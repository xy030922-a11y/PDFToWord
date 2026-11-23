#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Layout regression tests for PdfToWord.

Pipeline for every case:
    source PDF -> PdfToWord.exe (.docx) -> structural OOXML checks
               -> WPS exports the .docx back to PDF -> per-page pixel comparison

Run with the bundled Python environment (PyMuPDF, python-docx, numpy,
opencv-python-headless, pywin32):

    .conda\\python.exe tests\\layout_tests.py --exe build-vs\\Release\\PdfToWord.exe

Options:
    --no-wps     skip the WPS round-trip (structural checks only)
    --only NAME  run a single case, e.g. --only mixed_sizes
    --real-pdfs  verify the public PDFs downloaded into testPDF/ instead of
                 the synthetic cases (documents over 30 pages are skipped)
"""

from __future__ import annotations

import argparse
import pathlib
import re
import struct
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

import cv2
import numpy as np
import pymupdf

ROOT = pathlib.Path(__file__).resolve().parent.parent

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def q(prefix: str, tag: str) -> str:
    return "{%s}%s" % (NS[prefix], tag)


# --------------------------------------------------------------------------
# Test PDFs
# --------------------------------------------------------------------------

def text_page(doc, width=595, height=842, rotation=0, title="Page", zh=False):
    page = doc.new_page(width=width, height=height)
    if rotation:
        page.set_rotation(rotation)
    y = 70
    page.insert_text((50, y), "%s - PDF to Word layout test" % title,
                     fontsize=17, fontname="helv")
    y += 40
    for line in [
        "The quick brown fox jumps over the lazy dog 0123456789",
        "Second line with punctuation: !@#$%^&*() [] {} <>",
        "Third line, smaller type, still on the same page.",
    ]:
        page.insert_text((50, y), line, fontsize=11, fontname="helv")
        y += 26
    if zh:
        page.insert_text((50, y + 10), "中文测试：保留原版式", fontsize=14, fontname="china-s")
        y += 40
    # A table-like structure and a coloured block exercise vector graphics.
    for row in range(4):
        yy = y + row * 24
        page.draw_line((50, yy), (width - 50, yy), color=(0.2, 0.2, 0.2), width=0.6)
    page.draw_rect(pymupdf.Rect(50, y, 150, y + 72), color=(0, 0, 0), width=0.6)
    page.draw_rect(pymupdf.Rect(width - 160, height - 130, width - 50, height - 60),
                   color=None, fill=(0.85, 0.92, 0.98))
    return page


def image_page(doc, width=595, height=842, label="SCANNED PAGE"):
    img = np.full((1100, 780, 3), 255, np.uint8)
    cv2.rectangle(img, (30, 30), (750, 1070), (30, 30, 30), 5)
    cv2.putText(img, label, (70, 220), cv2.FONT_HERSHEY_SIMPLEX,
                2.0, (0, 0, 0), 4)
    cv2.putText(img, "NO TEXT LAYER", (70, 360), cv2.FONT_HERSHEY_SIMPLEX,
                1.7, (0, 0, 0), 3)
    for i in range(6):
        cv2.line(img, (70, 470 + i * 70), (700, 470 + i * 70), (90, 90, 90), 2)
    ok, buf = cv2.imencode(".png", img)
    assert ok
    page = doc.new_page(width=width, height=height)
    page.insert_image(page.rect, stream=buf.tobytes())
    return page


def build_cases() -> dict:
    def a4_text(doc):
        text_page(doc, title="A4 portrait", zh=True)

    def portrait_landscape(doc):
        text_page(doc, 595, 842, title="portrait 1")
        text_page(doc, 842, 595, title="landscape")
        text_page(doc, 595, 842, title="portrait 2")

    def mixed_sizes(doc):
        for i, (w, h) in enumerate([(595, 842), (420, 595), (612, 792), (500, 700)]):
            text_page(doc, w, h, title="size %d" % (i + 1))

    def oversize(doc):
        text_page(doc, 2400, 3200, title="oversized (33 x 44 inch)")
        text_page(doc, 595, 842, title="normal A4")

    def rotations(doc):
        text_page(doc, 595, 842, rotation=0, title="rotate 0")
        text_page(doc, 595, 842, rotation=90, title="rotate 90")
        text_page(doc, 595, 842, rotation=270, title="rotate 270")

    def many_pages(doc):
        for i in range(8):
            p = text_page(doc, title="page %d" % (i + 1))
            p.insert_text((50, 760), "footer %d" % (i + 1), fontsize=9, fontname="helv")

    def scanned(doc):
        image_page(doc)

    def mixed_content(doc):
        text_page(doc, title="text page 1", zh=True)
        image_page(doc, label="SCANNED MIDDLE")
        text_page(doc, title="text page 3")

    def tiny_blank(doc):
        text_page(doc, 100, 100, title="tiny")
        doc.new_page(width=595, height=842)  # completely blank

    return {
        "a4_text": a4_text,
        "portrait_landscape": portrait_landscape,
        "mixed_sizes": mixed_sizes,
        "oversize": oversize,
        "rotations": rotations,
        "many_pages": many_pages,
        "scanned": scanned,
        "mixed_content": mixed_content,
        "tiny_blank": tiny_blank,
    }


def make_pdfs(work: pathlib.Path) -> dict:
    out = {}
    for name, builder in build_cases().items():
        doc = pymupdf.open()
        builder(doc)
        path = work / (name + ".pdf")
        doc.save(path, garbage=4, deflate=True)
        doc.close()
        out[name] = path
    return out


# --------------------------------------------------------------------------
# Source PDF geometry
# --------------------------------------------------------------------------

MAX_PAGE_PT = 22 * 72          # Word caps a page at 22 inches
RENDER_DPI = 216
EMU_PER_TWIP = 635


def expected_twips(page) -> tuple[int, int]:
    """Mirror converter.cpp pageSizeInTwips(): pt -> DIPs -> capped -> twips."""
    w_pt, h_pt = page.rect.width, page.rect.height
    scale = min(1.0, MAX_PAGE_PT / max(w_pt, h_pt))
    return (int(round(w_pt * 20 * scale)), int(round(h_pt * 20 * scale)))


def expected_pixels(twips: int) -> int:
    return int(round(twips * RENDER_DPI / 1440))


# --------------------------------------------------------------------------
# DOCX parsing / structural checks
# --------------------------------------------------------------------------

def png_size(data: bytes) -> tuple[int, int]:
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValueError("not a PNG")
    return struct.unpack(">II", data[16:24])


def sect_info(sect) -> dict:
    if sect is None:
        return {}
    size = sect.find(q("w", "pgSz"))
    mar = sect.find(q("w", "pgMar"))
    info = {}
    if size is not None:
        info["w"] = int(size.get(q("w", "w"), "0"))
        info["h"] = int(size.get(q("w", "h"), "0"))
        info["orient"] = size.get(q("w", "orient"), "portrait")
    if mar is not None:
        info["mar"] = {k: int(mar.get(q("w", k), "-1"))
                       for k in ("top", "right", "bottom", "left", "header", "footer", "gutter")}
    return info


def parse_docx(path: pathlib.Path) -> dict:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if z.testzip() is not None:
            raise ValueError("corrupt zip entry")
        doc = ET.fromstring(z.read("word/document.xml"))
        rels = ET.fromstring(z.read("word/_rels/document.xml.rels")) \
            if "word/_rels/document.xml.rels" in names else None
        rid_target = {r.get("Id"): r.get("Target") for r in rels} if rels is not None else {}
        media = {n: z.read(n) for n in names if n.startswith("word/media/")}
        content_types = z.read("[Content_Types].xml").decode("utf-8", "replace")

    body = doc.find(q("w", "body"))
    paragraphs = []
    for p in body.findall(q("w", "p")):
        ppr = p.find(q("w", "pPr"))
        drawings = []
        for drawing in p.findall(".//" + q("w", "drawing")):
            anchor = drawing.find(q("wp", "anchor"))
            extent = anchor.find(q("wp", "extent")) if anchor is not None else None
            blip = drawing.find(".//" + q("a", "blip"))
            pos_h = anchor.find(q("wp", "positionH")) if anchor is not None else None
            pos_v = anchor.find(q("wp", "positionV")) if anchor is not None else None
            drawings.append({
                "cx": int(extent.get("cx", "0")) if extent is not None else 0,
                "cy": int(extent.get("cy", "0")) if extent is not None else 0,
                "rid": blip.get(q("r", "embed")) if blip is not None else None,
                "behindDoc": anchor.get("behindDoc") if anchor is not None else None,
                "allowOverlap": anchor.get("allowOverlap") if anchor is not None else None,
                "h_relative": pos_h.get("relativeFrom") if pos_h is not None else None,
                "v_relative": pos_v.get("relativeFrom") if pos_v is not None else None,
                "h_offset": pos_h.findtext(q("wp", "posOffset")) if pos_h is not None else None,
                "v_offset": pos_v.findtext(q("wp", "posOffset")) if pos_v is not None else None,
                "wrapNone": anchor.find(q("wp", "wrapNone")) is not None if anchor is not None else False,
            })
        paragraphs.append({
            "pageBreakBefore": ppr is not None and ppr.find(q("w", "pageBreakBefore")) is not None,
            "sect": sect_info(ppr.find(q("w", "sectPr"))) if ppr is not None else {},
            "drawings": drawings,
            "texts": [t.text or "" for t in p.findall(".//" + q("w", "t"))],
        })

    # pPr/sectPr terminates a section; the body-level sectPr describes the last one.
    groups, current = [], []
    for p in paragraphs:
        current.append(p)
        if p["sect"]:
            groups.append((current, p["sect"]))
            current = []
    if current:
        groups.append((current, sect_info(body.find(q("w", "sectPr")))))

    return {
        "names": names, "media": media, "rid_target": rid_target,
        "content_types": content_types, "paragraphs": paragraphs,
        "groups": groups,
        "fonts": sorted({int(s.get(q("w", "val"))) for s in doc.iter(q("w", "sz"))}),
    }


def check_layout_structure(docx: dict, source_pdf: pathlib.Path) -> list[str]:
    problems = []
    src = pymupdf.open(source_pdf)
    expected = [expected_twips(p) for p in src]
    src.close()
    total = len(expected)

    all_drawings = [d for p in docx["paragraphs"] for d in p["drawings"]]
    if len(docx["paragraphs"]) != total:
        problems.append("paragraph count %d != page count %d" % (len(docx["paragraphs"]), total))
    if len(all_drawings) != total:
        problems.append("full-page pictures %d != page count %d" % (len(all_drawings), total))
    if len(docx["media"]) != total:
        problems.append("media parts %d != page count %d" % (len(docx["media"]), total))

    if docx["groups"] and len(docx["groups"][-1][0]) == 0:
        problems.append("empty trailing section")

    page_no = 0
    for group_index, (paras, sect) in enumerate(docx["groups"]):
        if not sect:
            problems.append("section %d has no pgSz" % (group_index + 1))
            continue
        for p_index, p in enumerate(paras):
            if page_no >= total:
                break
            exp_w, exp_h = expected[page_no]
            if abs(sect.get("w", 0) - exp_w) > 1 or abs(sect.get("h", 0) - exp_h) > 1:
                problems.append("page %d pgSz %sx%s != expected %sx%s" %
                                (page_no + 1, sect.get("w"), sect.get("h"), exp_w, exp_h))
            orient = "landscape" if exp_w > exp_h else "portrait"
            if sect.get("orient") != orient:
                problems.append("page %d orientation %s != %s" %
                                (page_no + 1, sect.get("orient"), orient))
            mar = sect.get("mar", {})
            if any(mar.get(k, -1) != 0 for k in ("top", "right", "bottom", "left")):
                problems.append("page %d has non-zero margins %s" % (page_no + 1, mar))

            want_break = p_index > 0
            if p["pageBreakBefore"] != want_break:
                problems.append("page %d pageBreakBefore=%s, expected %s" %
                                (page_no + 1, p["pageBreakBefore"], want_break))

            if len(p["drawings"]) != 1:
                problems.append("page %d has %d pictures" % (page_no + 1, len(p["drawings"])))
                page_no += 1
                continue
            d = p["drawings"][0]
            if abs(d["cx"] - exp_w * EMU_PER_TWIP) > EMU_PER_TWIP or \
               abs(d["cy"] - exp_h * EMU_PER_TWIP) > EMU_PER_TWIP:
                problems.append("page %d picture extent %sx%s EMU does not match pgSz" %
                                (page_no + 1, d["cx"], d["cy"]))
            for key, want in [("h_relative", "page"), ("v_relative", "page"),
                              ("h_offset", "0"), ("v_offset", "0"),
                              ("behindDoc", "0"), ("allowOverlap", "1"), ("wrapNone", True)]:
                if d[key] != want:
                    problems.append("page %d anchor %s=%r, expected %r" %
                                    (page_no + 1, key, d[key], want))
            target = docx["rid_target"].get(d["rid"])
            if not target or ("word/" + target) not in docx["media"]:
                problems.append("page %d relationship %s -> %s is broken" %
                                (page_no + 1, d["rid"], target))
            else:
                data = docx["media"]["word/" + target]
                try:
                    px_w, px_h = png_size(data)
                    if abs(px_w - expected_pixels(exp_w)) > 1 or \
                       abs(px_h - expected_pixels(exp_h)) > 1:
                        problems.append("page %d PNG %dx%d != expected %dx%d" %
                                        (page_no + 1, px_w, px_h,
                                         expected_pixels(exp_w), expected_pixels(exp_h)))
                except ValueError as e:
                    problems.append("page %d PNG invalid: %s" % (page_no + 1, e))
            page_no += 1

    if total and docx["media"] and "image/png" not in docx["content_types"]:
        problems.append("image/png missing from [Content_Types].xml")
    if "wordprocessingml.document.main+xml" not in docx["content_types"]:
        problems.append("main document content type missing")
    return problems


def normalized_text(s: str) -> str:
    return re.sub(r"\s+", "", s)


def check_editable_structure(docx: dict, source_pdf: pathlib.Path) -> tuple[list[str], str]:
    problems = []
    src = pymupdf.open(source_pdf)
    expected = [expected_twips(p) for p in src]
    source_text = [p.get_text("text") for p in src]
    src.close()

    doc_text = "".join(t for p in docx["paragraphs"] for t in p["texts"])
    flat = normalized_text(doc_text).lower()
    tokens = set(re.findall(r"[A-Za-z0-9]{4,}", " ".join(source_text)))
    cjk = set(re.findall(r"[\u4e00-\u9fff]", " ".join(source_text)))
    wanted = {t.lower() for t in tokens} | cjk
    hit = sum(1 for t in wanted if t in flat)
    coverage = hit / len(wanted) if wanted else 1.0

    total = len(source_text)
    text_pages = [i for i, t in enumerate(source_text) if t.strip()]
    image_pages = [i for i, t in enumerate(source_text) if not t.strip()]
    drawings = [d for p in docx["paragraphs"] for d in p["drawings"]]
    # Only pages without a text layer may become pictures. This guards against
    # the regression where dense pages were all rasterised and became uneditable.
    if len(drawings) != len(image_pages):
        problems.append("picture pages %d != pages without a text layer %d" %
                        (len(drawings), len(image_pages)))
    all_images = len(drawings) == total
    if text_pages and all_images:
        problems.append("every page became a picture: the document is not editable")
    if text_pages and not all_images and wanted and coverage < 0.90:
        problems.append("text coverage %.1f%% (%d/%d source tokens missing)" %
                        (coverage * 100, len(wanted) - hit, len(wanted)))

    for i, sect in [(gi, s) for gi, (_, s) in enumerate(docx["groups"])]:
        if not sect:
            problems.append("editable section %d has no pgSz" % (i + 1))
            continue
        # Every section must describe one of the source pages' geometry.
        if not any(abs(sect.get("w", 0) - w) <= 1 and abs(sect.get("h", 0) - h) <= 1
                   for w, h in expected):
            problems.append("editable section %d pgSz %sx%s not in source sizes" %
                            (i + 1, sect.get("w"), sect.get("h")))
        mar = sect.get("mar", {})
        if any(mar.get(k, -1) < 0 or mar.get(k, 0) > 720
               for k in ("top", "right", "bottom", "left")):
            problems.append("editable section %d has unexpected margins %s" % (i + 1, mar))

    for d in drawings:
        if not any(abs(d["cx"] - w * EMU_PER_TWIP) <= EMU_PER_TWIP and
                   abs(d["cy"] - h * EMU_PER_TWIP) <= EMU_PER_TWIP
                   for w, h in expected):
            problems.append("editable picture extent %sx%s not in source sizes" % (d["cx"], d["cy"]))
        target = docx["rid_target"].get(d["rid"])
        if not target or ("word/" + target) not in docx["media"]:
            problems.append("editable picture relationship %s broken" % d["rid"])

    fonts = docx.get("fonts", [])
    font_desc = ("font=%.1f-%.1fpt" % (fonts[0] / 2, fonts[-1] / 2)) if fonts else "font=default"
    summary = "text pages=%d, picture pages=%d, %s, %s" % (
        len(text_pages), len(drawings), font_desc,
        "no editable text" if all_images else "token coverage=%.1f%%" % (coverage * 100))
    return problems, summary


# --------------------------------------------------------------------------
# WPS round-trip and visual comparison
# --------------------------------------------------------------------------

class WpsRenderer:
    def __init__(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        self.app = win32com.client.DispatchEx("KWps.Application")
        self.app.Visible = False
        try:
            self.app.DisplayAlerts = 0
        except Exception:
            pass

    def export(self, docx: pathlib.Path, pdf: pathlib.Path):
        if pdf.exists():
            pdf.unlink()
        doc = self.app.Documents.Open(str(docx), ReadOnly=True)
        try:
            doc.ExportAsFixedFormat(str(pdf), 17)  # 17 = wdExportFormatPDF
        finally:
            doc.Close(False)
        if not pdf.exists():
            raise RuntimeError("WPS did not produce %s" % pdf.name)

    def close(self):
        try:
            self.app.Quit()
        except Exception:
            pass
        try:
            self._pythoncom.CoUninitialize()
        except Exception:
            pass


def render_gray(page, width=700) -> np.ndarray:
    scale = width / page.rect.width
    pm = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale),
                         colorspace=pymupdf.csGRAY, alpha=False)
    return np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width)


def ink_iou(a: np.ndarray, b: np.ndarray, scale: float) -> float:
    if scale != 1.0:
        a = cv2.resize(a, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        b = cv2.resize(b, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    kernel = np.ones((3, 3), np.uint8)
    da = cv2.dilate((a < 210).astype(np.uint8), kernel)
    db = cv2.dilate((b < 210).astype(np.uint8), kernel)
    union = int(np.logical_or(da, db).sum())
    return float(np.logical_and(da, db).sum() / union) if union else 1.0


def check_image_geometry(page, index: int) -> list[str]:
    """The full-page picture must cover the whole page, unshifted and uncropped."""
    rects = [r for img in page.get_images(full=True) for r in page.get_image_rects(img[0])]
    if len(rects) != 1:
        return ["page %d has %d image placements, expected 1" % (index + 1, len(rects))]
    r = rects[0]
    if (abs(r.x0) > 2 or abs(r.y0) > 2 or
            abs(r.width - page.rect.width) > 2 or abs(r.height - page.rect.height) > 2):
        return ["page %d image rect (%.1f,%.1f %.1fx%.1f) does not cover page %.1fx%.1f" %
                (index + 1, r.x0, r.y0, r.width, r.height,
                 page.rect.width, page.rect.height)]
    return []


def visual_compare(source_pdf: pathlib.Path, rendered_pdf: pathlib.Path, *,
                   check_images=True, compare_pixels=True,
                   mae_limit=6.0) -> tuple[list[str], list[str]]:
    src = pymupdf.open(source_pdf)
    got = pymupdf.open(rendered_pdf)
    problems, stats = [], []
    if src.page_count != got.page_count:
        problems.append("WPS produced %d pages, source has %d" % (got.page_count, src.page_count))

    for i in range(min(src.page_count, got.page_count)):
        sw, sh = expected_twips(src[i])
        want_pt = (sw / 20.0, sh / 20.0)
        got_pt = (got[i].rect.width, got[i].rect.height)
        if abs(got_pt[0] - want_pt[0]) > 3 or abs(got_pt[1] - want_pt[1]) > 3:
            problems.append("page %d size %.1fx%.1f pt != expected %.1fx%.1f pt" %
                            (i + 1, got_pt[0], got_pt[1], want_pt[0], want_pt[1]))
        if check_images:
            problems += check_image_geometry(got[i], i)
        if not compare_pixels:
            continue

        a = render_gray(src[i])
        b = render_gray(got[i])
        if a.shape != b.shape:
            b = cv2.resize(b, (a.shape[1], a.shape[0]), interpolation=cv2.INTER_AREA)
        mae = float(np.abs(a.astype(np.float32) - b.astype(np.float32)).mean())
        try:
            (dx, dy), resp = cv2.phaseCorrelate(a.astype(np.float32), b.astype(np.float32))
        except cv2.error:
            dx = dy = 0.0
            resp = 1.0
        ious = [ink_iou(a, b, s) for s in (1.0, 0.5, 0.25)]
        iou = max(ious)  # multi-scale: tolerant to 1px anti-aliasing differences
        ink_a, ink_b = float((a < 210).mean()), float((b < 210).mean())
        blank = ink_a < 1e-4 and ink_b < 1e-4
        ink_delta = 0.0 if blank else abs(ink_a - ink_b) / max(ink_a, 1e-6)
        stats.append("p%d MAE=%4.2f IoU=%.3f shift=(%.2f,%.2f) resp=%.3f" %
                     (i + 1, mae, iou, dx, dy, resp))
        if blank:
            continue
        if mae > mae_limit:
            problems.append("page %d MAE %.2f > %.2f" % (i + 1, mae, mae_limit))
        if resp < 0.65:
            problems.append("page %d phase correlation %.3f < 0.65 (structure mismatch)" % (i + 1, resp))
        if abs(dx) > 2 or abs(dy) > 2:
            problems.append("page %d content shifted by (%.2f, %.2f) px" % (i + 1, dx, dy))
        if iou < 0.80:
            problems.append("page %d multi-scale ink IoU %.3f < 0.80" % (i + 1, iou))
        if ink_delta > 0.60:
            problems.append("page %d ink ratio differs by %.1f%%" % (i + 1, ink_delta * 100))
    src.close()
    got.close()
    return problems, stats


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------

def run_converter(exe: pathlib.Path, src: pathlib.Path, dst: pathlib.Path,
                  editable: bool) -> tuple[int, str]:
    if dst.exists():
        dst.unlink()
    args = [str(exe)]
    if editable:
        args.append("--editable")
    args += [str(src), str(dst)]
    proc = subprocess.run(args, capture_output=True, timeout=300)
    err = proc.stderr.decode("utf-8", "replace") + proc.stdout.decode("utf-8", "replace")
    return proc.returncode, err.strip()


def verify_real_pdfs(exe: pathlib.Path, work: pathlib.Path,
                     no_wps: bool, only: str) -> int:
    """Round-trip the public PDFs in testPDF/ (used with --real-pdfs)."""
    pdfdir = ROOT / "testPDF"
    files = sorted(pdfdir.glob("*.pdf"))
    if only:
        files = [f for f in files if only in f.name]
    if not files:
        print("no PDFs in", pdfdir, "- run tests/download_samples.py first")
        return 2
    work = work / "real-pdfs"
    work.mkdir(parents=True, exist_ok=True)
    print("real PDFs:", pdfdir)
    print("work    :", work)
    print("=" * 96)
    wps = None
    if not no_wps:
        try:
            wps = WpsRenderer()
        except Exception as e:
            print("WPS renderer unavailable (%s); visual checks skipped" % e)
    failures = []
    checked = 0
    try:
        for pdf in files:
            src = pymupdf.open(pdf)
            pages = src.page_count
            encrypted = src.needs_pass
            src.close()
            if encrypted:
                print("SKIP  %-38s (encrypted)" % pdf.name)
                continue
            if pages > 30:
                print("SKIP  %-38s %3d pages (>30, too slow for the round-trip)" %
                      (pdf.name, pages))
                continue
            checked += 1
            dst = work / (pdf.stem + "-layout.docx")
            rc, msg = run_converter(exe, pdf, dst, editable=False)
            if rc or not dst.is_file():
                print("FAIL  %-38s conversion rc=%d %s" % (pdf.name, rc, msg))
                failures.append((pdf.name, "conversion failed"))
                continue
            docx = parse_docx(dst)
            sproblems = check_layout_structure(docx, pdf)
            vproblems, stats = [], []
            if wps is not None:
                rendered = work / (pdf.stem + "-wps.pdf")
                try:
                    wps.export(dst, rendered)
                    # Real documents have denser text, so allow a slightly
                    # larger anti-aliasing MAE; geometry/IoU stay strict.
                    vproblems, stats = visual_compare(
                        pdf, rendered, check_images=True, compare_pixels=True,
                        mae_limit=10.0)
                except Exception as e:
                    vproblems = ["WPS export failed: %s" % e]
            mae = max((float(s.split("MAE=")[1].split()[0]) for s in stats), default=0.0)
            iou = min((float(s.split("IoU=")[1].split()[0]) for s in stats), default=1.0)

            # Editability regression guard: a PDF with a text layer must yield
            # real Word text, not rasterised full-page pictures.
            src = pymupdf.open(pdf)
            text_pages = sum(1 for p in src if p.get_text("text").strip())
            src.close()
            eproblems, edit_chars, edit_pics = [], 0, 0
            if text_pages:
                edst = work / (pdf.stem + "-editable.docx")
                erc, emsg = run_converter(exe, pdf, edst, editable=True)
                if erc or not edst.is_file():
                    eproblems.append("editable conversion failed: rc=%d %s" % (erc, emsg[:80]))
                else:
                    edocx = parse_docx(edst)
                    edit_chars = sum(len(t) for p in edocx["paragraphs"] for t in p["texts"])
                    edit_pics = sum(len(p["drawings"]) for p in edocx["paragraphs"])
                    if edit_chars == 0:
                        eproblems.append("editable output contains no text")
                    if edit_pics > pages - text_pages:
                        eproblems.append("%d pages rasterised (only %d have no text layer)" %
                                         (edit_pics, pages - text_pages))

            all_problems = sproblems + vproblems + eproblems
            status = "PASS" if not all_problems else "FAIL"
            edit_desc = ("%d chars, %d pics" % (edit_chars, edit_pics)
                         if text_pages else "n/a (no text layer)")
            print("%-4s  %-38s %3d pages | struct=%d visual=%d | MAE<=%.2f IoU>=%.3f | editable: %s" %
                  (status, pdf.name, pages, len(sproblems), len(vproblems), mae, iou,
                   edit_desc))
            for p in sproblems[:3]:
                print("       struct:", p)
            for p in vproblems[:3]:
                print("       visual:", p)
            for p in eproblems[:3]:
                print("       editable:", p)
            if all_problems:
                failures.append((pdf.name, len(all_problems)))
    finally:
        if wps is not None:
            wps.close()
    print("=" * 96)
    if failures:
        print("FAILED: %s" % failures)
        return 1
    print("ALL PASS: %d real PDFs (skipped %d large/encrypted)" %
          (checked, len(files) - checked))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe", default=str(ROOT / "build-vs" / "Release" / "PdfToWord.exe"))
    ap.add_argument("--work", default=str(ROOT / "build-check" / "layout-tests"))
    ap.add_argument("--no-wps", action="store_true")
    ap.add_argument("--only", default="")
    ap.add_argument("--real-pdfs", action="store_true")
    args = ap.parse_args()

    exe = pathlib.Path(args.exe).resolve()
    work = pathlib.Path(args.work).resolve()
    work.mkdir(parents=True, exist_ok=True)
    if not exe.is_file():
        print("executable not found:", exe)
        return 2
    if not (exe.parent / "pdfium.dll").is_file():
        print("pdfium.dll not found next to", exe)
        return 2

    if args.real_pdfs:
        return verify_real_pdfs(exe, work, args.no_wps, args.only)

    pdfs = make_pdfs(work)
    if args.only:
        pdfs = {k: v for k, v in pdfs.items() if k == args.only}
        if not pdfs:
            print("unknown case:", args.only)
            return 2

    print("exe :", exe)
    print("work:", work)
    print("cases:", ", ".join(pdfs))
    print("=" * 78)

    wps = None
    if not args.no_wps:
        try:
            wps = WpsRenderer()
            print("WPS renderer: ready")
        except Exception as e:
            print("WPS renderer unavailable (%s); visual checks skipped" % e)
            wps = None

    failures = []
    try:
        for name, pdf in pdfs.items():
            print("\n[%s]  %d source pages" % (name, pymupdf.open(pdf).page_count))
            layout_docx = work / (name + "-layout.docx")
            rc, msg = run_converter(exe, pdf, layout_docx, editable=False)
            if rc != 0 or not layout_docx.is_file():
                failures.append((name, "layout conversion failed: rc=%d %s" % (rc, msg)))
                print("  layout conversion FAILED:", msg)
                continue
            docx = parse_docx(layout_docx)
            problems = check_layout_structure(docx, pdf)
            print("  OOXML structure: %s (%d sections, %d pictures)" %
                  ("PASS" if not problems else "FAIL",
                   len(docx["groups"]), len(docx["media"])))
            for p in problems:
                print("     -", p)
            failures += [(name, "layout structure: " + p) for p in problems]

            if wps is not None:
                rendered = work / (name + "-wps.pdf")
                try:
                    wps.export(layout_docx, rendered)
                    vproblems, stats = visual_compare(pdf, rendered)
                    print("  WPS round-trip: %s | %s" %
                          ("PASS" if not vproblems else "FAIL", "  ".join(stats)))
                    for p in vproblems:
                        print("     -", p)
                    failures += [(name, "WPS visual: " + p) for p in vproblems]
                except Exception as e:
                    failures.append((name, "WPS export failed: %s" % e))
                    print("  WPS round-trip FAILED:", e)

            editable_docx = work / (name + "-editable.docx")
            src = pymupdf.open(pdf)
            has_text = any(p.get_text("text").strip() for p in src)
            src.close()
            rc, msg = run_converter(exe, pdf, editable_docx, editable=True)
            if has_text:
                if rc != 0 or not editable_docx.is_file():
                    failures.append((name, "editable conversion failed: rc=%d %s" % (rc, msg)))
                    print("  editable conversion FAILED:", msg)
                else:
                    edocx = parse_docx(editable_docx)
                    eproblems, summary = check_editable_structure(edocx, pdf)
                    print("  editable mode: %s (%s)" % ("PASS" if not eproblems else "FAIL", summary))
                    for p in eproblems:
                        print("     -", p)
                    failures += [(name, "editable: " + p) for p in eproblems]
                    if wps is not None:
                        erendered = work / (name + "-editable-wps.pdf")
                        try:
                            wps.export(editable_docx, erendered)
                            gproblems, _ = visual_compare(
                                pdf, erendered, check_images=False, compare_pixels=False)
                            print("  editable pagination: %s" % ("PASS" if not gproblems else "FAIL"))
                            for p in gproblems:
                                print("     -", p)
                            failures += [(name, "editable pagination: " + p) for p in gproblems]
                        except Exception as e:
                            failures.append((name, "editable WPS export failed: %s" % e))
                            print("  editable pagination FAILED:", e)
            else:
                if rc == 0 or "OCR" not in msg:
                    failures.append((name, "scanned PDF should fail with an OCR hint, got rc=%d %s" % (rc, msg)))
                    print("  scanned-PDF handling FAILED:", msg)
                else:
                    print("  scanned-PDF handling: PASS (clear OCR hint)")

            # Independent opener: python-docx must accept the package.
            try:
                from docx import Document
                for f in (layout_docx, editable_docx):
                    if f.is_file():
                        Document(str(f))
                print("  python-docx open: PASS")
            except Exception as e:
                failures.append((name, "python-docx failed: %s" % e))
                print("  python-docx open FAILED:", e)
    finally:
        if wps is not None:
            wps.close()

    print("\n" + "=" * 78)
    if failures:
        print("FAILED: %d problem(s)" % len(failures))
        for name, reason in failures:
            print("  [%s] %s" % (name, reason))
        return 1
    print("ALL PASS: %d cases" % len(pdfs))
    return 0


if __name__ == "__main__":
    sys.exit(main())
