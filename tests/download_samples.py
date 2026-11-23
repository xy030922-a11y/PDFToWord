#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download a varied set of public PDFs into testPDF/ for converter testing.

Run with the bundled Python environment (only PyMuPDF is needed for the report):

    .conda\\python.exe tests\\download_samples.py

All files are public samples: academic preprints, RFC / NIST / W3C documents,
the Unicode CJK code chart, and small layout edge cases from the pdf.js test
suite (fetched through the GitHub contents API because raw.githubusercontent.com
is not reachable from this machine).
"""

from __future__ import annotations

import base64
import json
import pathlib
import subprocess
import sys
import urllib.parse

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "testPDF"
TMP = ROOT / "build-check" / "download-tmp"

# (local filename, URL)
DIRECT = [
    ("01_tracemonkey_pldi09.pdf",
     "https://mozilla.github.io/pdf.js/web/compressed.tracemonkey-pldi-09.pdf"),
    ("02_attention_is_all_you_need.pdf",
     "https://arxiv.org/pdf/1706.03762"),
    ("03_rfc2616_http11.pdf",
     "https://www.rfc-editor.org/rfc/rfc2616.pdf"),
    ("04_nist_sp800_63_3.pdf",
     "https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-63-3.pdf"),
    ("05_w3c_dummy.pdf",
     "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf"),
    ("06_gpt4_technical_report.pdf",
     "https://arxiv.org/pdf/2303.08774"),
    ("07_unicode_cjk_codechart.pdf",
     "https://www.unicode.org/charts/PDF/U4E00.pdf"),
]

# (local filename, file in mozilla/pdf.js test/pdfs)
GITHUB = [
    ("08_pdfjs_hello_world_rotated.pdf", "hello_world_rotated.pdf"),
    ("09_pdfjs_empty.pdf", "empty.pdf"),
    ("10_pdfjs_canvas.pdf", "canvas.pdf"),
    ("11_pdfjs_ccitt_fax_scan.pdf", "ccitt_EndOfBlock_false.pdf"),
    ("12_pdfjs_images.pdf", "images.pdf"),
    ("13_pdfjs_bigboundingbox.pdf", "bigboundingbox.pdf"),
    ("14_pdfjs_alphatrans.pdf", "alphatrans.pdf"),
    ("15_pdfjs_form_two_pages.pdf", "form_two_pages.pdf"),
    ("16_pdfjs_image_rotated.pdf", "image-rotated-black-white-ratio.pdf"),
]

API = "https://api.github.com/repos/mozilla/pdf.js"


def curl(url: str, dest: pathlib.Path, timeout: int = 300):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    proc = subprocess.run(
        ["curl", "-sL", "--fail", "--max-time", str(timeout), "-o", str(dest), url],
        capture_output=True, timeout=timeout + 10)
    if proc.returncode or not dest.exists() or dest.stat().st_size == 0:
        raise RuntimeError("curl failed (%d) for %s: %s" %
                           (proc.returncode, url, proc.stderr.decode("utf-8", "replace")[:150]))
    return dest.read_bytes()


def github_pdf(remote_name: str) -> bytes:
    quoted = urllib.parse.quote(remote_name)
    meta = json.loads(curl(API + "/contents/test/pdfs/" + quoted,
                           TMP / "meta.json", 90).decode("utf-8"))
    if meta.get("content"):
        return base64.b64decode("".join(meta["content"].split()))
    # Files above ~1 MB have no inline content; use the git blob endpoint.
    blob = json.loads(curl(API + "/git/blobs/" + meta["sha"],
                           TMP / "blob.json", 120).decode("utf-8"))
    return base64.b64decode("".join(blob["content"].split()))


def inspect(path: pathlib.Path) -> str:
    import pymupdf
    doc = pymupdf.open(path)
    if doc.needs_pass:
        return "ENCRYPTED (skipped by converter)"
    pages = doc.page_count
    sizes = sorted({(round(p.rect.width), round(p.rect.height)) for p in doc})
    rotations = sorted({p.rotation for p in doc})
    chars = sum(len(p.get_text("text")) for p in doc)
    images = sum(len(p.get_images()) for p in doc)
    doc.close()
    size_desc = ", ".join("%dx%d" % s for s in sizes[:3])
    if len(sizes) > 3:
        size_desc += " (+%d more)" % (len(sizes) - 3)
    return ("%3d pages | %-22s | rot=%s | text=%d chars | images=%d" %
            (pages, size_desc, ",".join(map(str, rotations)), chars, images))


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    TMP.mkdir(parents=True, exist_ok=True)
    print("Destination:", OUT)
    print("=" * 100)
    ok, failed = 0, []
    for filename, url in DIRECT + [(n, None) for n, _ in GITHUB]:
        dest = OUT / filename
        try:
            if dest.exists() and dest.read_bytes()[:5] == b"%PDF-":
                data = dest.read_bytes()
                source = "cached"
            else:
                if url is None:
                    remote = next(r for n, r in GITHUB if n == filename)
                    data = github_pdf(remote)
                    source = "pdf.js test suite (GitHub API)"
                else:
                    data = curl(url, TMP / "download.bin")
                    source = url.split("/")[2]
                if data[:5] != b"%PDF-":
                    raise RuntimeError("not a PDF (starts with %r)" % data[:16])
                dest.write_bytes(data)
            report = inspect(dest)
            print("OK   %-38s %8.1f KB  %s" % (filename, len(data) / 1024, report))
            print("     source: %s" % source)
            ok += 1
        except Exception as e:
            failed.append((filename, str(e)))
            print("FAIL %-38s %s" % (filename, e))
            if dest.exists() and dest.read_bytes()[:5] != b"%PDF-":
                dest.unlink()
    print("=" * 100)
    print("downloaded/verified: %d, failed: %d" % (ok, len(failed)))
    for name, reason in failed:
        print("  %s: %s" % (name, reason))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
