import argparse
import os
import queue
import sys
import tempfile
import threading
import tkinter as tk
import warnings
from io import BytesIO
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

WD_FORMAT_DOCX = 16
WD_AUTOMATION_SECURITY_FORCE_DISABLE = 3
WORD_MAX_PAGE_POINTS = 22 * 72
PAGE_IMAGE_DPI = 300
CONVERSION_MODES = {
    "保留原版式": "layout",
    "可编辑文本": "auto",
}


class ConversionError(Exception):
    pass


def _validate_pdf(source):
    try:
        import pymupdf
    except ImportError as exc:
        raise ConversionError("缺少 PDF 读取依赖，请双击“启动.bat”完成安装。") from exc

    try:
        with pymupdf.open(source) as document:
            if document.needs_pass:
                raise ConversionError("暂不支持带密码的 PDF，请先解除密码保护。")
            if document.page_count == 0:
                raise ConversionError("这个 PDF 没有可转换的页面。")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"无法读取 PDF：{exc}") from exc


def _anchor_to_page(picture, page_number):
    from docx.oxml import OxmlElement

    # Use the existing picture payload, but place it outside Word's text flow.
    inline = picture._inline
    anchor = OxmlElement("wp:anchor")
    for name, value in {
        "simplePos": "0",
        "relativeHeight": "0",
        "behindDoc": "0",
        "locked": "1",
        "layoutInCell": "1",
        "allowOverlap": "1",
    }.items():
        anchor.set(name, value)

    origin = OxmlElement("wp:simplePos")
    origin.set("x", "0")
    origin.set("y", "0")
    anchor.append(origin)
    for axis in ("H", "V"):
        position = OxmlElement(f"wp:position{axis}")
        position.set("relativeFrom", "page")
        offset = OxmlElement("wp:posOffset")
        offset.text = "0"
        position.append(offset)
        anchor.append(position)

    anchor.append(inline.extent)
    effect_extent = OxmlElement("wp:effectExtent")
    for edge in ("l", "t", "r", "b"):
        effect_extent.set(edge, "0")
    anchor.append(effect_extent)
    anchor.append(OxmlElement("wp:wrapNone"))
    inline.docPr.set("descr", f"PDF 第 {page_number} 页")
    for child in list(inline):
        anchor.append(child)
    inline.getparent().replace(inline, anchor)


def _convert_preserving_layout(source, destination, progress):
    try:
        import pymupdf
        from docx import Document
        from docx.enum.section import WD_ORIENT, WD_SECTION_START
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt, Twips
    except ImportError as exc:
        raise ConversionError("缺少版式转换依赖，请双击“启动.bat”完成安装。") from exc

    document = Document()
    settings = document.settings.element
    view = OxmlElement("w:view")
    view.set(qn("w:val"), "print")
    settings.insert_element_before(view, "w:zoom")
    settings.insert_element_before(
        OxmlElement("w:doNotAutoCompressPictures"), "w:shapeDefaults", "w:decimalSymbol"
    )

    # The paragraphs hold the anchors and section breaks without taking up a full line.
    paragraph_format = document.styles["Normal"].paragraph_format
    paragraph_format.space_before = paragraph_format.space_after = Pt(0)
    paragraph_format.line_spacing = Pt(1)

    section = document.sections[0]
    section.top_margin = section.bottom_margin = Pt(0)
    section.left_margin = section.right_margin = Pt(0)
    section.header_distance = section.footer_distance = Pt(0)

    previous_size = None
    with pymupdf.open(source) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            rect = page.rect
            scale = min(1, WORD_MAX_PAGE_POINTS / max(rect.width, rect.height))
            width = Twips(max(1, round(rect.width * scale * 20)))
            height = Twips(max(1, round(rect.height * scale * 20)))
            size = (width, height)

            if size != previous_size:
                if previous_size is not None:
                    section = document.add_section(WD_SECTION_START.NEW_PAGE)
                section.orientation = WD_ORIENT.LANDSCAPE if width > height else WD_ORIENT.PORTRAIT
                section.page_width, section.page_height = size

            paragraph = document.add_paragraph()
            paragraph.paragraph_format.page_break_before = size == previous_size
            previous_size = size

            zoom = scale * PAGE_IMAGE_DPI / 72
            image = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            image.set_dpi(PAGE_IMAGE_DPI, PAGE_IMAGE_DPI)
            image_stream = BytesIO(image.tobytes("png"))
            del image
            picture = paragraph.add_run().add_picture(image_stream, width, height)
            _anchor_to_page(picture, page_number)
            if progress is not None:
                progress(page_number, pdf.page_count)

    document.save(destination)


def _convert_with_word(source, destination):
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise ConversionError("未安装 Microsoft Word 自动化组件 pywin32。") from exc

    initialized = False
    word = None
    document = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = WD_AUTOMATION_SECURITY_FORCE_DISABLE
        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        document.SaveAs2(str(destination), WD_FORMAT_DOCX)
        document.Close(False)
        document = None
    except Exception as exc:
        raise ConversionError(f"Microsoft Word 转换失败：{exc}") from exc
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        if initialized:
            pythoncom.CoUninitialize()


def _convert_with_builtin_engine(source, destination):
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", FutureWarning)
            from pdf2docx import Converter
    except ImportError as exc:
        raise ConversionError("缺少内置转换依赖，请双击“启动.bat”完成安装。") from exc

    converter = None
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", FutureWarning)
            converter = Converter(str(source))
            converter.convert(str(destination))
        converter.close()
        converter = None
    except Exception as exc:
        raise ConversionError(f"内置转换引擎失败：{exc}") from exc
    finally:
        if converter is not None:
            converter.close()


def convert_pdf(pdf_path, output_path, engine="layout", *, progress=None):
    source = Path(pdf_path).expanduser().resolve()
    destination = Path(output_path).expanduser().resolve()

    if source.suffix.lower() != ".pdf":
        raise ConversionError("请选择 PDF 文件。")
    if not source.is_file():
        raise ConversionError(f"PDF 文件不存在：{source}")
    if destination.suffix.lower() != ".docx":
        raise ConversionError("输出文件必须使用 .docx 扩展名。")
    if not destination.parent.is_dir():
        raise ConversionError(f"输出文件夹不存在：{destination.parent}")
    if engine not in {"layout", "auto", "word", "built-in"}:
        raise ConversionError("转换引擎只能是 layout、auto、word 或 built-in。")

    _validate_pdf(source)

    temp_file = tempfile.NamedTemporaryFile(
        prefix=f".{destination.stem}.",
        suffix=".docx",
        dir=destination.parent,
        delete=False,
    )
    temp_path = Path(temp_file.name)
    temp_file.close()

    try:
        if engine == "layout":
            _convert_preserving_layout(source, temp_path, progress)
            used_engine = "保留原版式"
        elif engine == "word":
            _convert_with_word(source, temp_path)
            used_engine = "Microsoft Word"
        elif engine == "built-in":
            _convert_with_builtin_engine(source, temp_path)
            used_engine = "内置引擎"
        else:
            try:
                _convert_with_word(source, temp_path)
                used_engine = "Microsoft Word"
            except ConversionError as word_error:
                temp_path.unlink(missing_ok=True)
                try:
                    _convert_with_builtin_engine(source, temp_path)
                    used_engine = "内置引擎"
                except ConversionError as builtin_error:
                    raise ConversionError(
                        f"{word_error}\n内置引擎也失败：{builtin_error}"
                    ) from builtin_error

        if not temp_path.is_file() or temp_path.stat().st_size == 0:
            raise ConversionError(f"{used_engine}没有生成有效的 Word 文件。")

        os.replace(temp_path, destination)
    except Exception as exc:
        if isinstance(exc, ConversionError):
            raise
        raise ConversionError(f"转换失败：{exc}") from exc
    finally:
        temp_path.unlink(missing_ok=True)

    return destination, used_engine


class PdfToWordApp:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.worker = None
        self.pdf_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择一个 PDF 文件")

        root.title("PDF 转 Word")
        root.geometry("680x405")
        root.minsize(620, 405)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Body.TLabel", font=("Microsoft YaHei UI", 10))
        style.configure("Action.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(18, 9))

        main = ttk.Frame(root, padding=26)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(1, weight=1)

        ttk.Label(main, text="PDF 转 Word", style="Title.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 22)
        )

        ttk.Label(main, text="PDF 文件", style="Body.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 12))
        self.pdf_entry = ttk.Entry(main, textvariable=self.pdf_var)
        self.pdf_entry.grid(row=1, column=1, sticky="ew", ipady=4)
        self.pdf_button = ttk.Button(main, text="浏览…", command=self.browse_pdf)
        self.pdf_button.grid(row=1, column=2, padx=(10, 0))

        ttk.Label(main, text="保存为", style="Body.TLabel").grid(
            row=2, column=0, sticky="w", padx=(0, 12), pady=(16, 0)
        )
        self.output_entry = ttk.Entry(main, textvariable=self.output_var)
        self.output_entry.grid(row=2, column=1, sticky="ew", ipady=4, pady=(16, 0))
        self.output_button = ttk.Button(main, text="浏览…", command=self.browse_output)
        self.output_button.grid(row=2, column=2, padx=(10, 0), pady=(16, 0))

        ttk.Label(main, text="转换模式", style="Body.TLabel").grid(
            row=3, column=0, sticky="w", padx=(0, 12), pady=(16, 0)
        )
        self.mode_box = ttk.Combobox(main, values=tuple(CONVERSION_MODES), state="readonly")
        self.mode_box.current(0)
        self.mode_box.grid(row=3, column=1, columnspan=2, sticky="ew", ipady=4, pady=(16, 0))
        ttk.Label(
            main,
            text="保留原版式适用于文字版和扫描版 PDF。需要修改文字时，请选“可编辑文本”。",
            foreground="#666666",
            wraplength=550,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(22, 10))
        ttk.Label(main, textvariable=self.status_var, style="Body.TLabel").grid(
            row=6, column=0, columnspan=3, sticky="w"
        )

        self.convert_button = ttk.Button(
            main, text="开始转换", style="Action.TButton", command=self.start_conversion
        )
        self.convert_button.grid(row=7, column=0, columnspan=3, pady=(22, 0))

        root.after(100, self.poll_events)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        )
        if filename:
            source = Path(filename)
            self.pdf_var.set(str(source))
            self.output_var.set(str(source.with_suffix(".docx")))
            self.status_var.set(f"已选择：{source.name}")

    def browse_output(self):
        current = Path(self.output_var.get()) if self.output_var.get().strip() else None
        source = Path(self.pdf_var.get()) if self.pdf_var.get().strip() else None
        initial_file = current.name if current else (source.with_suffix(".docx").name if source else "转换结果.docx")
        initial_dir = str(current.parent) if current and current.parent.is_dir() else None
        filename = filedialog.asksaveasfilename(
            title="保存 Word 文件",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialfile=initial_file,
            initialdir=initial_dir,
            confirmoverwrite=True,
        )
        if filename:
            self.output_var.set(filename)

    def start_conversion(self):
        source = self.pdf_var.get().strip()
        destination = self.output_var.get().strip()
        if not source:
            messagebox.showwarning("缺少 PDF", "请先选择要转换的 PDF 文件。")
            return
        if not destination:
            messagebox.showwarning("缺少保存位置", "请选择 Word 文件的保存位置。")
            return
        if Path(destination).exists() and not messagebox.askyesno(
            "覆盖文件", "目标 Word 文件已存在，转换完成后要覆盖它吗？"
        ):
            return

        self.set_busy(True)
        engine = CONVERSION_MODES[self.mode_box.get()]
        self.status_var.set("正在转换，大型 PDF 可能需要几分钟…")
        self.worker = threading.Thread(
            target=self.convert_in_background,
            args=(source, destination, engine),
            daemon=True,
        )
        self.worker.start()

    def convert_in_background(self, source, destination, engine):
        try:
            result = convert_pdf(source, destination, engine)
            self.events.put(("success", result))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def poll_events(self):
        try:
            while True:
                event, payload = self.events.get_nowait()
                self.set_busy(False)
                if event == "success":
                    result_path, engine_name = payload
                    self.status_var.set(f"转换完成（{engine_name}）：{result_path}")
                    if messagebox.askyesno(
                        "转换完成",
                        f"Word 文件已保存到：\n{result_path}\n\n转换引擎：{engine_name}\n是否现在打开？",
                    ):
                        try:
                            os.startfile(result_path)
                        except OSError as exc:
                            messagebox.showerror("无法打开文件", str(exc))
                else:
                    self.status_var.set("转换失败，请检查文件后重试。")
                    messagebox.showerror("转换失败", payload)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_events)

    def set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for widget in (self.pdf_entry, self.output_entry, self.pdf_button, self.output_button, self.convert_button):
            widget.configure(state=state)
        self.mode_box.configure(state="disabled" if busy else "readonly")
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def on_close(self):
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("正在转换", "请等待当前转换完成后再关闭窗口。")
            return
        self.root.destroy()


def main():
    parser = argparse.ArgumentParser(description="将 PDF 转换为尽量保留原始版面的 Word DOCX 文档。")
    parser.add_argument("pdf", nargs="?", help="要转换的 PDF 路径；不提供时启动图形界面")
    parser.add_argument("output", nargs="?", help="输出 DOCX 路径，默认与 PDF 同目录同名")
    parser.add_argument(
        "--engine",
        choices=("layout", "auto", "word", "built-in"),
        default="layout",
        help="默认 layout，整页图片保留版式。可编辑模式可选 auto、word 或 built-in。",
    )
    args = parser.parse_args()

    if not args.pdf:
        root = tk.Tk()
        PdfToWordApp(root)
        root.mainloop()
        return 0

    source = Path(args.pdf).expanduser().resolve()
    destination = Path(args.output).expanduser().resolve() if args.output else source.with_suffix(".docx")
    try:
        result_path, engine_name = convert_pdf(source, destination, args.engine)
    except ConversionError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1

    print(f"转换完成（{engine_name}）：{result_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
