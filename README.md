# PDF 转 Word

Windows 桌面应用，使用 C++、Qt 6 Widgets 和 CMake。**转换过程完全不需要 Microsoft Word**（也不调用 WPS、LibreOffice 等任何 Office 程序），生成的是标准 .docx 文件，可用任意支持 DOCX 的办公软件打开。

- **保留原版式**（默认）：用 Windows 自带的 PDF 渲染接口把每页渲染成整页图片，直接写入 DOCX。保留分页、方向和版面，支持扫描版 PDF；超出 Word 页面尺寸上限的页面会等比例缩小。
- **可编辑文本**：用 PDFium 直接提取 PDF 文字层，以**真正的 Word 文字**写入 DOCX，无需安装 Microsoft Word。字号会按页面可用空间自动缩小（11 pt 到最小 4 pt），避免文字溢出到额外页，保证「一页 PDF 对应一页 Word」；只有没有文字层的扫描页才会以整页图片保留。整份 PDF 都没有文字时会提示先做 OCR 或改用「保留原版式」。

## 依赖

| 依赖 | 用途 |
| --- | --- |
| MSVC、Windows 10/11 SDK（含 C++/WinRT）、CMake 3.21+、Ninja | 编译与 PDF 页面渲染 |
| Qt 6.5+ 的 MSVC 版本（Widgets、Concurrent） | 界面 |
| PDFium 预编译包（Windows x64） | 提取 PDF 文字，替代原来的 Word COM 自动化 |

PDFium 放在 `third_party/pdfium-win-x64/`，目录内应有 `PDFiumConfig.cmake`、`include/`、`lib/pdfium.dll.lib`、`bin/pdfium.dll`。当前使用的是 [pdfium-binaries](https://github.com/bblanchon/pdfium-binaries) 的 `chromium/8165`（`pdfium-win-x64.tgz`），许可证见包内 `LICENSE`。如果放在别处，配置时传入 `-DPDFIUM_PACKAGE_DIR=<路径>`。

`pdfium.dll` 采用延迟加载并会在构建后自动复制到 exe 同目录：缺少它时「保留原版式」仍可正常使用，「可编辑文本」会给出明确提示，程序不会启动失败。发布时把 exe、`pdfium.dll` 和 `windeployqt` 生成的 Qt DLL 放在一起即可。

## 编译和运行

在 Visual Studio 的 x64 开发者命令提示符中，从仓库目录运行。将以下 Qt 路径替换为本机的安装路径。

```bat
set "PATH=C:\Qt\6.12.0\msvc2022_64\bin;%PATH%"
cmake -S . -B build -G Ninja -DCMAKE_BUILD_TYPE=Release -DCMAKE_PREFIX_PATH=C:/Qt/6.12.0/msvc2022_64
cmake --build build
windeployqt --release build\PdfToWord.exe
build\PdfToWord.exe
```

也可以在 Qt Creator 中打开 `CMakeLists.txt`，选择对应的 MSVC Qt Kit 后构建和运行。

## 命令行

```bat
PdfToWord.exe 输入.pdf 输出.docx
PdfToWord.exe --editable 输入.pdf 输出.docx
```

不带参数启动图形界面；`--editable` 表示「可编辑文本」模式。无参数时输出文件默认保存在 PDF 所在文件夹。

## 布局一致性测试

`tests/layout_tests.py` 自动生成 9 组版式各异的 PDF：A4 单页、横竖混排、4 种页面尺寸、超过 22 英寸的超大页、旋转 0°/90°/270°、8 页长文档、纯扫描页、图文混排、极小页 + 空白页。每种都做三层校验：

1. **OOXML 结构**：分节数、`pgSz`/方向、零页边距、图片 extent 与锚点、分页符、PNG 像素尺寸、关系与内容类型；可编辑模式还会检查文字覆盖率，并断言「图片页数 == 无文字层页数」、有文字层的文档不能整页栅格化（防止再次出现生成结果无法编辑的回归）。
2. **WPS 渲染往返**：把生成的 DOCX 用 WPS 导出回 PDF，检查页数、每页尺寸，以及整页图片的摆放矩形必须精确覆盖页面（不裁切、不偏移、不缩放）。
3. **逐页像素对比**：MAE、相位相关（位移与结构响应）、多尺度墨迹 IoU、前景比例，容忍抗锯齿而不放过真实错位。

```bat
.conda\python.exe tests\layout_tests.py --exe build-vs\Release\PdfToWord.exe
```

需要项目自带的 `.conda` Python 环境（PyMuPDF、python-docx、numpy、OpenCV、pywin32）和本机 WPS 作为独立渲染器；`--no-wps` 可跳过渲染往返，`--only <用例名>` 可单独跑一个用例。任一检查失败时脚本以非零退出码结束。

真实 PDF 样本放在 `testPDF/`，由 `tests/download_samples.py` 下载 16 个公开文件（arXiv 论文、RFC 2616、NIST SP 800-63-3、W3C、Unicode CJK 汉字码表，以及 pdf.js 的旋转页/空白页/传真扫描/透明图形/表单等版式边界样例）。用 `--real-pdfs` 对它们跑同一套往返校验，超过 30 页的大文档自动跳过，MAE 阈值放宽到 10 以容纳密集文字的抗锯齿差异，页面几何与 IoU 仍然严格：

```bat
.conda\python.exe tests\download_samples.py
.conda\python.exe tests\layout_tests.py --real-pdfs
```
