# PDF 转 Word

Windows 桌面应用，使用 C++、Qt 6 Widgets 和 CMake。

- **保留原版式**是默认模式。每个 PDF 页面以整页图片插入 Word，保留分页、方向和版面，支持扫描版 PDF。超出 Word 尺寸上限的页面会等比例缩小。
- **可编辑文本**由桌面版 Microsoft Word 2013 或更新版本转换，分页和排版可能变化。扫描版如需修改文字，请先做 OCR。

## 编译和运行

需要 MSVC、Windows 10 或 11 SDK（含 C++/WinRT）、Qt 6.5 以上的 MSVC 版本（含 Widgets 和 Concurrent）、CMake 3.21 以上和 Ninja。

在 Visual Studio 的 x64 开发者命令提示符中，从仓库目录运行。将以下 Qt 路径替换为本机的安装路径。

```bat
set "PATH=C:\Qt\6.12.0\msvc2022_64\bin;%PATH%"
cmake -S . -B build -G Ninja -DCMAKE_BUILD_TYPE=Release -DCMAKE_PREFIX_PATH=C:/Qt/6.12.0/msvc2022_64
cmake --build build
windeployqt --release build\PdfToWord.exe
build\PdfToWord.exe
```

也可以在 Qt Creator 中打开 `CMakeLists.txt`，选择对应的 MSVC Qt Kit 后构建和运行。
