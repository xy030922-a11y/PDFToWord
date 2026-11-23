#include "converter.h"

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QMutex>
#include <QRegularExpression>
#include <QScopeGuard>
#include <QSize>
#include <QStringList>
#include <QTemporaryDir>
#include <QXmlStreamWriter>

#include <fpdf_text.h>
#include <fpdfview.h>

#include <windows.h>
#include <msopc.h>
#include <shcore.h>
#include <winrt/Windows.Data.Pdf.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Windows.Storage.h>
#include <winrt/Windows.Storage.Streams.h>
#include <winrt/Windows.UI.h>

#include <algorithm>
#include <cmath>
#include <limits>
#include <vector>

namespace {

using winrt::check_hresult;
using winrt::com_ptr;
using winrt::Windows::Data::Pdf::PdfDocument;
using winrt::Windows::Data::Pdf::PdfPage;
using winrt::Windows::Data::Pdf::PdfPageRenderOptions;
using winrt::Windows::Graphics::Imaging::BitmapEncoder;
using winrt::Windows::Storage::FileAccessMode;
using winrt::Windows::Storage::Streams::DataReader;
using winrt::Windows::Storage::Streams::InMemoryRandomAccessStream;
using winrt::Windows::Storage::Streams::IRandomAccessStream;

constexpr double dipsPerInch = 96;
constexpr double twipsPerInch = 1440;
constexpr int emusPerTwip = 635;
constexpr int renderDpi = 216;
constexpr double maxWordPageDips = 22 * dipsPerInch;

constexpr auto officeRelationship =
    L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr auto imageRelationship =
    L"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

class Apartment
{
public:
    explicit Apartment(winrt::apartment_type type) { winrt::init_apartment(type); }
    ~Apartment() { winrt::uninit_apartment(); }
    Apartment(const Apartment &) = delete;
    Apartment &operator=(const Apartment &) = delete;
};

class DocxPackage
{
public:
    DocxPackage()
    {
        check_hresult(CoCreateInstance(__uuidof(OpcFactory), nullptr, CLSCTX_INPROC_SERVER,
                                       __uuidof(IOpcFactory), factory.put_void()));
        check_hresult(factory->CreatePackage(package.put()));
        check_hresult(package->GetPartSet(parts.put()));
    }

    com_ptr<IOpcPart> addPart(const wchar_t *path, const wchar_t *contentType,
                              const QByteArray &data, OPC_COMPRESSION_OPTIONS compression)
    {
        com_ptr<IOpcPartUri> uri;
        check_hresult(factory->CreatePartUri(path, uri.put()));
        com_ptr<IOpcPart> part;
        check_hresult(parts->CreatePart(uri.get(), contentType, compression, part.put()));
        com_ptr<IStream> stream;
        check_hresult(part->GetContentStream(stream.put()));
        ULONG written = 0;
        check_hresult(stream->Write(data.constData(), static_cast<ULONG>(data.size()), &written));
        if (written != data.size())
            winrt::throw_hresult(STG_E_WRITEFAULT);
        return part;
    }

    void linkDocument()
    {
        com_ptr<IOpcRelationshipSet> relationships;
        check_hresult(package->GetRelationshipSet(relationships.put()));
        addRelationship(relationships.get(), L"rIdDocument", officeRelationship,
                        L"word/document.xml");
    }

    void linkImage(IOpcPart *document, const QString &id, const QString &relativePath)
    {
        com_ptr<IOpcRelationshipSet> relationships;
        check_hresult(document->GetRelationshipSet(relationships.put()));
        addRelationship(relationships.get(), id.toStdWString().c_str(), imageRelationship,
                        relativePath.toStdWString().c_str());
    }

    void save(const QString &path)
    {
        com_ptr<IStream> stream;
        check_hresult(factory->CreateStreamOnFile(QDir::toNativeSeparators(path).toStdWString().c_str(),
                                                  OPC_STREAM_IO_WRITE, nullptr, FILE_ATTRIBUTE_NORMAL,
                                                  stream.put()));
        check_hresult(factory->WritePackageToStream(package.get(), OPC_WRITE_DEFAULT, stream.get()));
        check_hresult(stream->Commit(STGC_DEFAULT));
    }

private:
    static void addRelationship(IOpcRelationshipSet *set, const wchar_t *id,
                                const wchar_t *type, const wchar_t *relativePath)
    {
        com_ptr<IUri> target;
        check_hresult(CreateUri(relativePath, Uri_CREATE_ALLOW_RELATIVE, 0, target.put()));
        com_ptr<IOpcRelationship> relationship;
        check_hresult(set->CreateRelationship(id, type, target.get(), OPC_URI_TARGET_MODE_INTERNAL,
                                              relationship.put()));
    }

    com_ptr<IOpcFactory> factory;
    com_ptr<IOpcPackage> package;
    com_ptr<IOpcPartSet> parts;
};

PdfDocument loadPdf(const QString &source)
{
    const auto path = QDir::toNativeSeparators(source).toStdWString();
    IRandomAccessStream stream{nullptr};
    const HRESULT status = CreateRandomAccessStreamOnFile(
        path.c_str(), static_cast<DWORD>(FileAccessMode::Read),
        winrt::guid_of<IRandomAccessStream>(), winrt::put_abi(stream));
    if (FAILED(status)) {
        const winrt::hresult_error error(status);
        throw QStringLiteral("无法打开 PDF 文件。%1")
            .arg(QString::fromWCharArray(error.message().c_str()));
    }

    try {
        const auto pdf = PdfDocument::LoadFromStreamAsync(stream).get();
        if (!pdf.PageCount())
            throw QStringLiteral("这个 PDF 没有可转换的页面。");
        return pdf;
    } catch (const winrt::hresult_error &error) {
        if (error.code() == HRESULT_FROM_WIN32(ERROR_WRONG_PASSWORD))
            throw QStringLiteral("请先解除 PDF 的密码保护。");
        throw QStringLiteral("无法读取 PDF。请确认文件完整，且未设置打开密码。\n%1")
            .arg(QString::fromWCharArray(error.message().c_str()));
    }
}

QByteArray renderPage(const PdfPage &page, const QSize &pageTwips)
{
    PdfPageRenderOptions options;
    options.DestinationWidth(std::max(1, qRound(pageTwips.width() * renderDpi / twipsPerInch)));
    options.DestinationHeight(std::max(1, qRound(pageTwips.height() * renderDpi / twipsPerInch)));
    options.BackgroundColor({255, 255, 255, 255});
    options.IsIgnoringHighContrast(true);
    options.BitmapEncoderId(BitmapEncoder::PngEncoderId());
    InMemoryRandomAccessStream stream;
    page.RenderToStreamAsync(stream, options).get();

    if (!stream.Size() || stream.Size() > std::numeric_limits<int>::max())
        throw QStringLiteral("无法读取页面图像。");
    const auto length = static_cast<uint32_t>(stream.Size());
    DataReader reader(stream.GetInputStreamAt(0));
    if (reader.LoadAsync(length).get() != length)
        throw QStringLiteral("PDF 页面图像读取不完整。");
    QByteArray image(length, Qt::Uninitialized);
    auto *bytes = reinterpret_cast<uint8_t *>(image.data());
    reader.ReadBytes({bytes, bytes + length});
    return image;
}

QString sectionXml(const QSize &size, int horizontalMargin = 0, int verticalMargin = 0)
{
    const auto orientation = size.width() > size.height() ? u"landscape" : u"portrait";
    return QStringLiteral(R"xml(<w:sectPr>
<w:type w:val="nextPage"/>
<w:pgSz w:w="%1" w:h="%2" w:orient="%3"/>
<w:pgMar w:top="%5" w:right="%4" w:bottom="%5" w:left="%4" w:header="0" w:footer="0" w:gutter="0"/>
</w:sectPr>)xml").arg(size.width()).arg(size.height()).arg(QStringView(orientation))
        .arg(horizontalMargin).arg(verticalMargin);
}

QString pictureXml(int number, const QSize &size, bool pageBreak, const QString &section = {})
{
    // The full-page picture is positioned independently of the paragraph's line height.
    return QStringLiteral(R"xml(<w:p><w:pPr>%1
<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>
%2
</w:pPr><w:r><w:drawing>
<wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="%3" cy="%4"/><wp:wrapNone/>
<wp:docPr id="%5" name="PDF page %5" descr="PDF 第 %5 页"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic><pic:nvPicPr><pic:cNvPr id="%5" name="PDF page %5"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="rId%5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%3" cy="%4"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>)xml")
        .arg(pageBreak ? QStringLiteral("<w:pageBreakBefore/>") : QString())
        .arg(section)
        .arg(size.width() * emusPerTwip).arg(size.height() * emusPerTwip)
        .arg(number);
}

QSize pageSizeInTwips(const PdfPage &page, int pageNumber)
{
    const auto dips = page.Size();
    if (!std::isfinite(dips.Width) || !std::isfinite(dips.Height)
        || dips.Width <= 0 || dips.Height <= 0)
        throw QStringLiteral("PDF 第 %1 页的尺寸无效。").arg(pageNumber);
    const double longestSide = std::max(dips.Width, dips.Height);
    const double scale = std::min(1.0, maxWordPageDips / longestSide)
                         * twipsPerInch / dipsPerInch;
    return QSize(std::max(1, qRound(dips.Width * scale)),
                 std::max(1, qRound(dips.Height * scale)));
}

QString startDocumentXml()
{
    return QStringLiteral(R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>)xml");
}

void convertLayout(const PdfDocument &pdf, const QString &output, const PageProgress &progress)
{
    DocxPackage package;
    package.linkDocument();
    const int total = static_cast<int>(pdf.PageCount());
    QString xml = startDocumentXml();

    QSize previousSize;
    bool previousPageBreak = false;
    for (int index = 0; index < total; ++index) {
        const auto page = pdf.GetPage(index);
        const auto size = pageSizeInTwips(page, index + 1);
        if (previousSize.isValid()) {
            // A section break belongs to the last picture paragraph in that section.
            const auto section = size != previousSize ? sectionXml(previousSize) : QString();
            xml += pictureXml(index, previousSize, previousPageBreak, section);
        }

        const auto relativePath = QStringLiteral("media/page-%1.png").arg(index + 1);
        const auto partPath = (QStringLiteral("/word/") + relativePath).toStdWString();
        package.addPart(partPath.c_str(), L"image/png", renderPage(page, size), OPC_COMPRESSION_NONE);
        previousPageBreak = size == previousSize;
        previousSize = size;
        if (progress)
            progress(index + 1, total);
    }

    xml += pictureXml(total, previousSize, previousPageBreak);
    xml += sectionXml(previousSize) + QStringLiteral("</w:body></w:document>");
    const auto document = package.addPart(
        L"/word/document.xml",
        L"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        xml.toUtf8(), OPC_COMPRESSION_NORMAL);
    for (int number = 1; number <= total; ++number)
        package.linkImage(document.get(), QStringLiteral("rId%1").arg(number),
                          QStringLiteral("media/page-%1.png").arg(number));
    package.save(output);
}

struct EditablePage
{
    QSize size;
    QStringList paragraphs;
};

QString editableRunXml(QString text, int fontTwips)
{
    // Extracted text can contain control characters that XML 1.0 cannot represent.
    static const QRegularExpression invalidCharacters(
        QStringLiteral("[\\x{0000}-\\x{0008}\\x{000B}\\x{000C}\\x{000E}-\\x{001F}\\x{FFFE}\\x{FFFF}]"));
    if (!text.isValidUtf16())
        text = QString::fromUtf8(text.toUtf8());
    text.remove(invalidCharacters);
    text.replace(QStringLiteral("\r\n"), QStringLiteral("\n"));
    text.replace(u'\r', u'\n');

    QString result;
    QXmlStreamWriter writer(&result);
    writer.writeStartElement(QStringLiteral("w:r"));
    writer.writeStartElement(QStringLiteral("w:rPr"));
    const auto halfPoints = QString::number(std::max(2, fontTwips / 10));
    writer.writeEmptyElement(QStringLiteral("w:sz"));
    writer.writeAttribute(QStringLiteral("w:val"), halfPoints);
    writer.writeEmptyElement(QStringLiteral("w:szCs"));
    writer.writeAttribute(QStringLiteral("w:val"), halfPoints);
    writer.writeEndElement();
    const auto lines = QStringView(text).split(u'\n');
    for (qsizetype lineIndex = 0; lineIndex < lines.size(); ++lineIndex) {
        if (lineIndex)
            writer.writeEmptyElement(QStringLiteral("w:br"));
        const auto runs = lines[lineIndex].split(u'\t');
        for (qsizetype runIndex = 0; runIndex < runs.size(); ++runIndex) {
            if (runIndex)
                writer.writeEmptyElement(QStringLiteral("w:tab"));
            if (runs[runIndex].isEmpty())
                continue;
            writer.writeStartElement(QStringLiteral("w:t"));
            writer.writeAttribute(QStringLiteral("xml:space"), QStringLiteral("preserve"));
            writer.writeCharacters(runs[runIndex].toString());
            writer.writeEndElement();
        }
    }
    writer.writeEndElement();
    if (writer.hasError())
        throw QStringLiteral("无法将 PDF 文本写入 Word 文档。");
    return result;
}

QSize editableMargins(const QSize &size)
{
    return {std::min(720, size.width() / 8), std::min(720, size.height() / 8)};
}

QString editableSectionXml(const EditablePage &page)
{
    if (page.paragraphs.isEmpty())
        return sectionXml(page.size);
    const auto margins = editableMargins(page.size);
    return sectionXml(page.size, margins.width(), margins.height());
}

// Word reflows extracted text. Pick the largest font (from 11 pt down to 4 pt)
// that still fits the page, so dense or tiny pages stay editable without
// spilling onto extra pages. Only pages with no text layer become pictures.
int editableFontTwips(const EditablePage &page)
{
    const auto margins = editableMargins(page.size);
    const int usableWidth = page.size.width() - 2 * margins.width();
    const int usableHeight = page.size.height() - 2 * margins.height();
    if (usableWidth <= 0 || usableHeight <= 0)
        return 80;

    constexpr int maxFont = 220; // 11 pt, in twips
    constexpr int minFont = 80;  // 4 pt
    for (int font = maxFont; font > minFont; font -= 20) {
        const int latin = std::max(1, font / 2);
        int lines = 0;
        for (const QString &paragraph : page.paragraphs) {
            int used = 0;
            int paragraphLines = 1;
            for (const QChar character : paragraph) {
                const int advance = character.unicode() > 0x2e80 ? font : latin;
                if (used + advance > usableWidth) {
                    ++paragraphLines;
                    used = 0;
                }
                used += advance;
            }
            lines += paragraphLines;
        }
        if (static_cast<qint64>(lines) * font * 12 / 10 <= usableHeight)
            return font;
    }
    return minFont;
}

bool shareEditableSection(const EditablePage &first, const EditablePage &second)
{
    return first.size == second.size && first.paragraphs.isEmpty() == second.paragraphs.isEmpty();
}

void saveEditableDocx(const PdfDocument &pdf, const std::vector<EditablePage> &pages,
                      const QString &output, const PageProgress &progress)
{
    if (pages.empty())
        throw QStringLiteral("这个 PDF 没有可转换的页面。");

    DocxPackage package;
    package.linkDocument();
    QString xml = startDocumentXml();
    std::vector<int> imagePages;

    for (std::size_t index = 0; index < pages.size(); ++index) {
        const auto &page = pages[index];
        const int number = static_cast<int>(index + 1);
        const bool pageBreak = index > 0 && shareEditableSection(pages[index - 1], page);
        const bool endsSection = index + 1 < pages.size()
                              && !shareEditableSection(page, pages[index + 1]);
        const auto section = endsSection ? editableSectionXml(page) : QString();
        if (page.paragraphs.isEmpty()) {
            const auto partPath = QStringLiteral("/word/media/page-%1.png").arg(number).toStdWString();
            const auto image = renderPage(pdf.GetPage(number - 1), page.size);
            package.addPart(partPath.c_str(), L"image/png", image, OPC_COMPRESSION_NONE);
            xml += pictureXml(number, page.size, pageBreak, section);
            imagePages.push_back(number);
        } else {
            const int font = editableFontTwips(page);
            for (qsizetype paragraphIndex = 0; paragraphIndex < page.paragraphs.size(); ++paragraphIndex) {
                xml += QStringLiteral("<w:p><w:pPr>");
                if (paragraphIndex == 0 && pageBreak)
                    xml += QStringLiteral("<w:pageBreakBefore/>");
                xml += QStringLiteral("<w:spacing w:after=\"0\"/>");
                if (paragraphIndex == page.paragraphs.size() - 1)
                    xml += section;
                xml += QStringLiteral("</w:pPr>");
                xml += editableRunXml(page.paragraphs[paragraphIndex], font);
                xml += QStringLiteral("</w:p>");
            }
        }
        if (progress)
            progress(number, static_cast<int>(pages.size()));
    }

    xml += editableSectionXml(pages.back()) + QStringLiteral("</w:body></w:document>");
    const auto document = package.addPart(
        L"/word/document.xml",
        L"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        xml.toUtf8(), OPC_COMPRESSION_NORMAL);
    for (const int number : imagePages) {
        package.linkImage(document.get(), QStringLiteral("rId%1").arg(number),
                          QStringLiteral("media/page-%1.png").arg(number));
    }
    package.save(output);
}

QStringList extractPageLines(FPDF_PAGE page, int pageNumber)
{
    const auto textPage = FPDFText_LoadPage(page);
    if (!textPage)
        throw QStringLiteral("无法读取 PDF 第 %1 页的文字。").arg(pageNumber);
    const auto closeTextPage = qScopeGuard([textPage] { FPDFText_ClosePage(textPage); });

    const int count = FPDFText_CountChars(textPage);
    if (count < 0)
        throw QStringLiteral("无法读取 PDF 第 %1 页的文字。").arg(pageNumber);

    QString text;
    text.reserve(count);
    for (int index = 0; index < count; ++index) {
        const auto character = static_cast<char32_t>(FPDFText_GetUnicode(textPage, index));
        if (character == 0 || QChar::isSurrogate(character))
            continue;
        if (character < 0x10000) {
            text += QChar(static_cast<char16_t>(character));
        } else if (character <= 0x10ffff) {
            text += QChar(QChar::highSurrogate(character));
            text += QChar(QChar::lowSurrogate(character));
        }
    }

    if (text.trimmed().isEmpty())
        return {};
    text.replace(QStringLiteral("\r\n"), QStringLiteral("\n"));
    text.replace(u'\r', u'\n');
    while (text.endsWith(u'\n'))
        text.chop(1);
    return text.split(u'\n');
}

void ensurePdfium()
{
    // pdfium.dll is delay-loaded, so a missing DLL would otherwise surface as a
    // low-level loader exception on the first PDFium call.
    if (GetModuleHandleW(L"pdfium.dll"))
        return;
    if (!LoadLibraryW(L"pdfium.dll"))
        throw QStringLiteral("缺少 pdfium.dll，无法提取可编辑文字。请把 pdfium.dll 放到程序所在文件夹，"
                             "或改用「保留原版式」模式。");
}

std::vector<EditablePage> readEditablePages(const PdfDocument &pdf, const QString &source)
{
    ensurePdfium();

    QFile file(source);
    if (!file.open(QIODevice::ReadOnly))
        throw QStringLiteral("无法打开 PDF 文件。%1").arg(file.errorString());

    const auto fileSize = file.size();
    if (fileSize <= 0)
        throw QStringLiteral("PDF 文件为空。");
    if (fileSize > std::numeric_limits<unsigned long>::max())
        throw QStringLiteral("此 PDF 超出可编辑文本模式支持的文件大小。");

    FPDF_FILEACCESS access{};
    access.m_FileLen = static_cast<unsigned long>(fileSize);
    access.m_Param = &file;
    access.m_GetBlock = [](void *context, unsigned long position, unsigned char *buffer,
                           unsigned long size) -> int {
        auto &sourceFile = *static_cast<QFile *>(context);
        return sourceFile.seek(position) && sourceFile.read(reinterpret_cast<char *>(buffer), size) == size;
    };

    // PDFium requires serialized access, including initialization and cleanup.
    static QMutex pdfiumMutex;
    const QMutexLocker lock(&pdfiumMutex);
    FPDF_InitLibrary();
    const auto destroyLibrary = qScopeGuard([] { FPDF_DestroyLibrary(); });

    const auto document = FPDF_LoadCustomDocument(&access, nullptr);
    if (!document) {
        if (FPDF_GetLastError() == FPDF_ERR_PASSWORD)
            throw QStringLiteral("请先解除 PDF 的密码保护。");
        throw QStringLiteral("无法提取文字。请确认 PDF 文件完整。");
    }
    const auto closeDocument = qScopeGuard([document] { FPDF_CloseDocument(document); });

    const int total = FPDF_GetPageCount(document);
    if (total <= 0 || static_cast<uint32_t>(total) != pdf.PageCount())
        throw QStringLiteral("无法确认 PDF 页数。请重新保存 PDF 后重试。");

    std::vector<EditablePage> pages;
    pages.reserve(total);
    for (int index = 0; index < total; ++index) {
        const auto page = FPDF_LoadPage(document, index);
        if (!page)
            throw QStringLiteral("无法读取 PDF 第 %1 页。").arg(index + 1);
        const auto closePage = qScopeGuard([page] { FPDF_ClosePage(page); });
        pages.push_back({pageSizeInTwips(pdf.GetPage(index), index + 1), extractPageLines(page, index + 1)});
    }
    return pages;
}

void convertEditable(const PdfDocument &pdf, const QString &source, const QString &output,
                     const PageProgress &progress)
{
    const auto pages = readEditablePages(pdf, source);
    if (std::none_of(pages.begin(), pages.end(), [](const EditablePage &page) {
            return !page.paragraphs.isEmpty();
        })) {
        throw QStringLiteral("这个 PDF 没有可提取的文字。扫描件可使用「保留原版式」，或先做 OCR 再转换。");
    }
    saveEditableDocx(pdf, pages, output, progress);
}

} // namespace

QString convertPdf(const QString &source, const QString &destination, ConversionMode mode,
                   const PageProgress &progress, bool replaceExisting)
{
    const QFileInfo input(source);
    const QFileInfo output(destination);
    const auto outputExistsError = QStringLiteral("输出位置已有同名文件。请选择其他文件名，或确认覆盖后重试。");
    if (input.suffix().compare(QStringLiteral("pdf"), Qt::CaseInsensitive) != 0 || !input.isFile())
        return QStringLiteral("请选择一个有效的 PDF 文件。");
    if (output.suffix().compare(QStringLiteral("docx"), Qt::CaseInsensitive) != 0)
        return QStringLiteral("请使用 .docx 扩展名保存 Word 文件。");
    if (!output.absoluteDir().exists())
        return QStringLiteral("输出文件夹不存在。");
    if (output.isDir())
        return QStringLiteral("输出路径指向文件夹。请指定要保存的 .docx 文件名。");
    if (!replaceExisting && (output.exists() || output.isSymLink()))
        return outputExistsError;

    QTemporaryDir staging(output.absoluteDir().filePath(QStringLiteral(".pdf-to-word-XXXXXX")));
    if (!staging.isValid())
        return QStringLiteral("无法创建临时文件。请检查输出文件夹的写入权限。");
    const QString temporaryDocx = staging.filePath(QStringLiteral("result.docx"));
    try {
        {
            Apartment apartment(winrt::apartment_type::multi_threaded);
            const auto pdf = loadPdf(input.absoluteFilePath());
            if (mode == ConversionMode::PreserveLayout)
                convertLayout(pdf, temporaryDocx, progress);
            else
                convertEditable(pdf, input.absoluteFilePath(), temporaryDocx, progress);
        }

        const auto from = QDir::toNativeSeparators(temporaryDocx).toStdWString();
        const auto to = QDir::toNativeSeparators(output.absoluteFilePath()).toStdWString();
        DWORD moveFlags = MOVEFILE_WRITE_THROUGH;
        if (replaceExisting)
            moveFlags |= MOVEFILE_REPLACE_EXISTING;
        if (!MoveFileExW(from.c_str(), to.c_str(), moveFlags)) {
            const DWORD error = GetLastError();
            if (!replaceExisting && (error == ERROR_FILE_EXISTS || error == ERROR_ALREADY_EXISTS))
                return outputExistsError;
            if (error == ERROR_SHARING_VIOLATION || error == ERROR_LOCK_VIOLATION)
                return QStringLiteral("目标文件正在使用。请关闭该文件后重试。");
            return QStringLiteral("无法保存 Word 文件。请检查输出位置的写入权限和可用空间。");
        }
        return {};
    } catch (const QString &error) {
        return error;
    } catch (const winrt::hresult_error &error) {
        return QStringLiteral("转换失败。%1").arg(QString::fromWCharArray(error.message().c_str()));
    } catch (const std::exception &error) {
        return QStringLiteral("转换失败。%1").arg(QString::fromLocal8Bit(error.what()));
    }
}
