#include "converter.h"

#include <QDir>
#include <QFileInfo>
#include <QScopeGuard>
#include <QSize>
#include <QTemporaryDir>

#include <windows.h>
#include <msopc.h>
#include <winrt/Windows.Data.Pdf.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Windows.Storage.h>
#include <winrt/Windows.Storage.Streams.h>
#include <winrt/Windows.UI.h>

#include <algorithm>
#include <cmath>
#include <initializer_list>
#include <limits>
#include <vector>

namespace {

using winrt::check_hresult;
using winrt::com_ptr;
using winrt::Windows::Data::Pdf::PdfDocument;
using winrt::Windows::Data::Pdf::PdfPage;
using winrt::Windows::Data::Pdf::PdfPageRenderOptions;
using winrt::Windows::Graphics::Imaging::BitmapEncoder;
using winrt::Windows::Storage::StorageFile;
using winrt::Windows::Storage::Streams::DataReader;
using winrt::Windows::Storage::Streams::InMemoryRandomAccessStream;

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
    try {
        const auto file = StorageFile::GetFileFromPathAsync(
            QDir::toNativeSeparators(source).toStdWString()).get();
        const auto pdf = PdfDocument::LoadFromFileAsync(file).get();
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

QString sectionXml(const QSize &size)
{
    const auto orientation = size.width() > size.height() ? u"landscape" : u"portrait";
    return QStringLiteral(R"xml(<w:sectPr>
<w:type w:val="nextPage"/>
<w:pgSz w:w="%1" w:h="%2" w:orient="%3"/>
<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>
</w:sectPr>)xml").arg(size.width()).arg(size.height()).arg(QStringView(orientation));
}

QString pictureXml(int number, const QSize &size, const QString &id, bool pageBreak)
{
    // The full-page picture is positioned independently of the paragraph's line height.
    return QStringLiteral(R"xml(<w:p><w:pPr>%1
<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>
</w:pPr><w:r><w:drawing>
<wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="%2" cy="%3"/><wp:wrapNone/>
<wp:docPr id="%4" name="PDF page %4" descr="PDF 第 %4 页"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic><pic:nvPicPr><pic:cNvPr id="%4" name="PDF page %4"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="%5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%2" cy="%3"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>)xml")
        .arg(pageBreak ? QStringLiteral("<w:pageBreakBefore/>") : QString())
        .arg(size.width() * emusPerTwip).arg(size.height() * emusPerTwip)
        .arg(number).arg(id);
}

void convertLayout(const PdfDocument &pdf, const QString &output, const PageProgress &progress)
{
    DocxPackage package;
    package.linkDocument();
    const int total = static_cast<int>(pdf.PageCount());
    QString xml = QStringLiteral(R"xml(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>)xml");

    QSize previousSize;
    for (int index = 0; index < total; ++index) {
        const auto page = pdf.GetPage(index);
        const auto dips = page.Size();
        const double longestSide = std::max(dips.Width, dips.Height);
        if (!std::isfinite(longestSide) || dips.Width <= 0 || dips.Height <= 0)
            throw QStringLiteral("PDF 第 %1 页的尺寸无效。").arg(index + 1);
        const double scale = std::min(1.0, maxWordPageDips / longestSide)
                             * twipsPerInch / dipsPerInch;
        const QSize size(std::max(1, qRound(dips.Width * scale)),
                         std::max(1, qRound(dips.Height * scale)));
        if (previousSize.isValid() && size != previousSize) {
            xml += QStringLiteral(R"xml(<w:p><w:pPr>
<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>%1
</w:pPr></w:p>)xml").arg(sectionXml(previousSize));
        }

        const auto relativePath = QStringLiteral("media/page-%1.png").arg(index + 1);
        const auto partPath = (QStringLiteral("/word/") + relativePath).toStdWString();
        package.addPart(partPath.c_str(), L"image/png", renderPage(page, size), OPC_COMPRESSION_NONE);
        xml += pictureXml(index + 1, size, QStringLiteral("rId%1").arg(index + 1), size == previousSize);
        previousSize = size;
        if (progress)
            progress(index + 1, total);
    }

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

class AutomationValue : public VARIANT
{
public:
    AutomationValue() { VariantInit(this); }
    AutomationValue(bool value) : AutomationValue()
    {
        vt = VT_BOOL;
        boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    }
    AutomationValue(int value) : AutomationValue()
    {
        vt = VT_I4;
        lVal = value;
    }
    AutomationValue(const QString &value) : AutomationValue()
    {
        vt = VT_BSTR;
        bstrVal = SysAllocStringLen(reinterpret_cast<const wchar_t *>(value.utf16()),
                                    static_cast<UINT>(value.size()));
        if (!bstrVal)
            winrt::throw_hresult(E_OUTOFMEMORY);
    }
    ~AutomationValue() { VariantClear(this); }
    AutomationValue(const AutomationValue &) = delete;
    AutomationValue &operator=(const AutomationValue &) = delete;
};

void invokeWord(IDispatch *object, const wchar_t *member, WORD flags,
                std::initializer_list<AutomationValue> arguments = {}, VARIANT *result = nullptr)
{
    auto name = const_cast<wchar_t *>(member);
    DISPID id;
    check_hresult(object->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id));

    // IDispatch takes positional arguments in reverse order.
    std::vector<VARIANTARG> reversed(arguments.size());
    std::reverse_copy(arguments.begin(), arguments.end(), reversed.begin());
    DISPPARAMS parameters{reversed.data(), nullptr, static_cast<UINT>(reversed.size()), 0};
    DISPID propertyPut = DISPID_PROPERTYPUT;
    if (flags == DISPATCH_PROPERTYPUT) {
        parameters.rgdispidNamedArgs = &propertyPut;
        parameters.cNamedArgs = 1;
    }

    EXCEPINFO details{};
    const auto clear = qScopeGuard([&details] {
        SysFreeString(details.bstrSource);
        SysFreeString(details.bstrDescription);
        SysFreeString(details.bstrHelpFile);
    });
    const HRESULT status = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                                          &parameters, result, &details, nullptr);
    if (FAILED(status)) {
        if (details.pfnDeferredFillIn)
            details.pfnDeferredFillIn(&details);
        if (details.bstrDescription)
            throw QStringLiteral("Microsoft Word 转换失败。%1").arg(
                QString::fromWCharArray(details.bstrDescription));
        check_hresult(status);
    }
}

com_ptr<IDispatch> wordObject(IDispatch *parent, const wchar_t *member, WORD flags,
                              std::initializer_list<AutomationValue> arguments = {})
{
    AutomationValue value;
    invokeWord(parent, member, flags, arguments, &value);
    if (value.vt != VT_DISPATCH || !value.pdispVal)
        throw QStringLiteral("Microsoft Word 无法打开文档。");
    com_ptr<IDispatch> object;
    object.copy_from(value.pdispVal);
    return object;
}

void closeWord(IDispatch *object, const wchar_t *method) noexcept
{
    try {
        invokeWord(object, method, DISPATCH_METHOD, {false});
    } catch (...) {
        // Word can exit as soon as its last automation document is closed.
    }
}

void convertEditable(const QString &source, const QString &output)
{
    Apartment apartment(winrt::apartment_type::single_threaded);
    CLSID wordId{};
    com_ptr<IDispatch> word;
    if (FAILED(CLSIDFromProgID(L"Word.Application", &wordId))
        || FAILED(CoCreateInstance(wordId, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                                   word.put_void()))) {
        throw QStringLiteral("可编辑文本模式需要安装桌面版 Microsoft Word。");
    }
    AutomationValue version;
    invokeWord(word.get(), L"Version", DISPATCH_PROPERTYGET, {}, &version);
    if (version.vt != VT_BSTR || !version.bstrVal
        || QString::fromWCharArray(version.bstrVal).section(u'.', 0, 0).toInt() < 15) {
        throw QStringLiteral("可编辑文本模式需要 Microsoft Word 2013 或更新版本提供 PDF 转换。");
    }
    const auto quit = qScopeGuard([&word] { closeWord(word.get(), L"Quit"); });
    invokeWord(word.get(), L"Visible", DISPATCH_PROPERTYPUT, {false});
    invokeWord(word.get(), L"DisplayAlerts", DISPATCH_PROPERTYPUT, {0});
    invokeWord(word.get(), L"AutomationSecurity", DISPATCH_PROPERTYPUT, {3});

    const auto documents = wordObject(word.get(), L"Documents", DISPATCH_PROPERTYGET);
    const auto document = wordObject(documents.get(), L"Open", DISPATCH_METHOD,
                                     {QDir::toNativeSeparators(source), false, true, false});
    const auto close = qScopeGuard([&document] { closeWord(document.get(), L"Close"); });
    invokeWord(document.get(), L"SaveAs2", DISPATCH_METHOD,
               {QDir::toNativeSeparators(output), 16});
    if (QFileInfo(output).size() <= 0)
        throw QStringLiteral("Microsoft Word 无法保存 DOCX。");
}

} // namespace

QString convertPdf(const QString &source, const QString &destination, ConversionMode mode,
                   const PageProgress &progress)
{
    const QFileInfo input(source);
    const QFileInfo output(destination);
    if (input.suffix().compare(QStringLiteral("pdf"), Qt::CaseInsensitive) != 0 || !input.isFile())
        return QStringLiteral("请选择一个有效的 PDF 文件。");
    if (output.suffix().compare(QStringLiteral("docx"), Qt::CaseInsensitive) != 0)
        return QStringLiteral("请使用 .docx 扩展名保存 Word 文件。");
    if (!output.absoluteDir().exists())
        return QStringLiteral("输出文件夹不存在。");

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
        }
        if (mode == ConversionMode::Editable)
            convertEditable(input.absoluteFilePath(), temporaryDocx);

        const auto from = QDir::toNativeSeparators(temporaryDocx).toStdWString();
        const auto to = QDir::toNativeSeparators(output.absoluteFilePath()).toStdWString();
        if (!MoveFileExW(from.c_str(), to.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
            return QStringLiteral("无法替换目标文件。请关闭已打开的 Word 文件并检查写入权限。");
        return {};
    } catch (const QString &error) {
        return error;
    } catch (const winrt::hresult_error &error) {
        return QStringLiteral("转换失败。%1").arg(QString::fromWCharArray(error.message().c_str()));
    } catch (const std::exception &error) {
        return QStringLiteral("转换失败。%1").arg(QString::fromLocal8Bit(error.what()));
    }
}
