#include "converter.h"

#include <QApplication>
#include <QCloseEvent>
#include <QCommandLineParser>
#include <QDesktopServices>
#include <QDir>
#include <QDragEnterEvent>
#include <QDragLeaveEvent>
#include <QDragMoveEvent>
#include <QDropEvent>
#include <QFileDialog>
#include <QFileInfo>
#include <QFrame>
#include <QFutureWatcher>
#include <QGridLayout>
#include <QHBoxLayout>
#include <QKeySequence>
#include <QLabel>
#include <QLineEdit>
#include <QMessageBox>
#include <QMimeData>
#include <QPainter>
#include <QProgressBar>
#include <QPushButton>
#include <QRadioButton>
#include <QShortcut>
#include <QStyle>
#include <QTextStream>
#include <QUrl>
#include <QVBoxLayout>
#include <QWidget>
#include <QtConcurrentRun>

namespace {

QString defaultOutput(const QString &source)
{
    const QFileInfo info(source);
    return info.dir().filePath(info.completeBaseName() + QStringLiteral(".docx"));
}

QString pathFromField(const QString &text)
{
    QString path = text.trimmed();
    if (path.size() >= 2 && path.startsWith(u'"') && path.endsWith(u'"'))
        path = path.mid(1, path.size() - 2);
    return QDir::fromNativeSeparators(path);
}

bool isPdfPath(const QString &path)
{
    return path.endsWith(QStringLiteral(".pdf"), Qt::CaseInsensitive);
}

QString droppedPdf(const QMimeData &data)
{
    const auto urls = data.urls();
    if (urls.size() != 1 || !urls.first().isLocalFile())
        return {};
    const QString path = urls.first().toLocalFile();
    return isPdfPath(path) ? path : QString{};
}

constexpr auto windowStyle = R"qss(
QWidget#ConverterWindow { background: #f3f5f8; color: #172b40; }
QLabel#Title { color: #13273d; font-size: 30px; font-weight: 700; }
QFrame#Form, QFrame#Feedback {
    background: #ffffff;
    border: 1px solid #d9e1eb;
    border-radius: 14px;
}
QFrame#DropArea {
    background: #f0f6fc;
    border: 1px dashed #8aa9c9;
    border-radius: 9px;
}
QFrame#DropArea[dragging="true"] { background: #dfedfb; border-color: #1761a7; }
QFrame#DropArea:disabled { background: #f3f5f8; border-color: #d2dce7; }
QLabel[fieldLabel="true"], QLabel#DropTitle, QLabel#StatusTitle {
    color: #203852;
    font-weight: 600;
}
QLabel#DropTitle { font-size: 15px; }
QLabel#ModeHint, QLabel#StatusDetail { color: #4b6076; font-size: 13px; }
QLineEdit, QPushButton, QRadioButton {
    color: #203852;
    font-size: 14px;
    border: 1px solid #c7d5e3;
    border-radius: 7px;
}
QLineEdit { background: #f9fbfd; padding: 8px 11px; selection-background-color: #2367a5; }
QPushButton { background: #ffffff; padding: 8px 13px; font-weight: 600; }
QPushButton:hover { background: #ecf4fd; border-color: #91b5d9; }
QPushButton:pressed { background: #dfedfa; }
QRadioButton { background: #f9fbfd; padding: 10px 11px; spacing: 8px; }
QRadioButton:hover { background: #f0f6fc; }
QRadioButton:checked { background: #ecf4fd; border-color: #73a3d1; color: #174d81; }
QLineEdit:focus, QPushButton:focus, QRadioButton:focus { border-color: #1761a7; }
QRadioButton::indicator {
    width: 8px;
    height: 8px;
    border: 4px solid #d3dfec;
    border-radius: 8px;
    background: #ffffff;
}
QRadioButton::indicator:checked { border-color: #2367a5; }
QLineEdit:disabled, QPushButton:disabled, QRadioButton:disabled { color: #66788a; }
QPushButton#UseDefault { padding: 4px 7px; font-size: 12px; }
QPushButton#Convert {
    background: #1761a7;
    border-color: #1761a7;
    color: #ffffff;
    padding: 10px 19px;
    font-size: 15px;
}
QPushButton#Convert:hover { background: #124f8a; }
QPushButton#Convert:pressed { background: #104478; }
QPushButton#Convert:focus { border-color: #092d51; }
QPushButton#Convert:disabled { background: #e5edf6; border-color: #e5edf6; color: #496681; }
QProgressBar { background: #e0e9f3; border: none; border-radius: 3px; }
QProgressBar::chunk { background: #2367a5; border-radius: 3px; }
)qss";

class PdfIcon final : public QWidget
{
public:
    explicit PdfIcon(QWidget *parent) : QWidget(parent) { setFixedSize(40, 40); }

protected:
    void paintEvent(QPaintEvent *) override
    {
        QPainter painter(this);
        painter.setRenderHint(QPainter::Antialiasing);
        painter.setPen(Qt::NoPen);
        painter.setBrush(QColor(QStringLiteral("#dceafa")));
        painter.drawRoundedRect(rect(), 9, 9);

        painter.setPen(QPen(QColor(QStringLiteral("#2367a5")), 1.5, Qt::SolidLine, Qt::RoundCap));
        painter.setBrush(Qt::NoBrush);
        painter.drawRoundedRect(QRectF(12, 8, 16, 24), 3, 3);
        painter.drawLine(QPointF(16, 16), QPointF(24, 16));
        painter.drawLine(QPointF(16, 21), QPointF(24, 21));
        painter.drawLine(QPointF(16, 26), QPointF(21, 26));
    }
};

class MainWindow final : public QWidget
{
public:
    MainWindow()
    {
        setWindowTitle(QStringLiteral("PDF 转 Word"));
        setObjectName(QStringLiteral("ConverterWindow"));
        setAcceptDrops(true);
        setMinimumWidth(540);
        setStyleSheet(QString::fromLatin1(windowStyle));

        auto *layout = new QVBoxLayout(this);
        layout->setContentsMargins(24, 22, 24, 24);
        layout->setSpacing(14);
        auto *title = new QLabel(windowTitle(), this);
        title->setObjectName(QStringLiteral("Title"));
        layout->addWidget(title);

        auto *form = new QFrame(this);
        form->setObjectName(QStringLiteral("Form"));
        auto *formLayout = new QGridLayout(form);
        formLayout->setContentsMargins(18, 16, 18, 16);
        formLayout->setHorizontalSpacing(9);
        formLayout->setVerticalSpacing(8);
        formLayout->setColumnStretch(0, 1);

        auto *sourceLabel = new QLabel(QStringLiteral("源 PDF"), form);
        sourceLabel->setProperty("fieldLabel", true);
        sourceLabel->setBuddy(&input);
        formLayout->addWidget(sourceLabel, 0, 0, 1, 2);

        dropArea.setObjectName(QStringLiteral("DropArea"));
        auto *dropLayout = new QHBoxLayout(&dropArea);
        dropLayout->setContentsMargins(12, 10, 12, 10);
        dropLayout->setSpacing(11);
        auto *dropTitle = new QLabel(QStringLiteral("将 PDF 文件拖到这里"), &dropArea);
        dropTitle->setObjectName(QStringLiteral("DropTitle"));
        dropLayout->addWidget(new PdfIcon(&dropArea));
        dropLayout->addWidget(dropTitle, 1);
        dropLayout->addWidget(&browseInput);
        formLayout->addWidget(&dropArea, 1, 0, 1, 2);

        input.setPlaceholderText(QStringLiteral("也可以粘贴 PDF 文件路径"));
        input.setAcceptDrops(false);
        input.setClearButtonEnabled(true);
        formLayout->addWidget(&input, 2, 0, 1, 2);
        formLayout->setRowMinimumHeight(3, 3);

        auto *outputLabel = new QLabel(QStringLiteral("保存为"), form);
        outputLabel->setProperty("fieldLabel", true);
        outputLabel->setBuddy(&output);
        useDefault.setObjectName(QStringLiteral("UseDefault"));
        formLayout->addWidget(outputLabel, 4, 0);
        formLayout->addWidget(&useDefault, 4, 1, Qt::AlignRight);
        output.setPlaceholderText(QStringLiteral("默认保存在 PDF 所在文件夹"));
        output.setAcceptDrops(false);
        formLayout->addWidget(&output, 5, 0);
        formLayout->addWidget(&browseOutput, 5, 1);
        formLayout->setRowMinimumHeight(6, 3);

        auto *modeLabel = new QLabel(QStringLiteral("转换模式"), form);
        modeLabel->setProperty("fieldLabel", true);
        modeLabel->setBuddy(&preserveLayout);
        formLayout->addWidget(modeLabel, 7, 0, 1, 2);
        auto *modeLayout = new QHBoxLayout;
        modeLayout->setSpacing(9);
        preserveLayout.setChecked(true);
        modeLayout->addWidget(&preserveLayout, 1);
        modeLayout->addWidget(&editable, 1);
        formLayout->addLayout(modeLayout, 8, 0, 1, 2);
        hint.setObjectName(QStringLiteral("ModeHint"));
        hint.setWordWrap(true);
        formLayout->addWidget(&hint, 9, 0, 1, 2);
        layout->addWidget(form);

        auto *feedback = new QFrame(this);
        feedback->setObjectName(QStringLiteral("Feedback"));
        auto *feedbackLayout = new QHBoxLayout(feedback);
        feedbackLayout->setContentsMargins(14, 12, 14, 12);
        feedbackLayout->setSpacing(14);
        auto *statusLayout = new QVBoxLayout;
        statusLayout->setSpacing(6);
        statusTitle.setObjectName(QStringLiteral("StatusTitle"));
        statusDetail.setObjectName(QStringLiteral("StatusDetail"));
        statusDetail.setTextFormat(Qt::PlainText);
        statusDetail.setTextInteractionFlags(Qt::TextSelectableByMouse);
        statusDetail.setSizePolicy(QSizePolicy::Ignored, QSizePolicy::Preferred);
        statusDetail.setWordWrap(true);
        progress.setAccessibleName(QStringLiteral("转换进度"));
        progress.setTextVisible(false);
        progress.setFixedHeight(6);
        statusLayout->addWidget(&statusTitle);
        statusLayout->addWidget(&statusDetail);
        statusLayout->addWidget(&progress);

        auto *actions = new QHBoxLayout(&resultActions);
        actions->setContentsMargins(0, 0, 0, 0);
        actions->setSpacing(7);
        actions->addWidget(&openDocument);
        actions->addWidget(&openFolder);
        actions->addStretch();
        statusLayout->addWidget(&resultActions);
        feedbackLayout->addLayout(statusLayout, 1);
        convert.setObjectName(QStringLiteral("Convert"));
        convert.setMinimumWidth(120);
        feedbackLayout->addWidget(&convert, 0, Qt::AlignTop);
        layout->addWidget(feedback);
        layout->addStretch();

        setTabOrder(&browseInput, &input);
        setTabOrder(&input, &useDefault);
        setTabOrder(&useDefault, &output);
        setTabOrder(&output, &browseOutput);
        setTabOrder(&browseOutput, &preserveLayout);
        setTabOrder(&editable, &convert);
        setTabOrder(&convert, &openDocument);
        setTabOrder(&openDocument, &openFolder);

        auto *openShortcut = new QShortcut(QKeySequence::Open, this);
        connect(openShortcut, &QShortcut::activated, &browseInput, &QPushButton::click);
        connect(&browseInput, &QPushButton::clicked, this, [this] { selectInput(); });
        connect(&input, &QLineEdit::textChanged, this, [this] { refreshSettings(); });
        connect(&output, &QLineEdit::textEdited, this, [this] {
            customOutput = !pathFromField(output.text()).isEmpty();
            refreshSettings(false);
        });
        connect(&output, &QLineEdit::editingFinished, this, [this] {
            if (!customOutput && pathFromField(output.text()).isEmpty())
                updateDefaultOutput();
        });
        connect(&browseOutput, &QPushButton::clicked, this, [this] { selectOutput(); });
        connect(&useDefault, &QPushButton::clicked, this, [this] {
            customOutput = false;
            refreshSettings();
        });
        connect(&preserveLayout, &QRadioButton::toggled, this, [this] { refreshSettings(false); });
        connect(&convert, &QPushButton::clicked, this, [this] { startConversion(); });
        connect(&watcher, &QFutureWatcher<QString>::finished, this, [this] { finishConversion(); });
        connect(&openDocument, &QPushButton::clicked, this, [this] {
            if (!QDesktopServices::openUrl(QUrl::fromLocalFile(resultPath)))
                setFeedback(QStringLiteral("无法打开文档。请安装可打开 .docx 的程序，例如 WPS 或 LibreOffice"),
                            QDir::toNativeSeparators(resultPath));
        });
        connect(&openFolder, &QPushButton::clicked, this, [this] {
            const QString folder = QFileInfo(resultPath).absolutePath();
            if (!QDesktopServices::openUrl(QUrl::fromLocalFile(folder)))
                setFeedback(QStringLiteral("无法打开所在文件夹"), QDir::toNativeSeparators(folder));
        });

        refreshSettings();
        resize(700, sizeHint().height());
    }

    ~MainWindow() override { watcher.waitForFinished(); }

protected:
    void dragEnterEvent(QDragEnterEvent *event) override
    {
        if (busy || !event->possibleActions().testFlag(Qt::CopyAction) || droppedPdf(*event->mimeData()).isEmpty())
            return;
        setDropHighlight(true);
        event->setDropAction(Qt::CopyAction);
        event->accept();
    }

    void dragMoveEvent(QDragMoveEvent *event) override
    {
        if (busy || !event->possibleActions().testFlag(Qt::CopyAction)) {
            event->ignore();
            return;
        }
        event->setDropAction(Qt::CopyAction);
        event->accept();
    }

    void dragLeaveEvent(QDragLeaveEvent *event) override
    {
        setDropHighlight(false);
        event->accept();
    }

    void dropEvent(QDropEvent *event) override
    {
        setDropHighlight(false);
        const QString path = droppedPdf(*event->mimeData());
        if (busy || !event->possibleActions().testFlag(Qt::CopyAction) || path.isEmpty() || !QFileInfo(path).isFile())
            return;
        input.setText(QDir::toNativeSeparators(path));
        event->setDropAction(Qt::CopyAction);
        event->accept();
    }

    void closeEvent(QCloseEvent *event) override
    {
        if (busy) {
            statusDetail.setText(QStringLiteral("转换完成后即可关闭窗口。"));
            statusDetail.show();
            event->ignore();
            return;
        }
        QWidget::closeEvent(event);
    }

private:
    ConversionMode currentMode() const
    {
        return preserveLayout.isChecked() ? ConversionMode::PreserveLayout : ConversionMode::Editable;
    }

    void setDropHighlight(bool active)
    {
        dropArea.setProperty("dragging", active);
        dropArea.style()->unpolish(&dropArea);
        dropArea.style()->polish(&dropArea);
        dropArea.update();
    }

    void selectInput()
    {
        const QString path = QFileDialog::getOpenFileName(
            this, QStringLiteral("选择 PDF 文件"), pathFromField(input.text()), QStringLiteral("PDF 文件 (*.pdf)"));
        if (!path.isEmpty())
            input.setText(QDir::toNativeSeparators(path));
    }

    void selectOutput()
    {
        QFileDialog dialog(this, QStringLiteral("保存 Word 文档"), {}, QStringLiteral("Word 文档 (*.docx)"));
        dialog.setAcceptMode(QFileDialog::AcceptSave);
        dialog.setDefaultSuffix(QStringLiteral("docx"));
        dialog.setOption(QFileDialog::DontConfirmOverwrite);
        const QString path = pathFromField(output.text());
        if (!path.isEmpty())
            dialog.selectFile(path);
        if (dialog.exec() == QDialog::Accepted) {
            customOutput = true;
            output.setText(QDir::toNativeSeparators(dialog.selectedFiles().first()));
            refreshSettings(false);
        }
    }

    void updateDefaultOutput()
    {
        const QString source = pathFromField(input.text());
        output.setText(isPdfPath(source)
            ? QDir::toNativeSeparators(defaultOutput(QFileInfo(source).absoluteFilePath()))
            : QString{});
    }

    void setFeedback(const QString &title, const QString &detail = {})
    {
        statusTitle.setText(title);
        statusDetail.setText(detail);
        statusDetail.setVisible(!detail.isEmpty());
    }

    void refreshSettings(bool updateOutput = true)
    {
        if (updateOutput && !customOutput)
            updateDefaultOutput();

        resultPath.clear();
        resultActions.hide();
        progress.hide();
        hint.setText(currentMode() == ConversionMode::PreserveLayout
            ? QStringLiteral("以整页图片保留分页和方向，适用于扫描版 PDF。无需安装 Microsoft Word。")
            : QStringLiteral("直接从 PDF 提取文字生成可编辑文档，无需安装 Microsoft Word。"
                             "会自动缩小字号以适应页面；只有无文字层的扫描页才转为图片，排版可能变化。"));

        const QString source = pathFromField(input.text());
        if (source.isEmpty())
            setFeedback(QStringLiteral("等待选择 PDF"));
        else if (!isPdfPath(source))
            setFeedback(QStringLiteral("请选择 .pdf 文件"));
        else
            setFeedback(QStringLiteral("等待转换"));
        updateControls();
    }

    void updateControls()
    {
        const bool ready = !busy;
        input.setEnabled(ready);
        output.setEnabled(ready);
        dropArea.setEnabled(ready);
        browseOutput.setEnabled(ready);
        useDefault.setEnabled(ready && customOutput);
        preserveLayout.setEnabled(ready);
        editable.setEnabled(ready);
        convert.setEnabled(ready && isPdfPath(pathFromField(input.text())));
        convert.setText(busy ? QStringLiteral("转换中…") : QStringLiteral("开始转换"));
    }

    void startConversion()
    {
        if (busy)
            return;

        const QFileInfo source(pathFromField(input.text()));
        if (!isPdfPath(source.filePath()) || !source.isFile() || !source.isReadable()) {
            setFeedback(QStringLiteral("无法读取源 PDF"), QStringLiteral("请检查文件路径和读取权限。"));
            input.setFocus();
            return;
        }

        QString path = pathFromField(output.text());
        if (path.isEmpty())
            path = defaultOutput(source.absoluteFilePath());
        if (QFileInfo(path).isDir() || path.endsWith(u'/')) {
            setFeedback(QStringLiteral("请填写保存文件名"), QStringLiteral("保存位置需要包含 .docx 文件名。"));
            output.setFocus();
            return;
        }
        if (QFileInfo(path).suffix().isEmpty())
            path += QStringLiteral(".docx");
        const QFileInfo destination(path);
        if (destination.suffix().compare(QStringLiteral("docx"), Qt::CaseInsensitive) != 0
            || !destination.absoluteDir().exists()) {
            setFeedback(QStringLiteral("请检查保存位置"), QStringLiteral("请选择已有文件夹并使用 .docx 扩展名。"));
            output.setFocus();
            return;
        }

        const bool replaceExisting = destination.exists() || destination.isSymLink();
        if (replaceExisting) {
            QMessageBox confirm(QMessageBox::Question, QStringLiteral("覆盖文件"),
                                QStringLiteral("转换完成后覆盖现有文件吗？"),
                                QMessageBox::Yes | QMessageBox::No, this);
            confirm.setTextFormat(Qt::PlainText);
            confirm.setInformativeText(QDir::toNativeSeparators(destination.absoluteFilePath()));
            confirm.setDefaultButton(QMessageBox::No);
            confirm.button(QMessageBox::Yes)->setText(QStringLiteral("覆盖"));
            confirm.button(QMessageBox::No)->setText(QStringLiteral("返回"));
            if (confirm.exec() != QMessageBox::Yes)
                return;
        }

        const QString sourcePath = source.absoluteFilePath();
        const QString destinationPath = destination.absoluteFilePath();
        const auto mode = currentMode();
        output.setText(QDir::toNativeSeparators(destinationPath));
        resultPath = destinationPath;
        resultActions.hide();
        busy = true;
        updateControls();
        progress.setRange(0, 0);
        progress.show();
        setFeedback(mode == ConversionMode::PreserveLayout
            ? QStringLiteral("正在读取 PDF")
            : QStringLiteral("正在提取 PDF 文字"));
        watcher.setFuture(QtConcurrent::run([this, sourcePath, destinationPath, mode, replaceExisting] {
            return convertPdf(sourcePath, destinationPath, mode, [this](int completed, int total) {
                QMetaObject::invokeMethod(this, [this, completed, total] {
                    if (!busy)
                        return;
                    progress.setRange(0, total);
                    progress.setValue(completed);
                    setFeedback(completed == total ? QStringLiteral("正在保存 Word 文档")
                                                   : QStringLiteral("正在处理页面"),
                                QStringLiteral("已处理 %1 / %2 页").arg(completed).arg(total));
                }, Qt::QueuedConnection);
            }, replaceExisting);
        }));
    }

    void finishConversion()
    {
        const QString error = watcher.result();
        busy = false;
        progress.hide();
        updateControls();
        if (!error.isEmpty()) {
            resultPath.clear();
            setFeedback(QStringLiteral("转换失败"), error);
            return;
        }

        setFeedback(QStringLiteral("转换完成"), QDir::toNativeSeparators(resultPath));
        resultActions.show();
    }

    QFrame dropArea;
    QLineEdit input;
    QLineEdit output;
    QPushButton browseInput{QStringLiteral("选择文件…")};
    QPushButton browseOutput{QStringLiteral("选择位置…")};
    QPushButton useDefault{QStringLiteral("恢复默认")};
    QRadioButton preserveLayout{QStringLiteral("保留原版式")};
    QRadioButton editable{QStringLiteral("可编辑文本")};
    QLabel hint;
    QLabel statusTitle;
    QLabel statusDetail;
    QProgressBar progress;
    QWidget resultActions;
    QPushButton openDocument{QStringLiteral("打开文档")};
    QPushButton openFolder{QStringLiteral("打开文件夹")};
    QPushButton convert{QStringLiteral("开始转换")};
    QFutureWatcher<QString> watcher;
    QString resultPath;
    bool customOutput = false;
    bool busy = false;
};

} // namespace

int main(int argc, char **argv)
{
    QCommandLineParser parser;
    parser.setApplicationDescription(QStringLiteral("将 PDF 转换为 Word DOCX"));
    parser.addHelpOption();
    const QCommandLineOption editable(QStringLiteral("editable"),
                                      QStringLiteral("从 PDF 提取文字生成可编辑文本，无需 Microsoft Word。"));
    parser.addOption(editable);
    parser.addPositionalArgument(QStringLiteral("pdf"), QStringLiteral("待转换的 PDF 文件"),
                                 QStringLiteral("[pdf]"));
    parser.addPositionalArgument(QStringLiteral("docx"), QStringLiteral("输出 Word 文件"),
                                 QStringLiteral("[docx]"));

    if (argc > 1) {
        QCoreApplication app(argc, argv);
        parser.process(app);
        const auto paths = parser.positionalArguments();
        if (paths.isEmpty() || paths.size() > 2)
            parser.showHelp(1);
        const auto destination = paths.size() == 2 ? paths[1] : defaultOutput(paths[0]);
        const auto mode = parser.isSet(editable) ? ConversionMode::Editable : ConversionMode::PreserveLayout;
        const auto error = convertPdf(paths[0], destination, mode);
        if (!error.isEmpty()) {
            QTextStream(stderr) << error << Qt::endl;
            return 1;
        }
        QTextStream(stdout) << QDir::toNativeSeparators(destination) << Qt::endl;
        return 0;
    }

    QApplication app(argc, argv);
    app.setFont(QFont(QStringLiteral("Microsoft YaHei UI"), 10));
    MainWindow window;
    window.show();
    return app.exec();
}
