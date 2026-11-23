#include "converter.h"

#include <QApplication>
#include <QCloseEvent>
#include <QCommandLineParser>
#include <QDesktopServices>
#include <QDir>
#include <QFileDialog>
#include <QFileInfo>
#include <QFrame>
#include <QFutureWatcher>
#include <QGridLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QMessageBox>
#include <QPainter>
#include <QProgressBar>
#include <QPushButton>
#include <QRadioButton>
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

constexpr auto windowStyle = R"qss(
QWidget#ConverterWindow { background: #eff3f9; color: #182338; }
QFrame#Header {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #15233f, stop:1 #253c72);
    border-radius: 16px;
}
QLabel#Title { color: #f7faff; font-size: 29px; font-weight: 700; }
QLabel#Subtitle { color: #c9d8f4; font-size: 13px; }
QFrame#Form {
    background: #ffffff;
    border: 1px solid #dce4f0;
    border-radius: 14px;
}
QLabel[fieldLabel="true"] { color: #23324d; font-weight: 600; }
QLabel#ModeHint, QLabel#Status { color: #4d5c75; font-size: 13px; }
QLineEdit, QPushButton, QRadioButton {
    font-size: 14px;
    border: 1px solid #cbd6e7;
    border-radius: 8px;
    color: #1d2d48;
}
QLineEdit { background: #f8fafd; padding: 9px 12px; selection-background-color: #315cdd; }
QLineEdit:focus, QPushButton:focus, QRadioButton:focus { border-color: #315cdd; }
QLineEdit:disabled { color: #596983; }
QPushButton { background: #f8fafd; padding: 9px 15px; font-weight: 600; }
QPushButton:hover { background: #ecf2ff; border-color: #9db5e9; }
QPushButton:pressed { background: #e2ebff; }
QRadioButton { background: #f8fafd; padding: 11px 12px; spacing: 9px; }
QRadioButton:hover { background: #f1f5ff; }
QRadioButton:checked { background: #edf3ff; border-color: #6490ef; color: #2047a8; }
QRadioButton:focus { border-color: #254bb6; }
QRadioButton::indicator { width: 8px; height: 8px; border: 4px solid #d6e0ef; border-radius: 8px; background: #ffffff; }
QRadioButton::indicator:checked { border-color: #315cdd; }
QPushButton:disabled, QRadioButton:disabled { color: #66748c; }
QPushButton#Convert {
    background: #2854d2; border-color: #2854d2; color: #ffffff;
    padding: 11px 23px; font-size: 15px;
}
QPushButton#Convert:hover { background: #1e46bd; }
QPushButton#Convert:pressed { background: #173a9f; }
QPushButton#Convert:focus { border-color: #122968; }
QPushButton#Convert:disabled { background: #d7e2fb; border-color: #d7e2fb; color: #304d94; }
QProgressBar { background: #dce5f3; border: none; border-radius: 3px; }
QProgressBar::chunk { background: #315cdd; border-radius: 3px; }
)qss";

class DocumentGraphic final : public QWidget
{
public:
    explicit DocumentGraphic(QWidget *parent) : QWidget(parent) { setFixedSize(120, 72); }

protected:
    void paintEvent(QPaintEvent *) override
    {
        QPainter painter(this);
        painter.setRenderHint(QPainter::Antialiasing);

        const auto drawDocument = [&painter](int x, int y, const QColor &fill, const QColor &line) {
            painter.setPen(Qt::NoPen);
            painter.setBrush(fill);
            painter.drawRoundedRect(QRectF(x, y, 34, 45), 7, 7);
            painter.setPen(QPen(line, 2, Qt::SolidLine, Qt::RoundCap));
            painter.drawLine(QPointF(x + 9, y + 14), QPointF(x + 20, y + 14));
            painter.drawLine(QPointF(x + 9, y + 22), QPointF(x + 25, y + 22));
            painter.drawLine(QPointF(x + 9, y + 30), QPointF(x + 21, y + 30));
        };
        drawDocument(4, 16, QColor(255, 255, 255, 34), QColor("#c9d9fd"));
        drawDocument(79, 8, QColor("#dce8ff"), QColor("#335dcb"));
        painter.setPen(QPen(QColor("#a9c2ff"), 2, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        painter.drawLine(47, 34, 68, 34);
        painter.drawLine(63, 29, 68, 34);
        painter.drawLine(68, 34, 63, 39);
    }
};

void addFileField(QGridLayout &form, int row, const QString &label, QLineEdit &field,
                  QPushButton &browse)
{
    auto *caption = new QLabel(label);
    caption->setProperty("fieldLabel", true);
    caption->setBuddy(&field);
    form.addWidget(caption, row, 0, 1, 2);
    form.addWidget(&field, row + 1, 0);
    form.addWidget(&browse, row + 1, 1);
}

class MainWindow final : public QWidget
{
public:
    MainWindow()
    {
        setWindowTitle(QStringLiteral("PDF 转 Word"));
        setObjectName(QStringLiteral("ConverterWindow"));
        setMinimumWidth(600);
        setStyleSheet(QString::fromLatin1(windowStyle));

        auto *layout = new QVBoxLayout(this);
        layout->setContentsMargins(26, 24, 26, 22);
        layout->setSpacing(16);

        auto *header = new QFrame(this);
        header->setObjectName(QStringLiteral("Header"));
        auto *headerLayout = new QHBoxLayout(header);
        headerLayout->setContentsMargins(25, 16, 24, 16);
        auto *heading = new QVBoxLayout;
        heading->setSpacing(5);
        auto *title = new QLabel(windowTitle(), header);
        title->setObjectName(QStringLiteral("Title"));
        auto *subtitle = new QLabel(QStringLiteral("保留分页、纸张方向和内容位置"), header);
        subtitle->setObjectName(QStringLiteral("Subtitle"));
        heading->addWidget(title);
        heading->addWidget(subtitle);
        headerLayout->addLayout(heading, 1);
        headerLayout->addWidget(new DocumentGraphic(header));
        layout->addWidget(header);

        auto *form = new QFrame(this);
        form->setObjectName(QStringLiteral("Form"));
        auto *formLayout = new QGridLayout(form);
        formLayout->setContentsMargins(20, 18, 20, 18);
        formLayout->setHorizontalSpacing(10);
        formLayout->setVerticalSpacing(8);
        formLayout->setColumnStretch(0, 1);
        input.setPlaceholderText(QStringLiteral("选择要转换的 PDF 文件"));
        output.setPlaceholderText(QStringLiteral("默认保存在 PDF 所在文件夹"));
        addFileField(*formLayout, 0, QStringLiteral("源 PDF"), input, browseInput);
        formLayout->setRowMinimumHeight(2, 3);
        addFileField(*formLayout, 3, QStringLiteral("输出 Word"), output, browseOutput);
        formLayout->setRowMinimumHeight(5, 3);

        auto *modeLabel = new QLabel(QStringLiteral("转换模式"), form);
        modeLabel->setProperty("fieldLabel", true);
        modeLabel->setBuddy(&preserveLayout);
        formLayout->addWidget(modeLabel, 6, 0, 1, 2);
        auto *modeLayout = new QHBoxLayout;
        modeLayout->setSpacing(10);
        preserveLayout.setChecked(true);
        modeLayout->addWidget(&preserveLayout, 1);
        modeLayout->addWidget(&editable, 1);
        formLayout->addLayout(modeLayout, 7, 0, 1, 2);
        hint.setObjectName(QStringLiteral("ModeHint"));
        hint.setWordWrap(true);
        formLayout->addWidget(&hint, 8, 0, 1, 2);
        layout->addWidget(form);
        layout->addStretch();

        auto *footer = new QHBoxLayout;
        auto *feedback = new QVBoxLayout;
        feedback->setSpacing(8);
        status.setObjectName(QStringLiteral("Status"));
        status.setWordWrap(true);
        progress.setTextVisible(false);
        progress.setFixedHeight(6);
        progress.setRange(0, 1);
        progress.setValue(0);
        feedback->addWidget(&status);
        feedback->addWidget(&progress);
        footer->addLayout(feedback, 1);
        footer->setSpacing(18);
        convert.setObjectName(QStringLiteral("Convert"));
        convert.setMinimumWidth(132);
        footer->addWidget(&convert, 0, Qt::AlignVCenter);
        layout->addLayout(footer);

        setTabOrder(&input, &browseInput);
        setTabOrder(&browseInput, &output);
        setTabOrder(&output, &browseOutput);
        setTabOrder(&browseOutput, &preserveLayout);
        setTabOrder(&editable, &convert);

        connect(&browseInput, &QPushButton::clicked, this, [this] {
            const auto path = QFileDialog::getOpenFileName(this, QStringLiteral("选择 PDF 文件"),
                                                          input.text(), QStringLiteral("PDF 文件 (*.pdf)"));
            if (!path.isEmpty()) {
                input.setText(QDir::toNativeSeparators(path));
                output.setText(QDir::toNativeSeparators(defaultOutput(path)));
            }
        });
        connect(&input, &QLineEdit::editingFinished, this, [this] {
            if (output.text().trimmed().isEmpty() && !input.text().trimmed().isEmpty())
                output.setText(QDir::toNativeSeparators(defaultOutput(input.text().trimmed())));
        });
        connect(&browseOutput, &QPushButton::clicked, this, [this] {
            auto path = QFileDialog::getSaveFileName(this, QStringLiteral("保存 Word 文档"),
                                                    output.text(), QStringLiteral("Word 文档 (*.docx)"),
                                                    nullptr, QFileDialog::DontConfirmOverwrite);
            if (!path.isEmpty()) {
                if (QFileInfo(path).suffix().isEmpty())
                    path += QStringLiteral(".docx");
                output.setText(QDir::toNativeSeparators(path));
            }
        });
        connect(&preserveLayout, &QRadioButton::toggled, this, [this] { updateHint(); });
        connect(&convert, &QPushButton::clicked, this, [this] { startConversion(); });
        connect(&watcher, &QFutureWatcher<QString>::finished, this, [this] { finishConversion(); });
        updateHint();
        resize(708, sizeHint().height());
    }

    ~MainWindow() override { watcher.waitForFinished(); }

protected:
    void closeEvent(QCloseEvent *event) override
    {
        if (busy) {
            status.setText(QStringLiteral("请等待当前转换完成后再关闭窗口。"));
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

    void updateHint()
    {
        hint.setText(currentMode() == ConversionMode::PreserveLayout
            ? QStringLiteral("每页以整页图片写入 Word，适用于文字版和扫描版 PDF。")
            : QStringLiteral("需要桌面版 Microsoft Word 2013 或更新版本。分页和排版可能变化。"));
    }

    void setBusy(bool value)
    {
        busy = value;
        input.setEnabled(!busy);
        output.setEnabled(!busy);
        browseInput.setEnabled(!busy);
        browseOutput.setEnabled(!busy);
        preserveLayout.setEnabled(!busy);
        editable.setEnabled(!busy);
        convert.setEnabled(!busy);
        convert.setText(busy ? QStringLiteral("转换中…") : QStringLiteral("开始转换"));
    }

    void startConversion()
    {
        const QString source = input.text().trimmed();
        QString destination = output.text().trimmed();
        if (source.isEmpty() || !QFileInfo(source).isFile()) {
            QMessageBox::warning(this, windowTitle(), QStringLiteral("请先选择要转换的 PDF 文件。"));
            return;
        }
        if (destination.isEmpty()) {
            destination = defaultOutput(source);
            output.setText(QDir::toNativeSeparators(destination));
        }
        if (QFileInfo::exists(destination)
            && QMessageBox::question(this, QStringLiteral("覆盖文件"),
                                     QStringLiteral("转换完成后覆盖现有的 Word 文件吗？"),
                                     QMessageBox::Yes | QMessageBox::No,
                                     QMessageBox::No) != QMessageBox::Yes) {
            return;
        }

        const auto selectedMode = currentMode();
        resultPath = QFileInfo(destination).absoluteFilePath();
        setBusy(true);
        progress.setRange(0, 0);
        status.setText(QStringLiteral("正在转换…"));
        watcher.setFuture(QtConcurrent::run([this, source, destination, selectedMode] {
            return convertPdf(source, destination, selectedMode, [this](int completed, int total) {
                QMetaObject::invokeMethod(this, [this, completed, total] {
                    if (!busy)
                        return;
                    progress.setRange(0, total);
                    progress.setValue(completed);
                    status.setText(QStringLiteral("已处理 %1 / %2 页").arg(completed).arg(total));
                }, Qt::QueuedConnection);
            });
        }));
    }

    void finishConversion()
    {
        const auto error = watcher.result();
        setBusy(false);
        progress.setRange(0, 1);
        progress.setValue(error.isEmpty() ? 1 : 0);
        if (!error.isEmpty()) {
            status.setText(QStringLiteral("转换失败"));
            QMessageBox::warning(this, QStringLiteral("转换失败"), error);
            return;
        }

        status.setText(QStringLiteral("已保存到 %1").arg(QDir::toNativeSeparators(resultPath)));
        if (QMessageBox::question(this, QStringLiteral("转换完成"),
                                 QStringLiteral("现在打开 Word 文档吗？"),
                                 QMessageBox::Yes | QMessageBox::No,
                                 QMessageBox::No) == QMessageBox::Yes
            && !QDesktopServices::openUrl(QUrl::fromLocalFile(resultPath))) {
            QMessageBox::warning(this, windowTitle(), QStringLiteral("无法打开 Word 文档。"));
        }
    }

    QLineEdit input;
    QLineEdit output;
    QPushButton browseInput{QStringLiteral("选择文件…")};
    QPushButton browseOutput{QStringLiteral("选择位置…")};
    QRadioButton preserveLayout{QStringLiteral("保留原版式")};
    QRadioButton editable{QStringLiteral("可编辑文本")};
    QLabel hint;
    QProgressBar progress;
    QLabel status{QStringLiteral("等待选择文件")};
    QPushButton convert{QStringLiteral("开始转换")};
    QFutureWatcher<QString> watcher;
    QString resultPath;
    bool busy = false;
};

} // namespace

int main(int argc, char **argv)
{
    QCommandLineParser parser;
    parser.setApplicationDescription(QStringLiteral("将 PDF 转换为 Word DOCX"));
    parser.addHelpOption();
    const QCommandLineOption editable(QStringLiteral("editable"),
                                      QStringLiteral("使用 Microsoft Word 生成可编辑文本。"));
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
