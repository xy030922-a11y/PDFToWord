#include "converter.h"

#include <QApplication>
#include <QCloseEvent>
#include <QComboBox>
#include <QCommandLineParser>
#include <QDesktopServices>
#include <QDir>
#include <QFileDialog>
#include <QFileInfo>
#include <QFutureWatcher>
#include <QGridLayout>
#include <QLabel>
#include <QLineEdit>
#include <QMessageBox>
#include <QProgressBar>
#include <QPushButton>
#include <QTextStream>
#include <QUrl>
#include <QWidget>
#include <QtConcurrentRun>

namespace {

QString defaultOutput(const QString &source)
{
    const QFileInfo info(source);
    return info.dir().filePath(info.completeBaseName() + QStringLiteral(".docx"));
}

class MainWindow final : public QWidget
{
public:
    MainWindow()
    {
        setWindowTitle(QStringLiteral("PDF 转 Word"));
        setMinimumWidth(560);

        auto *layout = new QGridLayout(this);
        layout->setContentsMargins(24, 24, 24, 24);
        layout->setSpacing(12);
        layout->setColumnStretch(1, 1);
        layout->setRowStretch(7, 1);

        auto *title = new QLabel(QStringLiteral("PDF 转 Word"), this);
        QFont titleFont = title->font();
        titleFont.setPointSize(18);
        titleFont.setBold(true);
        title->setFont(titleFont);
        layout->addWidget(title, 0, 0, 1, 3);

        auto *inputLabel = new QLabel(QStringLiteral("PDF 文件"), this);
        inputLabel->setBuddy(&input);
        layout->addWidget(inputLabel, 1, 0);
        layout->addWidget(&input, 1, 1);
        layout->addWidget(&browseInput, 1, 2);

        auto *outputLabel = new QLabel(QStringLiteral("保存为"), this);
        outputLabel->setBuddy(&output);
        layout->addWidget(outputLabel, 2, 0);
        layout->addWidget(&output, 2, 1);
        layout->addWidget(&browseOutput, 2, 2);

        auto *modeLabel = new QLabel(QStringLiteral("转换模式"), this);
        modeLabel->setBuddy(&mode);
        layout->addWidget(modeLabel, 3, 0);
        mode.addItem(QStringLiteral("保留原版式（整页图片）"), int(ConversionMode::PreserveLayout));
        mode.addItem(QStringLiteral("可编辑文本（Microsoft Word）"), int(ConversionMode::Editable));
        layout->addWidget(&mode, 3, 1, 1, 2);

        hint.setWordWrap(true);
        layout->addWidget(&hint, 4, 0, 1, 3);
        progress.setTextVisible(false);
        progress.setRange(0, 1);
        progress.setValue(0);
        layout->addWidget(&progress, 5, 0, 1, 3);
        status.setWordWrap(true);
        layout->addWidget(&status, 6, 0, 1, 3);
        layout->addWidget(&convert, 7, 0, 1, 3, Qt::AlignHCenter | Qt::AlignTop);

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
        connect(&mode, &QComboBox::currentIndexChanged, this, [this] { updateHint(); });
        connect(&convert, &QPushButton::clicked, this, [this] { startConversion(); });
        connect(&watcher, &QFutureWatcher<QString>::finished, this, [this] { finishConversion(); });
        updateHint();
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
    ConversionMode currentMode() const { return ConversionMode(mode.currentData().toInt()); }

    void updateHint()
    {
        hint.setText(currentMode() == ConversionMode::PreserveLayout
            ? QStringLiteral("每个 PDF 页面对应一个 Word 页面，适用于扫描版和文字版。")
            : QStringLiteral("需要桌面版 Microsoft Word 2013 或更新版本。分页和排版可能变化。"));
    }

    void setBusy(bool value)
    {
        busy = value;
        input.setEnabled(!busy);
        output.setEnabled(!busy);
        browseInput.setEnabled(!busy);
        browseOutput.setEnabled(!busy);
        mode.setEnabled(!busy);
        convert.setEnabled(!busy);
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
    QPushButton browseInput{QStringLiteral("浏览…")};
    QPushButton browseOutput{QStringLiteral("浏览…")};
    QComboBox mode;
    QLabel hint;
    QProgressBar progress;
    QLabel status{QStringLiteral("请选择一个 PDF 文件")};
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
    MainWindow window;
    window.show();
    return app.exec();
}
