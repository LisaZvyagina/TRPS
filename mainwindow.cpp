#include "mainwindow.h"
#include "simpledocparser.h"
#include "simpletranslator.h"

#include <QTime>
#include <QFileInfo>
#include <QFileDialog>
#include <QRegularExpression>
#include <QMessageBox>
#include <QFont>
#include <QVBoxLayout>
#include <QHBoxLayout>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent),
    m_parser(new SimpleDocParser(this)),
    m_translator(new SimpleTranslator(this)),
    m_isTranslating(false)
{
    setupUI();
    setAcceptDrops(true);
    appendLog("Приложение запущено.");

    connect(m_translator, &SimpleTranslator::translationFinished,
            this, &MainWindow::onTranslationFinished);

    connect(m_translator, &SimpleTranslator::progressUpdated,
            this, &MainWindow::onProgressUpdated);
}

MainWindow::~MainWindow()
{
}

void MainWindow::setupUI()
{
    setWindowTitle("DocTrans");
    setMinimumSize(800, 600);

    m_centralWidget = new QWidget(this);
    m_mainLayout = new QVBoxLayout(m_centralWidget);

    m_topLayout = new QHBoxLayout();

    m_selectFileBtn = new QPushButton("Выбрать файл");
    m_fileLabel = new QLabel("Файл не выбран");

    m_sourceLangCombo = new QComboBox();
    m_sourceLangCombo->addItems({"ru", "en"});

    m_targetLangCombo = new QComboBox();
    m_targetLangCombo->addItems({"en", "ru"});

    m_translateBtn = new QPushButton("Перевести");
    m_translateBtn->setEnabled(false);

    m_topLayout->addWidget(m_selectFileBtn);
    m_topLayout->addWidget(m_fileLabel);
    m_topLayout->addWidget(m_sourceLangCombo);
    m_topLayout->addWidget(m_targetLangCombo);
    m_topLayout->addWidget(m_translateBtn);

    m_logTextEdit = new QTextEdit();
    m_logTextEdit->setReadOnly(true);
    m_logTextEdit->setFont(QFont("Consolas", 10));

    m_progressBar = new QProgressBar();
    m_progressBar->setVisible(false);

    m_mainLayout->addLayout(m_topLayout);
    m_mainLayout->addWidget(m_logTextEdit);
    m_mainLayout->addWidget(m_progressBar);

    setCentralWidget(m_centralWidget);

    connect(m_selectFileBtn, &QPushButton::clicked,
            this, &MainWindow::onSelectFileClicked);

    connect(m_translateBtn, &QPushButton::clicked,
            this, &MainWindow::onTranslateClicked);
}

void MainWindow::appendLog(const QString &message)
{
    m_logTextEdit->append(
        QString("[%1] %2")
            .arg(QTime::currentTime().toString("hh:mm:ss"), message));
}

void MainWindow::onSelectFileClicked()
{
    QStringList filePaths = QFileDialog::getOpenFileNames(
        this,
        "Выберите файлы (можно несколько)",
        "",
        "Files (*.txt *.docx *.pdf)");

    if (filePaths.isEmpty())
        return;

    // Добавляем все выбранные файлы в очередь
    for (const QString &filePath : filePaths) {
        m_fileQueue.enqueue(filePath);
        appendLog("Добавлен в очередь: " + QFileInfo(filePath).fileName());
    }

    // Если перевод не идёт, начинаем обработку очереди
    if (!m_isTranslating) {
        processNextFile();
    }
}

void MainWindow::onTranslateClicked()
{
    QString text;
    QString fmt;

    if (!m_parser->parse(m_currentFilePath, text, fmt)) {
        appendLog("Ошибка чтения файла");
        // Переходим к следующему файлу при ошибке
        processNextFile();
        return;
    }

    // Автоопределение языка (оставляем как есть)
    QString lower = text.toLower();
    if (lower.contains(QRegularExpression("[a-z]"))) {
        m_sourceLangCombo->setCurrentText("en");
        m_targetLangCombo->setCurrentText("ru");
    } else {
        m_sourceLangCombo->setCurrentText("ru");
        m_targetLangCombo->setCurrentText("en");
    }

    appendLog("Начат перевод...");
    m_progressBar->setVisible(true);
    m_progressBar->setValue(10);

    m_translator->translate(text,
                            m_sourceLangCombo->currentText(),
                            m_targetLangCombo->currentText());
}

void MainWindow::onTranslationFinished(const QString &translatedText, bool success, const QString &error)
{
    if (!success) {
        appendLog("Ошибка перевода: " + error);
        m_progressBar->setVisible(false);
        // Переходим к следующему файлу в очереди
        processNextFile();
        return;
    }

    QFileInfo info(m_currentFilePath);
    QString outputPath;

    if (info.suffix() == "docx") {
        outputPath = info.path() + "/" + info.completeBaseName() + "_translated.docx";
    }
    else if (info.suffix() == "pdf") {
        outputPath = info.path() + "/" + info.completeBaseName() + "_translated.pdf";
    }
    else {
        outputPath = info.path() + "/" + info.completeBaseName() + "_translated.txt";
    }

    if (!m_parser->build(translatedText, outputPath)) {
        appendLog("Ошибка сохранения");
        m_progressBar->setVisible(false);
        processNextFile();
        return;
    }

    appendLog("Готово: " + outputPath);
    m_progressBar->setValue(100);
    m_progressBar->setVisible(false);

    // Показываем уведомление только для последнего файла
    if (m_fileQueue.isEmpty()) {
        QMessageBox::information(this, "Готово", "Файл переведён!");
    } else {
        appendLog("Файл переведён. Переход к следующему...");
    }

    // Запускаем следующий файл
    processNextFile();
}

void MainWindow::onProgressUpdated(int percent)
{
    m_progressBar->setValue(percent);
}

void MainWindow::dragEnterEvent(QDragEnterEvent *event)
{
    if (event->mimeData()->hasUrls()) {
        event->acceptProposedAction();
    }
}

void MainWindow::dropEvent(QDropEvent *event)
{
    QList<QUrl> urls = event->mimeData()->urls();

    if (urls.isEmpty()) {
        appendLog("Drag&Drop: не удалось получить файл");
        return;
    }

    int addedCount = 0;
    for (const QUrl &url : urls) {
        QString filePath = url.toLocalFile();

        if (filePath.endsWith(".docx", Qt::CaseInsensitive) ||
            filePath.endsWith(".pdf", Qt::CaseInsensitive) ||
            filePath.endsWith(".txt", Qt::CaseInsensitive)) {

            m_fileQueue.enqueue(filePath);
            appendLog("Добавлен в очередь (Drag&Drop): " + QFileInfo(filePath).fileName());
            addedCount++;
        } else {
            appendLog("Пропущен (неподдерживаемый формат): " + QFileInfo(filePath).fileName());
        }
    }

    if (addedCount > 0 && !m_isTranslating) {
        processNextFile();
    }
}

void MainWindow::processNextFile()
{
    if (m_fileQueue.isEmpty()) {
        m_isTranslating = false;
        m_progressBar->setVisible(false);
        appendLog("Все файлы обработаны!");
        QMessageBox::information(this, "Готово", "Все файлы в очереди переведены!");
        return;
    }

    m_isTranslating = true;
    m_currentFilePath = m_fileQueue.dequeue();

    QString fileName = QFileInfo(m_currentFilePath).fileName();
    appendLog(QString("=== Обработка файла %1 (осталось в очереди: %2) ===")
                  .arg(fileName)
                  .arg(m_fileQueue.size()));

    m_fileLabel->setText(fileName);

    // Запускаем перевод
    onTranslateClicked();
}

