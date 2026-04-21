#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTextEdit>
#include <QPushButton>
#include <QComboBox>
#include <QLabel>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QProgressBar>
#include <QFileDialog>
#include <QMessageBox>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QQueue>

class SimpleDocParser;
class SimpleTranslator;

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void onSelectFileClicked();
    void onTranslateClicked();
    void onTranslationFinished(const QString &translatedText, bool success, const QString &error);
    void onProgressUpdated(int percent);

protected:
    void dragEnterEvent(QDragEnterEvent *event) override;
    void dropEvent(QDropEvent *event) override;

private:
    void setupUI();
    void appendLog(const QString &message);
    void processNextFile();  // новая функция для пакетной обработки

    QWidget *m_centralWidget;
    QVBoxLayout *m_mainLayout;

    // Верхняя панель
    QHBoxLayout *m_topLayout;
    QPushButton *m_selectFileBtn;
    QLabel *m_fileLabel;
    QComboBox *m_sourceLangCombo;
    QComboBox *m_targetLangCombo;
    QPushButton *m_translateBtn;

    // Центральная область
    QTextEdit *m_logTextEdit;

    // Нижняя панель
    QHBoxLayout *m_bottomLayout;
    QProgressBar *m_progressBar;

    // Данные
    QString m_currentFilePath;
    SimpleDocParser *m_parser;
    SimpleTranslator *m_translator;

    // Пакетная обработка
    QQueue<QString> m_fileQueue;
    bool m_isTranslating;
};

#endif // MAINWINDOW_H
