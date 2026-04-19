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

private:
    void setupUI();
    void appendLog(const QString &message);

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
};

#endif // MAINWINDOW_H