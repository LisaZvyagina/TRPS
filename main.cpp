#include <QApplication>
#include "mainwindow.h"

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    // Стили приложения
    app.setStyleSheet("QMainWindow { background-color: #f5f5f5; }"
                      "QPushButton { background-color: #0078d4; color: white; border: none; padding: 8px 16px; border-radius: 4px; }"
                      "QPushButton:hover { background-color: #106ebe; }"
                      "QTextEdit { font-family: monospace; font-size: 10pt; }");

    MainWindow window;
    window.show();

    return app.exec();
}