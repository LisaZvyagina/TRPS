#ifndef SIMPLEDOCPARSER_H
#define SIMPLEDOCPARSER_H

#include <QObject>
#include <QString>

class SimpleDocParser : public QObject
{
    Q_OBJECT

public:
    explicit SimpleDocParser(QObject *parent = nullptr);

    bool parse(const QString &filePath, QString &outText, QString &outFormatInfo);
    bool build(const QString &translatedText, const QString &outputPath);

private:
    bool parseTXT(const QString &filePath, QString &outText);
    bool parseDOCX(const QString &filePath, QString &outText);
    bool parsePDF(const QString &filePath, QString &outText);

    bool buildTXT(const QString &translatedText, const QString &outputPath);
    bool buildDOCX(const QString &translatedText, const QString &outputPath);
    bool buildPDF(const QString &translatedText, const QString &outputPath);
};

#endif