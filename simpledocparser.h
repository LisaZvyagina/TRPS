#ifndef SIMPLEDOCPARSER_H
#define SIMPLEDOCPARSER_H

#include <QObject>
#include <QString>
#include <QList>

class SimpleDocParser : public QObject
{
    Q_OBJECT

public:
    explicit SimpleDocParser(QObject *parent = nullptr);

    bool parse(const QString &filePath, QString &outText, QString &outFormatInfo);
    bool build(const QString &translatedText, const QString &outputPath);

private:
    struct TextRun {
        QString text;
        bool bold = false;
        bool italic = false;
        bool underline = false;
        bool isHeading = false;
        int headingLevel = 0;
        bool isTable = false;
        int tableCols = 0;
        QString fontName;
        int fontSize = 0;
    };

    QList<TextRun> m_runs;

    bool parseTXT(const QString &filePath, QString &outText);
    bool parseDOCX(const QString &filePath, QString &outText);
    bool parsePDF(const QString &filePath, QString &outText);

    bool buildTXT(const QString &translatedText, const QString &outputPath);
    bool buildDOCX(const QString &translatedText, const QString &outputPath);
    bool buildPDF(const QString &translatedText, const QString &outputPath);
};

#endif
