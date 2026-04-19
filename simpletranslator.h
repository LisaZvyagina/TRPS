#ifndef SIMPLETRANSLATOR_H
#define SIMPLETRANSLATOR_H

#include <QObject>
#include <QNetworkAccessManager>
#include <QStringList>

class QNetworkReply;

class SimpleTranslator : public QObject
{
    Q_OBJECT

public:
    explicit SimpleTranslator(QObject *parent = nullptr);

    void translate(const QString &text,
                   const QString &sourceLang,
                   const QString &targetLang);

signals:
    void translationFinished(const QString &translatedText, bool success, const QString &error);
    void progressUpdated(int percent);

private slots:
    void onReplyFinished(QNetworkReply *reply);

private:
    void translateNextChunk();

    QNetworkAccessManager *m_networkManager;

    QStringList m_chunks;
    QStringList m_translatedChunks;

    int m_currentChunkIndex;
    int m_totalChunks;

    QString m_sourceLang;
    QString m_targetLang;
};

#endif