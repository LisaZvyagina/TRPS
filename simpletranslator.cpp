#include "simpletranslator.h"

#include <QNetworkRequest>
#include <QNetworkReply>
#include <QUrl>
#include <QTimer>
#include <QJsonDocument>
#include <QJsonArray>

SimpleTranslator::SimpleTranslator(QObject *parent)
    : QObject(parent),
    m_networkManager(new QNetworkAccessManager(this)),
    m_currentChunkIndex(0),
    m_totalChunks(0)
{
    connect(m_networkManager, &QNetworkAccessManager::finished,
            this, &SimpleTranslator::onReplyFinished);
}
void SimpleTranslator::translate(const QString &text,
                                 const QString &sourceLang,
                                 const QString &targetLang)
{
    m_sourceLang = sourceLang;
    m_targetLang = targetLang;

    m_chunks.clear();
    m_translatedChunks.clear();

    // 🔥 УСКОРЕНИЕ
    const int CHUNK_SIZE = 1500;

    for (int i = 0; i < text.length(); i += CHUNK_SIZE)
        m_chunks << text.mid(i, CHUNK_SIZE);

    m_currentChunkIndex = 0;
    m_totalChunks = m_chunks.size();

    emit progressUpdated(5);

    QTimer::singleShot(0, this, &SimpleTranslator::translateNextChunk);
}
void SimpleTranslator::translateNextChunk()
{
    if (m_currentChunkIndex >= m_totalChunks) {

        QString result;

        for (const QString &chunk : m_translatedChunks) {

            if (chunk == "__LINE__" || chunk == "__EMPTY__")
                result += "\n";
            else
                result += chunk;
        }

        emit translationFinished(result, true, "");
        return;
    }

    QString chunk = m_chunks[m_currentChunkIndex];

    if (chunk.startsWith("__")) {
        m_translatedChunks << chunk;
        m_currentChunkIndex++;
        QTimer::singleShot(0, this, &SimpleTranslator::translateNextChunk);
        return;
    }

    QString encoded = QUrl::toPercentEncoding(chunk);

    QUrl url("https://translate.googleapis.com/translate_a/single?client=gtx&sl="
             + m_sourceLang + "&tl=" + m_targetLang + "&dt=t&q=" + encoded);

    QNetworkRequest request(url);
    request.setHeader(QNetworkRequest::UserAgentHeader, "Mozilla/5.0");

    m_networkManager->get(request);
}

void SimpleTranslator::onReplyFinished(QNetworkReply *reply)
{
    if (reply->error()) {
        reply->deleteLater();

        QTimer::singleShot(1000, this, &SimpleTranslator::translateNextChunk);
        return;
    }

    QJsonDocument doc = QJsonDocument::fromJson(reply->readAll());

    QString result;

    if (doc.isArray()) {
        QJsonArray root = doc.array();

        if (!root.isEmpty() && root[0].isArray()) {
            QJsonArray arr = root[0].toArray();

            for (const auto &v : arr) {
                QJsonArray part = v.toArray();
                if (!part.isEmpty())
                    result += part[0].toString();
            }
        }
    }

    if (result.isEmpty()) {
        result = m_chunks[m_currentChunkIndex];
    }

    m_translatedChunks << result;
    m_currentChunkIndex++;

    emit progressUpdated(5 + (m_currentChunkIndex * 90 / m_totalChunks));

    QTimer::singleShot(100, this, &SimpleTranslator::translateNextChunk);
}