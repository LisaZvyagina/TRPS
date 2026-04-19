#include "simpledocparser.h"

#include <QFile>
#include <QTextStream>
#include <QProcess>
#include <QPdfWriter>
#include <QPainter>
#include <QFont>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#include <QDebug>

SimpleDocParser::SimpleDocParser(QObject *parent) : QObject(parent) {}

bool SimpleDocParser::parse(const QString &filePath, QString &outText, QString &outFormatInfo)
{
    Q_UNUSED(outFormatInfo);
    m_runs.clear();

    if (filePath.endsWith(".txt", Qt::CaseInsensitive))
        return parseTXT(filePath, outText);

    if (filePath.endsWith(".docx", Qt::CaseInsensitive))
        return parseDOCX(filePath, outText);

    if (filePath.endsWith(".pdf", Qt::CaseInsensitive))
        return parsePDF(filePath, outText);

    return false;
}

bool SimpleDocParser::parseTXT(const QString &filePath, QString &outText)
{
    QFile file(filePath);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
        return false;

    QTextStream in(&file);
    outText = in.readAll();
    return true;
}

bool SimpleDocParser::parseDOCX(const QString &filePath, QString &outText)
{
    m_runs.clear();

    QProcess process;

    QString script =
        "import json\n"
        "import docx\n"
        "import re\n"
        "import sys\n"
        "sys.stdout.reconfigure(encoding='utf-8')\n"
        "from docx.table import Table\n"
        "from docx.text.paragraph import Paragraph\n"
        "\n"
        "def iter_block_items(parent):\n"
        "    for child in parent.element.body:\n"
        "        if child.tag.endswith('p'):\n"
        "            yield Paragraph(child, parent)\n"
        "        elif child.tag.endswith('tbl'):\n"
        "            yield Table(child, parent)\n"
        "\n"
        "def get_heading_level(para):\n"
        "    if not para.style or not para.style.name:\n"
        "        return 0\n"
        "    style_name = para.style.name\n"
        "    if style_name.startswith('Heading'):\n"
        "        match = re.search(r'Heading\\s*(\\d+)', style_name)\n"
        "        if match:\n"
        "            return int(match.group(1))\n"
        "        return 1\n"
        "    return 0\n"
        "\n"
        "doc = docx.Document(r'''" + filePath + "''')\n"
                     "runs_data = []\n"
                     "texts = []\n"
                     "\n"
                     "for block in iter_block_items(doc):\n"
                     "    if isinstance(block, Paragraph):\n"
                     "        para_text = block.text.strip()\n"
                     "        if not para_text:\n"
                     "            texts.append('')\n"
                     "            continue\n"
                     "        \n"
                     "        heading_level = get_heading_level(block)\n"
                     "        is_heading = heading_level > 0\n"
                     "        \n"
                     "        # Каждый run — отдельная строка с форматированием\n"
                     "        para_runs = []\n"
                     "        for run in block.runs:\n"
                     "            text = run.text\n"
                     "            if not text:\n"
                     "                continue\n"
                     "            \n"
                     "            # Определяем форматирование run\n"
                     "            bold = bool(run.bold) if run.bold is not None else False\n"
                     "            italic = bool(run.italic) if run.italic is not None else False\n"
                     "            underline = bool(run.underline) if run.underline is not None else False\n"
                     "            \n"
                     "            run_data = {\n"
                     "                'text': text,\n"
                     "                'bold': bold,\n"
                     "                'italic': italic,\n"
                     "                'underline': underline,\n"
                     "                'headingLevel': heading_level,\n"
                     "                'isHeading': is_heading\n"
                     "            }\n"
                     "            para_runs.append(run_data)\n"
                     "            runs_data.append(run_data)\n"
                     "        \n"
                     "        if para_runs:\n"
                     "            # Собираем текст параграфа из runs\n"
                     "            para_full_text = ''.join([r['text'] for r in para_runs])\n"
                     "            texts.append(para_full_text)\n"
                     "        else:\n"
                     "            # Fallback\n"
                     "            run_data = {\n"
                     "                'text': block.text,\n"
                     "                'bold': False,\n"
                     "                'italic': False,\n"
                     "                'underline': False,\n"
                     "                'headingLevel': heading_level,\n"
                     "                'isHeading': is_heading\n"
                     "            }\n"
                     "            runs_data.append(run_data)\n"
                     "            texts.append(block.text)\n"
                     "            \n"
                     "    elif isinstance(block, Table):\n"
                     "        for row in block.rows:\n"
                     "            row_text = []\n"
                     "            for cell in row.cells:\n"
                     "                cell_text = cell.text.strip()\n"
                     "                if cell_text:\n"
                     "                    row_text.append(cell_text)\n"
                     "                    runs_data.append({\n"
                     "                        'text': cell_text,\n"
                     "                        'bold': False,\n"
                     "                        'italic': False,\n"
                     "                        'underline': False,\n"
                     "                        'headingLevel': 0,\n"
                     "                        'isHeading': False\n"
                     "                    })\n"
                     "            if row_text:\n"
                     "                texts.append(' | '.join(row_text))\n"
                     "\n"
                     "print('###JSON###')\n"
                     "print(json.dumps(runs_data, ensure_ascii=False))\n"
                     "print('###TEXT###')\n"
                     "print('\\n'.join(texts))\n";

    process.start("py", QStringList() << "-3" << "-c" << script);
    process.waitForFinished();

    QString output = QString::fromUtf8(process.readAllStandardOutput());

    int jsonStart = output.indexOf("###JSON###");
    int textStart = output.indexOf("###TEXT###");

    if (jsonStart == -1 || textStart == -1)
        return false;

    QString jsonPart = output.mid(jsonStart + 10, textStart - (jsonStart + 10)).trimmed();
    QString textPart = output.mid(textStart + 10).trimmed();

    outText = textPart;

    QJsonDocument doc = QJsonDocument::fromJson(jsonPart.toUtf8());
    QJsonArray arr = doc.array();

    for (auto v : arr) {
        QJsonObject obj = v.toObject();

        TextRun r;
        r.text = obj["text"].toString();
        r.bold = obj["bold"].toBool();
        r.italic = obj["italic"].toBool();
        r.underline = obj["underline"].toBool();
        r.isHeading = obj["isHeading"].toBool();
        r.headingLevel = obj["headingLevel"].toInt();

        m_runs.append(r);
    }

    return true;
}

bool SimpleDocParser::parsePDF(const QString &filePath, QString &outText)
{
    QProcess process;

    QString program =
        "C:/Users/Denis/Downloads/Release-25.12.0-0/poppler-25.12.0/Library/bin/pdftotext.exe";

    process.start(program, QStringList() << filePath << "-");
    process.waitForFinished();

    outText = QString::fromUtf8(process.readAllStandardOutput());
    return !outText.isEmpty();
}

bool SimpleDocParser::build(const QString &text, const QString &path)
{
    if (path.endsWith(".docx"))
        return buildDOCX(text, path);

    if (path.endsWith(".pdf"))
        return buildPDF(text, path);

    return buildTXT(text, path);
}

bool SimpleDocParser::buildTXT(const QString &text, const QString &path)
{
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text))
        return false;

    QTextStream out(&file);
    out << text;
    return true;
}

bool SimpleDocParser::buildDOCX(const QString &translatedText, const QString &path)
{
    QString normalizedText = translatedText;
    normalizedText.replace("\r\n", "\n");
    normalizedText.replace("\r", "\n");

    while (normalizedText.contains("\n\n")) {
        normalizedText.replace("\n\n", "\n");
    }

    QStringList lines = normalizedText.split("\n", Qt::KeepEmptyParts);

    while (!lines.isEmpty() && lines.first().isEmpty()) {
        lines.removeFirst();
    }
    while (!lines.isEmpty() && lines.last().isEmpty()) {
        lines.removeLast();
    }

    qDebug() << "translatedLines:" << lines.size() << "m_runs:" << m_runs.size();

    QStringList translatedParts;

    if (lines.size() == m_runs.size()) {
        translatedParts = lines;
    } else if (lines.size() > 0 && m_runs.size() > 0) {
        translatedParts = distributeLines(lines, m_runs.size());
    } else {
        translatedParts = lines;
    }

    qDebug() << "translatedParts:" << translatedParts.size();

    QJsonArray arr;

    int minSize = qMin(translatedParts.size(), m_runs.size());

    for (int i = 0; i < minSize; ++i) {
        QString cleanText = translatedParts[i];
        cleanText.replace("\r", " ");
        cleanText.replace("\n", " ");
        cleanText = cleanText.simplified(); // Убирает лишние пробелы

        QJsonObject obj;
        obj["text"] = cleanText;
        obj["bold"] = m_runs[i].bold;
        obj["italic"] = m_runs[i].italic;
        obj["underline"] = m_runs[i].underline;
        obj["heading"] = m_runs[i].isHeading;
        obj["headingLevel"] = m_runs[i].headingLevel;
        arr.append(obj);
    }

    QJsonDocument doc(arr);
    QString json = QString::fromUtf8(doc.toJson(QJsonDocument::Compact));
    json.replace("\\", "\\\\");
    json.replace("'''", " ");

    QString safePath = path;
    safePath.replace("\\", "\\\\");

    QProcess process;

    QString script =
        "import json\n"
        "from docx import Document\n"
        "doc = Document()\n"
        "data = json.loads(r'''" + json + "''')\n"
                 "\n"
                 "for item in data:\n"
                 "    text = item['text']\n"
                 "    bold = item['bold']\n"
                 "    italic = item['italic']\n"
                 "    underline = item['underline']\n"
                 "    isHeading = item['heading']\n"
                 "    headingLevel = item.get('headingLevel', 1)\n"
                 "\n"
                 "    if not text and not isHeading:\n"
                 "        doc.add_paragraph()\n"
                 "        continue\n"
                 "\n"
                 "    if isHeading:\n"
                 "        p = doc.add_heading(level=headingLevel)\n"
                 "    else:\n"
                 "        p = doc.add_paragraph()\n"
                 "\n"
                 "    run = p.add_run(text)\n"
                 "    if bold:\n"
                 "        run.bold = True\n"
                 "    if italic:\n"
                 "        run.italic = True\n"
                 "    if underline:\n"
                 "        run.underline = True\n"
                 "\n"
                 "doc.save(r'''" + safePath + "''')\n"
                     "print('OK')\n";

    process.start("py", QStringList() << "-3" << "-c" << script);
    process.waitForFinished();

    QString error = QString::fromUtf8(process.readAllStandardError());
    if (!error.isEmpty()) {
        qDebug() << "Python error:" << error;
        return false;
    }

    return true;
}

QStringList SimpleDocParser::distributeLines(const QStringList &lines, int targetCount)
{
    QStringList result;

    if (targetCount <= 0) return result;
    if (lines.isEmpty()) return result;

    QStringList normalizedLines;
    for (const QString &line : lines) {
        QString normalized = line;
        normalized.replace("\r\n", " ");
        normalized.replace("\r", " ");
        normalized.replace("\n", " ");
        normalized = normalized.simplified();
        if (!normalized.isEmpty()) {
            normalizedLines.append(normalized);
        }
    }

    // Объединяем все строки в один текст
    QString fullText = normalizedLines.join(" ");

    // Разбиваем на слова
    QStringList words = fullText.split(" ", Qt::SkipEmptyParts);

    if (words.isEmpty()) {
        for (int i = 0; i < targetCount; ++i) {
            result.append("");
        }
        return result;
    }

    // Распределяем слова поровну
    int wordsPerRun = words.size() / targetCount;
    int remainder = words.size() % targetCount;

    int wordIndex = 0;
    for (int i = 0; i < targetCount; ++i) {
        int count = wordsPerRun + (i < remainder ? 1 : 0);

        QStringList runWords;
        for (int j = 0; j < count && wordIndex < words.size(); ++j) {
            runWords.append(words[wordIndex++]);
        }

        result.append(runWords.join(" "));
    }

    return result;
}

bool SimpleDocParser::buildPDF(const QString &text, const QString &path)
{
    QPdfWriter writer(path);
    writer.setPageSize(QPageSize(QPageSize::A4));
    writer.setResolution(300);

    QPainter painter(&writer);

    QFont font("Times New Roman", 12);
    painter.setFont(font);

    QFontMetrics fm(font);

    int left = 60;
    int top = 80;
    int right = writer.width() - 60;
    int bottom = writer.height() - 80;

    int maxWidth = right - left;
    int y = top;
    int lineHeight = fm.height() + 6;

    QStringList paragraphs = text.split("\n");

    for (const QString &paragraph : paragraphs) {

        QString trimmed = paragraph.trimmed();

        if (trimmed.isEmpty()) {
            y += lineHeight;
            continue;
        }

        QString currentLine;
        QStringList words = trimmed.split(" ", Qt::SkipEmptyParts);

        for (const QString &word : words) {

            QString testLine = currentLine.isEmpty()
            ? word
            : currentLine + " " + word;

            if (fm.horizontalAdvance(testLine) > maxWidth) {

                painter.drawText(left, y, currentLine);
                y += lineHeight;

                currentLine = word;

                if (y > bottom) {
                    writer.newPage();
                    painter.setFont(font);
                    y = top;
                }
            }
            else {
                currentLine = testLine;
            }
        }

        if (!currentLine.isEmpty()) {
            painter.drawText(left, y, currentLine);
            y += lineHeight;
        }

        y += lineHeight / 2;

        if (y > bottom) {
            writer.newPage();
            painter.setFont(font);
            y = top;
        }
    }

    painter.end();
    return true;
}
