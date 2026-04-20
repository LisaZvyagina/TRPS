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
        "def get_font_info(run, para):\n"
        "    font_name = None\n"
        "    font_size_pt = None\n"
        "    \n"
        "    if run.font:\n"
        "        if run.font.name:\n"
        "            font_name = run.font.name\n"
        "        if run.font.size:\n"
        "            font_size_pt = run.font.size / 12700\n"
        "    \n"
        "    if not font_name and para.style and para.style.font:\n"
        "        if para.style.font.name:\n"
        "            font_name = para.style.font.name\n"
        "    \n"
        "    if not font_size_pt and para.style and para.style.font and para.style.font.size:\n"
        "        font_size_pt = para.style.font.size / 12700\n"
        "    \n"
        "    try:\n"
        "        rPr = run._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')\n"
        "        if rPr is not None:\n"
        "            rFonts = rPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')\n"
        "            if rFonts is not None:\n"
        "                ascii_font = rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')\n"
        "                if ascii_font and not font_name:\n"
        "                    font_name = ascii_font\n"
        "            \n"
        "            sz = rPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')\n"
        "            if sz is not None:\n"
        "                val = sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')\n"
        "                if val:\n"
        "                    font_size_pt = int(val) / 2\n"
        "    except:\n"
        "        pass\n"
        "    \n"
        "    return font_name, font_size_pt\n"
        "\n"
        "doc = docx.Document(r'''" + filePath + "''')\n"
                     "blocks_data = []\n"
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
                     "        para_runs = []\n"
                     "        para_font_name = None\n"
                     "        para_font_size = None\n"
                     "        \n"
                     "        for run in block.runs:\n"
                     "            text = run.text\n"
                     "            if not text:\n"
                     "                continue\n"
                     "            \n"
                     "            bold = bool(run.bold) if run.bold is not None else False\n"
                     "            italic = bool(run.italic) if run.italic is not None else False\n"
                     "            underline = bool(run.underline) if run.underline is not None else False\n"
                     "            \n"
                     "            font_name, font_size = get_font_info(run, block)\n"
                     "            if font_name:\n"
                     "                para_font_name = font_name\n"
                     "            if font_size:\n"
                     "                para_font_size = font_size\n"
                     "            \n"
                     "            para_runs.append({\n"
                     "                'text': text,\n"
                     "                'bold': bold,\n"
                     "                'italic': italic,\n"
                     "                'underline': underline\n"
                     "            })\n"
                     "        \n"
                     "        has_bold = any(r['bold'] for r in para_runs)\n"
                     "        has_italic = any(r['italic'] for r in para_runs)\n"
                     "        has_underline = any(r['underline'] for r in para_runs)\n"
                     "        \n"
                     "        alignment = ''\n"
                     "        if block.alignment is not None:\n"
                     "            alignment_map = {\n"
                     "                0: 'left',\n"
                     "                1: 'center',\n"
                     "                2: 'right',\n"
                     "                3: 'justify'\n"
                     "            }\n"
                     "            alignment = alignment_map.get(block.alignment, '')\n"
                     "        \n"
                     "        block_data = {\n"
                     "            'type': 'paragraph',\n"
                     "            'text': block.text,\n"
                     "            'bold': has_bold,\n"
                     "            'italic': has_italic,\n"
                     "            'underline': has_underline,\n"
                     "            'headingLevel': heading_level,\n"
                     "            'isHeading': is_heading,\n"
                     "            'fontName': para_font_name if para_font_name else '',\n"
                     "            'fontSize': round(para_font_size) if para_font_size else 0,\n"
                     "            'alignment': alignment\n"
                     "        }\n"
                     "        blocks_data.append(block_data)\n"
                     "        texts.append(block.text)\n"
                     "            \n"
                     "    elif isinstance(block, Table):\n"
                     "        table_data = []\n"
                     "        max_cols = 0\n"
                     "        \n"
                     "        for row in block.rows:\n"
                     "            row_data = []\n"
                     "            for cell in row.cells:\n"
                     "                cell_text = cell.text.strip()\n"
                     "                row_data.append(cell_text)\n"
                     "            if row_data:\n"
                     "                table_data.append(row_data)\n"
                     "                max_cols = max(max_cols, len(row_data))\n"
                     "        \n"
                     "        if table_data:\n"
                     "            block_data = {\n"
                     "                'type': 'table',\n"
                     "                'text': json.dumps(table_data, ensure_ascii=False),\n"
                     "                'bold': False,\n"
                     "                'italic': False,\n"
                     "                'underline': False,\n"
                     "                'headingLevel': 0,\n"
                     "                'isHeading': False,\n"
                     "                'tableCols': max_cols,\n"
                     "                'tableRows': len(table_data),\n"
                     "                'fontName': '',\n"
                     "                'fontSize': 0,\n"
                     "                'alignment': ''\n"
                     "            }\n"
                     "            blocks_data.append(block_data)\n"
                     "            \n"
                     "            table_text = []\n"
                     "            for row in table_data:\n"
                     "                table_text.append(' | '.join(row))\n"
                     "            texts.append('\\n'.join(table_text))\n"
                     "\n"
                     "print('###JSON###')\n"
                     "print(json.dumps(blocks_data, ensure_ascii=False))\n"
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
        r.isTable = (obj["type"].toString() == "table");
        r.tableCols = obj["tableCols"].toInt();
        r.fontName = obj["fontName"].toString();
        r.fontSize = obj["fontSize"].toInt();
        r.alignment = obj["alignment"].toString();

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

    QStringList textLines;
    QStringList tableLines;

    for (const QString &line : lines) {
        if (line.contains(" | ")) {
            tableLines.append(line);
        } else {
            textLines.append(line);
        }
    }

    qDebug() << "Text lines:" << textLines.size() << "Table lines:" << tableLines.size();

    int textRunsCount = 0;
    for (const auto &run : m_runs) {
        if (!run.isTable) textRunsCount++;
    }

    qDebug() << "Text runs count:" << textRunsCount;

    QStringList distributedText;
    if (textLines.size() == textRunsCount) {
        distributedText = textLines;
    } else if (textLines.size() > 0 && textRunsCount > 0) {
        QString fullText = textLines.join(" ");
        QStringList words = fullText.split(" ", Qt::SkipEmptyParts);

        int wordsPerRun = words.size() / textRunsCount;
        int remainder = words.size() % textRunsCount;

        int wordIdx = 0;
        for (int i = 0; i < textRunsCount; ++i) {
            int count = wordsPerRun + (i < remainder ? 1 : 0);
            QStringList runWords;
            for (int j = 0; j < count && wordIdx < words.size(); ++j) {
                runWords.append(words[wordIdx++]);
            }
            distributedText.append(runWords.join(" "));
        }
    }

    QJsonArray arr;
    int textIdx = 0;
    int tableIdx = 0;

    for (const auto &run : m_runs) {
        if (run.isTable) {
            QJsonObject obj;
            obj["isTable"] = true;
            obj["tableCols"] = run.tableCols;

            QJsonDocument tableDoc = QJsonDocument::fromJson(run.text.toUtf8());
            QJsonArray origTableData = tableDoc.array();

            QJsonArray translatedTableData;
            int rowsNeeded = origTableData.size();

            for (int r = 0; r < rowsNeeded && tableIdx < tableLines.size(); ++r) {
                QString line = tableLines[tableIdx++];
                QStringList cells = line.split("|", Qt::SkipEmptyParts);

                QJsonArray rowArray;
                for (const QString &cell : cells) {
                    rowArray.append(cell.trimmed());
                }
                translatedTableData.append(rowArray);
            }

            obj["tableData"] = translatedTableData;
            arr.append(obj);
        } else {
            QString cleanText = (textIdx < distributedText.size())
            ? distributedText[textIdx++]
            : run.text;

            cleanText = cleanText.simplified();

            QJsonObject obj;
            obj["text"] = cleanText;
            obj["bold"] = run.bold;
            obj["italic"] = run.italic;
            obj["underline"] = run.underline;
            obj["heading"] = run.isHeading;
            obj["headingLevel"] = run.headingLevel;
            obj["isTable"] = false;
            obj["fontName"] = run.fontName;
            obj["fontSize"] = run.fontSize;
            obj["alignment"] = run.alignment;
            arr.append(obj);
        }
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
        "from docx.shared import Pt\n"
        "from docx.enum.text import WD_ALIGN_PARAGRAPH\n"
        "doc = Document()\n"
        "data = json.loads(r'''" + json + "''')\n"
                 "\n"
                 "for item in data:\n"
                 "    is_table = item.get('isTable', False)\n"
                 "    \n"
                 "    if is_table:\n"
                 "        table_data = item.get('tableData', [])\n"
                 "        cols = item.get('tableCols', 1)\n"
                 "        \n"
                 "        if table_data:\n"
                 "            rows = len(table_data)\n"
                 "            table = doc.add_table(rows=rows, cols=cols)\n"
                 "            table.style = 'Table Grid'\n"
                 "            \n"
                 "            for r in range(rows):\n"
                 "                row_data = table_data[r]\n"
                 "                for c in range(min(len(row_data), cols)):\n"
                 "                    table.rows[r].cells[c].text = row_data[c]\n"
                 "    else:\n"
                 "        text = item['text']\n"
                 "        bold = item['bold']\n"
                 "        italic = item['italic']\n"
                 "        underline = item['underline']\n"
                 "        isHeading = item['heading']\n"
                 "        headingLevel = item.get('headingLevel', 1)\n"
                 "        font_name = item.get('fontName', '')\n"
                 "        font_size = item.get('fontSize', 0)\n"
                 "        alignment = item.get('alignment', '')\n"
                 "        \n"
                 "        if not text and not isHeading:\n"
                 "            doc.add_paragraph()\n"
                 "            continue\n"
                 "        \n"
                 "        if isHeading:\n"
                 "            p = doc.add_heading(level=headingLevel)\n"
                 "        else:\n"
                 "            p = doc.add_paragraph()\n"
                 "        \n"
                 "        align_map = {\n"
                 "            'left': WD_ALIGN_PARAGRAPH.LEFT,\n"
                 "            'center': WD_ALIGN_PARAGRAPH.CENTER,\n"
                 "            'right': WD_ALIGN_PARAGRAPH.RIGHT,\n"
                 "            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY\n"
                 "        }\n"
                 "        if alignment in align_map:\n"
                 "            p.alignment = align_map[alignment]\n"
                 "        \n"
                 "        run = p.add_run(text)\n"
                 "        if bold:\n"
                 "            run.bold = True\n"
                 "        if italic:\n"
                 "            run.italic = True\n"
                 "        if underline:\n"
                 "            run.underline = True\n"
                 "        \n"
                 "        if font_name:\n"
                 "            run.font.name = font_name\n"
                 "        if font_size > 0:\n"
                 "            run.font.size = Pt(font_size)\n"
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
