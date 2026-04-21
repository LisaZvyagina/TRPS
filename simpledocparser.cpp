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
        "def get_cell_info(cell):\n"
        "    font_name = None\n"
        "    font_size_pt = None\n"
        "    bold = False\n"
        "    italic = False\n"
        "    underline = False\n"
        "    alignment = ''\n"
        "    \n"
        "    for para in cell.paragraphs:\n"
        "        if not alignment and para.alignment is not None:\n"
        "            alignment_map = {\n"
        "                0: 'left',\n"
        "                1: 'center',\n"
        "                2: 'right',\n"
        "                3: 'justify'\n"
        "            }\n"
        "            alignment = alignment_map.get(para.alignment, '')\n"
        "        \n"
        "        for run in para.runs:\n"
        "            if run.text.strip():\n"
        "                fn, fs = get_font_info(run, para)\n"
        "                if fn and not font_name:\n"
        "                    font_name = fn\n"
        "                if fs and not font_size_pt:\n"
        "                    font_size_pt = fs\n"
        "                if run.bold:\n"
        "                    bold = True\n"
        "                if run.italic:\n"
        "                    italic = True\n"
        "                if run.underline:\n"
        "                    underline = True\n"
        "                if font_name and font_size_pt and alignment:\n"
        "                    break\n"
        "        if font_name and font_size_pt and alignment:\n"
        "            break\n"
        "    \n"
        "    return font_name, font_size_pt, bold, italic, underline, alignment\n"
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
                     "            'alignment': alignment,\n"
                     "            'xPos': 0\n"
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
                     "                fn, fs, bold, italic, underline, align = get_cell_info(cell)\n"
                     "                row_data.append({\n"
                     "                    'text': cell_text,\n"
                     "                    'fontName': fn if fn else '',\n"
                     "                    'fontSize': round(fs) if fs else 0,\n"
                     "                    'bold': bold,\n"
                     "                    'italic': italic,\n"
                     "                    'underline': underline,\n"
                     "                    'alignment': align\n"
                     "                })\n"
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
                     "                'alignment': '',\n"
                     "                'xPos': 0\n"
                     "            }\n"
                     "            blocks_data.append(block_data)\n"
                     "            \n"
                     "            table_text = []\n"
                     "            for row in table_data:\n"
                     "                table_text.append(' | '.join([cell['text'] for cell in row]))\n"
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
        r.xPos = obj["xPos"].toDouble();

        m_runs.append(r);
    }

    return true;
}

bool SimpleDocParser::parsePDF(const QString &filePath, QString &outText)
{
    m_runs.clear();

    QProcess process;

    QString script =
        "import json\n"
        "import fitz\n"
        "import sys\n"
        "sys.stdout.reconfigure(encoding='utf-8')\n"
        "\n"
        "def calc_page_margins(blocks, page_width):\n"
        "    \"\"\"Вычисляет реальное левое и правое поле страницы по минимальным x0/x1.\"\"\"\n"
        "    x0_vals = []\n"
        "    x1_vals = []\n"
        "    for b in blocks:\n"
        "        if 'lines' not in b:\n"
        "            continue\n"
        "        for ln in b['lines']:\n"
        "            for sp in ln['spans']:\n"
        "                if sp['text'].strip():\n"
        "                    x0_vals.append(sp['bbox'][0])\n"
        "                    x1_vals.append(sp['bbox'][2])\n"
        "    left_margin = min(x0_vals) if x0_vals else 50\n"
        "    right_margin = page_width - max(x1_vals) if x1_vals else 50\n"
        "    return left_margin, right_margin\n"
        "\n"
        "def detect_alignment(line, page_width, left_margin, right_margin):\n"
        "    x0 = line['bbox'][0]\n"
        "    x1 = line['bbox'][2]\n"
        "    line_width = x1 - x0\n"
        "    # Порог: строка считается начинающейся у левого поля если отступ <= left_margin + 5pt\n"
        "    left_threshold = left_margin + 5\n"
        "    # Правый край страницы с учетом поля\n"
        "    right_edge = page_width - right_margin\n"
        "    page_center = page_width / 2\n"
        "    \n"
        "    # 1. Начинается у левого поля → left (высший приоритет)\n"
        "    if x0 <= left_threshold:\n"
        "        return 'left'\n"
        "    \n"
        "    # 2. Правый край прижат к правому полю, а левый далеко от левого → right\n"
        "    if abs(x1 - right_edge) <= 15 and x0 > left_threshold + 20:\n"
        "        return 'right'\n"
        "    \n"
        "    # 3. Строка симметрична относительно центра страницы → center\n"
        "    line_center = (x0 + x1) / 2\n"
        "    if abs(line_center - page_center) < page_width * 0.06:\n"
        "        return 'center'\n"
        "    \n"
        "    # 4. Сдвинута вправо, но не по центру и не у правого поля\n"
        "    if x0 > left_threshold + 50:\n"
        "        return 'center'\n"
        "    \n"
        "    return 'left'\n"
        "\n"
        "def build_underline_set(page):\n"
        "    \"\"\"Собирает Y-координаты горизонтальных линий (подчеркивания) на странице.\"\"\"\n"
        "    underline_ys = []\n"
        "    try:\n"
        "        for drawing in page.get_drawings():\n"
        "            for item in drawing.get('items', []):\n"
        "                if item[0] == 'l':  # line\n"
        "                    p1, p2 = item[1], item[2]\n"
        "                    if abs(p1.y - p2.y) < 2 and abs(p1.x - p2.x) > 5:\n"
        "                        underline_ys.append((min(p1.x, p2.x), max(p1.x, p2.x), p1.y))\n"
        "    except Exception:\n"
        "        pass\n"
        "    return underline_ys\n"
        "\n"
        "def is_underlined(span, underline_ys, tolerance=4):\n"
        "    \"\"\"Проверяет, есть ли горизонтальная линия сразу под span.\"\"\"\n"
        "    sx0 = span['bbox'][0]\n"
        "    sx1 = span['bbox'][2]\n"
        "    sy1 = span['bbox'][3]  # нижний край спана\n"
        "    for (lx0, lx1, ly) in underline_ys:\n"
        "        if abs(ly - sy1) < tolerance:\n"
        "            overlap = min(sx1, lx1) - max(sx0, lx0)\n"
        "            if overlap > (sx1 - sx0) * 0.4:\n"
        "                return True\n"
        "    return False\n"
        "\n"
        "doc = fitz.open(r'''" + filePath + "''')\n"
                     "blocks_data = []\n"
                     "texts = []\n"
                     "global_left_margin = 85.0\n"
                     "first_page = True\n"
                     "\n"
                     "for page in doc:\n"
                     "    page_width = page.rect.width\n"
                     "    underline_ys = build_underline_set(page)\n"
                     "    blocks = page.get_text('dict')['blocks']\n"
                     "    left_margin, right_margin = calc_page_margins(blocks, page_width)\n"
                     "    if first_page:\n"
                     "        global_left_margin = left_margin\n"
                     "        first_page = False\n"
                     "    \n"
                     "    for block in blocks:\n"
                     "        if 'lines' not in block:\n"
                     "            continue\n"
                     "        \n"
                     "        for line in block['lines']:\n"
                     "            line_text_parts = []\n"
                     "            line_runs = []\n"
                     "            \n"
                     "            for span in line['spans']:\n"
                     "                text = span['text']\n"
                     "                if not text.strip():\n"
                     "                    continue\n"
                     "                \n"
                     "                flags = span['flags']\n"
                     "                bold = bool(flags & 2 ** 4)\n"
                     "                italic = bool(flags & 2 ** 1)\n"
                     "                underline = is_underlined(span, underline_ys)\n"
                     "                font = span['font']\n"
                     "                size = span['size']\n"
                     "                \n"
                     "                line_runs.append({\n"
                     "                    'text': text,\n"
                     "                    'bold': bold,\n"
                     "                    'italic': italic,\n"
                     "                    'underline': underline,\n"
                     "                    'fontName': font,\n"
                     "                    'fontSize': round(size, 1)\n"
                     "                })\n"
                     "                line_text_parts.append(text)\n"
                     "            \n"
                     "            if line_runs:\n"
                     "                full_text = ''.join(line_text_parts)\n"
                     "                \n"
                     "                has_bold = any(r['bold'] for r in line_runs)\n"
                     "                has_italic = any(r['italic'] for r in line_runs)\n"
                     "                has_underline = any(r['underline'] for r in line_runs)\n"
                     "                \n"
                     "                font_counts = {}\n"
                     "                size_counts = {}\n"
                     "                for r in line_runs:\n"
                     "                    font_counts[r['fontName']] = font_counts.get(r['fontName'], 0) + len(r['text'])\n"
                     "                    size_counts[r['fontSize']] = size_counts.get(r['fontSize'], 0) + len(r['text'])\n"
                     "                \n"
                     "                main_font = max(font_counts, key=font_counts.get) if font_counts else ''\n"
                     "                main_size = max(size_counts, key=size_counts.get) if size_counts else 0\n"
                     "                \n"
                     "                alignment = detect_alignment(line, page_width, left_margin, right_margin)\n"
                     "                x0_real = line['bbox'][0]\n"
                     "                \n"
                     "                block_data = {\n"
                     "                    'type': 'paragraph',\n"
                     "                    'text': full_text,\n"
                     "                    'bold': has_bold,\n"
                     "                    'italic': has_italic,\n"
                     "                    'underline': has_underline,\n"
                     "                    'headingLevel': 0,\n"
                     "                    'isHeading': False,\n"
                     "                    'fontName': main_font,\n"
                     "                    'fontSize': main_size,\n"
                     "                    'alignment': alignment,\n"
                     "                    'xPos': x0_real\n"
                     "                }\n"
                     "                blocks_data.append(block_data)\n"
                     "                texts.append(full_text)\n"
                     "\n"
                     "doc.close()\n"
                     "\n"
                     "actual_left_margin = global_left_margin\n"
                     "\n"
                     "print('###JSON###')\n"
                     "print(json.dumps(blocks_data, ensure_ascii=False))\n"
                     "print('###MARGIN###')\n"
                     "print(actual_left_margin)\n"
                     "print('###TEXT###')\n"
                     "print('\\n'.join(texts))\n";

    process.start("py", QStringList() << "-3" << "-c" << script);
    process.waitForFinished();

    QString output = QString::fromUtf8(process.readAllStandardOutput());

    int jsonStart  = output.indexOf("###JSON###");
    int marginStart = output.indexOf("###MARGIN###");
    int textStart  = output.indexOf("###TEXT###");

    if (jsonStart == -1 || textStart == -1)
        return false;

    int jsonEnd = (marginStart != -1) ? marginStart : textStart;
    QString jsonPart = output.mid(jsonStart + 10, jsonEnd - (jsonStart + 10)).trimmed();

    if (marginStart != -1) {
        QString marginPart = output.mid(marginStart + 12, textStart - (marginStart + 12)).trimmed();
        bool ok = false;
        double margin = marginPart.toDouble(&ok);
        if (ok && margin > 10.0)
            m_pdfLeftMargin = margin;
    }

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
        r.xPos = obj["xPos"].toDouble();

        m_runs.append(r);
    }

    return true;
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
                int cellIdx = 0;
                for (const QString &cell : cells) {
                    QJsonObject cellObj;
                    cellObj["text"] = cell.trimmed();

                    if (r < origTableData.size()) {
                        QJsonArray origRow = origTableData[r].toArray();
                        if (cellIdx < origRow.size()) {
                            QJsonObject origCell = origRow[cellIdx].toObject();
                            cellObj["fontName"] = origCell["fontName"].toString();
                            cellObj["fontSize"] = origCell["fontSize"].toInt();
                            cellObj["bold"] = origCell["bold"].toBool();
                            cellObj["italic"] = origCell["italic"].toBool();
                            cellObj["underline"] = origCell["underline"].toBool();
                            cellObj["alignment"] = origCell["alignment"].toString();
                        }
                    }

                    rowArray.append(cellObj);
                    cellIdx++;
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
            obj["xPos"] = run.xPos;
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
                 "                    cell = table.rows[r].cells[c]\n"
                 "                    cell_data = row_data[c]\n"
                 "                    \n"
                 "                    cell.text = ''\n"
                 "                    p = cell.paragraphs[0]\n"
                 "                    run = p.add_run(cell_data['text'])\n"
                 "                    \n"
                 "                    if cell_data.get('bold', False):\n"
                 "                        run.bold = True\n"
                 "                    if cell_data.get('italic', False):\n"
                 "                        run.italic = True\n"
                 "                    if cell_data.get('underline', False):\n"
                 "                        run.underline = True\n"
                 "                    \n"
                 "                    font_name = cell_data.get('fontName', '')\n"
                 "                    font_size = cell_data.get('fontSize', 0)\n"
                 "                    if font_name:\n"
                 "                        run.font.name = font_name\n"
                 "                    if font_size > 0:\n"
                 "                        run.font.size = Pt(font_size)\n"
                 "                    \n"
                 "                    alignment = cell_data.get('alignment', '')\n"
                 "                    align_map = {\n"
                 "                        'left': WD_ALIGN_PARAGRAPH.LEFT,\n"
                 "                        'center': WD_ALIGN_PARAGRAPH.CENTER,\n"
                 "                        'right': WD_ALIGN_PARAGRAPH.RIGHT,\n"
                 "                        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY\n"
                 "                    }\n"
                 "                    if alignment in align_map:\n"
                 "                        p.alignment = align_map[alignment]\n"
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
    QString normalizedText = text;
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

    QStringList textLines;
    QStringList tableLines;

    for (const QString &line : lines) {
        if (line.contains(" | ")) {
            tableLines.append(line);
        } else {
            textLines.append(line);
        }
    }

    int textRunsCount = 0;
    for (const auto &run : m_runs) {
        if (!run.isTable) textRunsCount++;
    }

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
                int cellIdx = 0;
                for (const QString &cell : cells) {
                    QJsonObject cellObj;
                    cellObj["text"] = cell.trimmed();

                    if (r < origTableData.size()) {
                        QJsonArray origRow = origTableData[r].toArray();
                        if (cellIdx < origRow.size()) {
                            QJsonObject origCell = origRow[cellIdx].toObject();
                            cellObj["fontName"] = origCell["fontName"].toString();
                            cellObj["fontSize"] = origCell["fontSize"].toInt();
                            cellObj["bold"] = origCell["bold"].toBool();
                            cellObj["italic"] = origCell["italic"].toBool();
                            cellObj["underline"] = origCell["underline"].toBool();
                            cellObj["alignment"] = origCell["alignment"].toString();
                        }
                    }

                    rowArray.append(cellObj);
                    cellIdx++;
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
            obj["xPos"] = run.xPos;
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
        "from fpdf import FPDF\n"
        "import sys\n"
        "import traceback\n"
        "sys.stdout.reconfigure(encoding='utf-8')\n"
        "sys.stderr.reconfigure(encoding='utf-8')\n"
        "\n"
        "try:\n"
        "    class PDF(FPDF):\n"
        "        def header(self):\n"
        "            pass\n"
        "        def footer(self):\n"
        "            pass\n"
        "\n"
        "    pdf = PDF()\n"
        "    pdf.set_margins(left=" + QString::number(m_pdfLeftMargin / 2.835, 'f', 1) +
        ", top=25, right=" + QString::number(m_pdfLeftMargin / 2.835 * 0.9, 'f', 1) + ")\n"
                                                                                      "    pdf.set_auto_page_break(auto=True, margin=20)\n"
                                                                                      "    pdf.add_page()\n"
                                                                                      "\n"
                                                                                      "    data = json.loads(r'''" + json + "''')\n"
                 "\n"
                 "    for item in data:\n"
                 "        is_table = item.get('isTable', False)\n"
                 "        \n"
                 "        if is_table:\n"
                 "            table_data = item.get('tableData', [])\n"
                 "            cols = item.get('tableCols', 1)\n"
                 "            \n"
                 "            if table_data and cols > 0:\n"
                 "                rows = len(table_data)\n"
                 "                col_width = (pdf.w - 30) / cols\n"
                 "                line_height = 6\n"
                 "                \n"
                 "                for r in range(rows):\n"
                 "                    row_data = table_data[r]\n"
                 "                    \n"
                 "                    max_lines = 1\n"
                 "                    for c in range(min(len(row_data), cols)):\n"
                 "                        cell_data = row_data[c]\n"
                 "                        text = cell_data.get('text', '')\n"
                 "                        font_size = cell_data.get('fontSize', 11)\n"
                 "                        if font_size <= 0:\n"
                 "                            font_size = 11\n"
                 "                        pdf.set_font_size(font_size)\n"
                 "                        text_width = pdf.get_string_width(text)\n"
                 "                        lines_needed = max(1, int(text_width / (col_width - 2)) + 1)\n"
                 "                        max_lines = max(max_lines, lines_needed)\n"
                 "                    \n"
                 "                    row_height = max_lines * line_height\n"
                 "                    \n"
                 "                    x_start = pdf.l_margin\n"
                 "                    y_start = pdf.get_y()\n"
                 "                    \n"
                 "                    for c in range(min(len(row_data), cols)):\n"
                 "                        cell_data = row_data[c]\n"
                 "                        text = cell_data.get('text', '')\n"
                 "                        bold = cell_data.get('bold', False)\n"
                 "                        italic = cell_data.get('italic', False)\n"
                 "                        underline = cell_data.get('underline', False)\n"
                 "                        font_size = cell_data.get('fontSize', 11)\n"
                 "                        alignment = cell_data.get('alignment', '')\n"
                 "                        \n"
                 "                        if font_size <= 0:\n"
                 "                            font_size = 11\n"
                 "                        \n"
                 "                        font_name = cell_data.get('fontName', '')\n"
                 "                        \n"
                 "                        if 'times' in font_name.lower() or 'timesnewroman' in font_name.lower():\n"
                 "                            family = 'Times'\n"
                 "                        elif 'arial' in font_name.lower():\n"
                 "                            family = 'Arial'\n"
                 "                        elif 'courier' in font_name.lower():\n"
                 "                            family = 'Courier'\n"
                 "                        else:\n"
                 "                            family = 'Arial'\n"
                 "                        \n"
                 "                        style = ''\n"
                 "                        if bold:\n"
                 "                            style += 'B'\n"
                 "                        if italic:\n"
                 "                            style += 'I'\n"
                 "                        if underline:\n"
                 "                            style += 'U'\n"
                 "                        \n"
                 "                        pdf.set_font(family, style, font_size)\n"
                 "                        \n"
                 "                        x = x_start + c * col_width\n"
                 "                        pdf.rect(x, y_start, col_width, row_height)\n"
                 "                        \n"
                 "                        if alignment == 'center':\n"
                 "                            align = 'C'\n"
                 "                        elif alignment == 'right':\n"
                 "                            align = 'R'\n"
                 "                        else:\n"
                 "                            align = 'L'\n"
                 "                        \n"
                 "                        pdf.set_xy(x, y_start)\n"
                 "                        pdf.multi_cell(col_width, line_height, text, border=0, align=align)\n"
                 "                    \n"
                 "                    pdf.set_y(y_start + row_height)\n"
                 "                    \n"
                 "                    if pdf.get_y() > pdf.h - 30:\n"
                 "                        pdf.add_page()\n"
                 "        else:\n"
                 "            text = item.get('text', '')\n"
                 "            bold = item.get('bold', False)\n"
                 "            italic = item.get('italic', False)\n"
                 "            underline = item.get('underline', False)\n"
                 "            font_name = item.get('fontName', '')\n"
                 "            font_size = item.get('fontSize', 11)\n"
                 "            alignment = item.get('alignment', '')\n"
                 "            x_pos = item.get('xPos', 0)\n"
                 "            \n"
                 "            if font_size <= 0:\n"
                 "                font_size = 11\n"
                 "            \n"
                 "            if 'times' in font_name.lower() or 'timesnewroman' in font_name.lower():\n"
                 "                family = 'Times'\n"
                 "            elif 'arial' in font_name.lower():\n"
                 "                family = 'Arial'\n"
                 "            elif 'courier' in font_name.lower():\n"
                 "                family = 'Courier'\n"
                 "            else:\n"
                 "                family = 'Arial'\n"
                 "            \n"
                 "            style = ''\n"
                 "            if bold:\n"
                 "                style += 'B'\n"
                 "            if italic:\n"
                 "                style += 'I'\n"
                 "            if underline:\n"
                 "                style += 'U'\n"
                 "            \n"
                 "            pdf.set_font(family, style, font_size)\n"
                 "            \n"
                 "            if not text.strip():\n"
                 "                pdf.ln(font_size / 2)\n"
                 "                continue\n"
                 "            \n"
                 "            text_width = pdf.get_string_width(text)\n"
                 "            page_width = pdf.w - pdf.l_margin - pdf.r_margin\n"
                 "            line_h = font_size * 0.35 + 2\n"
                 "            \n"
                 "            if alignment == 'center':\n"
                 "                pdf.multi_cell(0, line_h, text, align='C')\n"
                 "            elif alignment == 'right':\n"
                 "                pdf.multi_cell(0, line_h, text, align='R')\n"
                 "            elif alignment == 'justify':\n"
                 "                pdf.multi_cell(0, line_h, text, align='J')\n"
                 "            else:\n"
                 "                pdf.multi_cell(0, line_h, text, align='L')\n"
                 "            \n"
                 "            pdf.ln(font_size * 0.15)\n"
                 "            \n"
                 "            if pdf.get_y() > pdf.h - 30:\n"
                 "                pdf.add_page()\n"
                 "\n"
                 "    pdf.output(r'''" + safePath + "''')\n"
                     "    print('OK')\n"
                     "except Exception as e:\n"
                     "    print('ERROR: ' + str(e), file=sys.stderr)\n"
                     "    traceback.print_exc(file=sys.stderr)\n"
                     "    sys.exit(1)\n";

    process.start("py", QStringList() << "-3" << "-c" << script);

    if (!process.waitForFinished(60000)) {
        qDebug() << "PDF build timeout or crash";
        return false;
    }

    int exitCode = process.exitCode();
    QString stdOut = QString::fromUtf8(process.readAllStandardOutput()).trimmed();
    QString stdErr = QString::fromUtf8(process.readAllStandardError()).trimmed();

    qDebug() << "PDF build exit code:" << exitCode;
    qDebug() << "PDF build stdout:" << stdOut;
    qDebug() << "PDF build stderr:" << stdErr;

    if (exitCode != 0) {
        qDebug() << "PDF build failed with exit code:" << exitCode;
        return false;
    }

    if (!stdOut.contains("OK")) {
        qDebug() << "PDF build did not return OK, stdout:" << stdOut;
        return false;
    }

    return true;
}
