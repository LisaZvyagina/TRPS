#include "simpledocparser.h"

#include <QFile>
#include <QTextStream>
#include <QProcess>
#include <QPdfWriter>
#include <QPainter>
#include <QFont>

SimpleDocParser::SimpleDocParser(QObject *parent) : QObject(parent) {}

bool SimpleDocParser::parse(const QString &filePath, QString &outText, QString &outFormatInfo)
{
    Q_UNUSED(outFormatInfo);

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
    QProcess process;

    QString script =
        "import sys\n"
        "import docx\n"
        "from docx.table import Table\n"
        "from docx.text.paragraph import Paragraph\n"
        "from docx.document import Document as _Document\n"
        "sys.stdout.reconfigure(encoding='utf-8')\n"

        "def iter_block_items(parent):\n"
        "    for child in parent.element.body:\n"
        "        if child.tag.endswith('p'):\n"
        "            yield Paragraph(child, parent)\n"
        "        elif child.tag.endswith('tbl'):\n"
        "            yield Table(child, parent)\n"

        "doc = docx.Document(r'''"+ filePath + "''')\n"
                     "texts = []\n"

                     "for block in iter_block_items(doc):\n"
                     "    if isinstance(block, Paragraph):\n"
                     "        if block.text.strip():\n"
                     "            texts.append(block.text)\n"

                     "    elif isinstance(block, Table):\n"
                     "        for row in block.rows:\n"
                     "            row_text = []\n"
                     "            for cell in row.cells:\n"
                     "                row_text.append(cell.text.strip())\n"
                     "            texts.append(' | '.join(row_text))\n"

                     "print('\\n'.join(texts))\n";

    process.start("python", QStringList() << "-c" << script);
    process.waitForFinished();

    outText = QString::fromUtf8(process.readAllStandardOutput());
    return !outText.isEmpty();
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

bool SimpleDocParser::buildDOCX(const QString &text, const QString &path)
{
    QString safe = text;
    safe.replace("\\", "\\\\");
    safe.replace("'''", " ");

    QProcess process;

    QString script =
        "from docx import Document\n"
        "doc = Document()\n"
        "lines = r'''"+ safe + "'''.split('\\n')\n"

                 "for line in lines:\n"
                 "    if '|' in line:\n"
                 "        cells = [c.strip() for c in line.split('|')]\n"
                 "        table = doc.add_table(rows=1, cols=len(cells))\n"
                 "        row = table.rows[0].cells\n"
                 "        for i in range(len(cells)):\n"
                 "            row[i].text = cells[i]\n"
                 "    else:\n"
                 "        doc.add_paragraph(line)\n"

                 "doc.save(r'''"+ path + "''')\n"
                 "print('OK')\n";

    process.start("python", QStringList() << "-c" << script);
    process.waitForFinished();

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

        // 🔥 пустая строка = абзац
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

        // 🔥 отступ между абзацами
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