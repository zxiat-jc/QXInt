#include "QXlnt.h"

#include <QFile>
#include <QMetaType>
#include <QVariant>

QXlnt::QXlnt()
{
}

void QXlnt::setSheetsTitle(const QStringList& sheetsTitle)
{
    int i = 0;
    for (int i = 0; i < sheetsTitle.length(); i++) {
        if (i == 0) {
            // 获取活动表单
            _sheetMap[sheetsTitle[i]] = _wb.active_sheet();
        } else {
            // 新建表单
            _sheetMap[sheetsTitle[i]] = _wb.create_sheet();
        }
        // 设置表单标题
        _sheetMap[sheetsTitle[i]].title(sheetsTitle[i].toStdString());
    }
}

void QXlnt::setCell(QString sheetTitle, int row, int column, QVariant value)
{
    if (_sheetMap.contains(sheetTitle)) {
        auto&& worksheet = _sheetMap[sheetTitle];
        if (value.typeId() == QMetaType::Double) {
            worksheet.cell(xlnt::cell_reference(column, row)).value(QString::number(value.toDouble(), 'f', 5).toDouble());
        } else {
            worksheet.cell(xlnt::cell_reference(column, row)).value(value.toString().toStdString());
        }
        worksheet.cell(xlnt::cell_reference(column, row)).alignment(xlnt::alignment().wrap(true).horizontal(xlnt::horizontal_alignment::center).vertical(xlnt::vertical_alignment::center));
    }
}

void QXlnt::setHeaders(QString sheetTitle, int column)
{
    if (_sheetMap.contains(sheetTitle)) {
        auto currentRow = this->currentRowLength(sheetTitle);
        if (currentRow <= 0) {
            return;
        }
        auto&& worksheet = _sheetMap[sheetTitle];
        xlnt::cell titleBlock = worksheet.cell(xlnt::cell_reference(1, currentRow));
        titleBlock.value(sheetTitle.toStdString());
        titleBlock.font(xlnt::font().size(25).bold(true).name("Arial"));
        titleBlock.alignment(xlnt::alignment().wrap(false).horizontal(xlnt::horizontal_alignment::center).vertical(xlnt::vertical_alignment::center));
        worksheet.merge_cells(xlnt::range_reference(xlnt::cell_reference(1, currentRow), xlnt::cell_reference(column, currentRow)));
    }
}

void QXlnt::save(QString path)
{
    _wb.save(path.toStdString());
}

void QXlnt::setTitle(QString title)
{
    _wb.title(title.toStdString());
}

void QXlnt::setDatas(QString sheetTitle, QList<QList<QVariant>> datas)
{

    if (_sheetMap.contains(sheetTitle)) {
        auto currentRow = this->currentRowLength(sheetTitle);
        if (currentRow <= 0) {
            return;
        }
        for (int i = 1; i <= datas.length(); i++) {
            for (int j = 1; j <= datas[i - 1].length(); j++) {
                this->setCell(sheetTitle, currentRow + i, j, datas[i - 1][j - 1]);
            }
        }
    }
}

size_t QXlnt::currentRowLength(QString sheetTitle)
{
    if (!_sheetMap.contains(sheetTitle)) {
        return 0;
    }
    auto&& worksheet = _sheetMap[sheetTitle];
    return worksheet.rows(true).length();
}

QList<QVariantList> QXlnt::readExcel(QString path)
{
    QList<QVariantList> result;
    try {
        xlnt::workbook wb;
        wb.load(path.toStdString());
        auto ws = wb.active_sheet();
        for (auto row : ws.rows(true)) {
            QVariantList str;
            for (auto cell : row) {
                QString content = QString::fromStdString(cell.to_string());
                if (content.trimmed().isEmpty()) {
                    continue;
                }
                str.append(QString::fromStdString(cell.to_string()));
            }
            qDebug() << str;
            result.append(str);
        }
        qDebug() << "____process completed_____________";
    } catch (std::exception& e) {
        qDebug() << "___error: " << QString::fromStdString(e.what());
        return {};
    }
    return result;
}

QString QXlnt::ConvertExcel2Txt(QString path, QString splitter, QString suffix, bool s)
{
    auto&& result = this->readExcel(path);
    if (result.isEmpty()) {
        return {};
    }
    QFile excelFile(path);
    path.replace(path.length() - 4, 4, suffix);
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        return {};
    }
    QTextStream out(&file);
    for (int i = 0; i < result.length(); i++) {
        for (int j = 0; j < result[i].length(); j++) {
            if (j == result[i].length() - 1 && !s) {
                out << result[i][j].toString();
            } else {
                out << result[i][j].toString() << splitter;
            }
        }
        out << "\n";
    }
    file.close();
    return path;
}
