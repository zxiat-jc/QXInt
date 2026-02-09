#include "QXlnt.h"

#include <QFile>
#include <QMetaType>
#include <QVariant>

QXlnt::QXlnt(QObject* parent)
{
}

QXlnt::~QXlnt()
{
    _sheetMap.clear();
}

bool QXlnt::load(QString path)
{
    try {
        _wb.load(path.toStdString());
        _sheetMap.clear();

        // 加载所有工作表到映射
        for (const auto& sheet : _wb) {
            QString title = QString::fromStdString(sheet.title());
            _sheetMap[title] = sheet;
        }

        return true;
    } catch (const std::exception& e) {
        emit errored(QString("加载文件失败: %1").arg(e.what()));
        return false;
    }
}

bool QXlnt::save(QString path)
{
    try {
        _wb.save(path.toStdString());

        return true;
    } catch (const std::exception& e) {
        emit errored(QString("保存文件失败: %1").arg(e.what()));
        return false;
    }
}

void QXlnt::setTitle(QString title)
{
    _wb.title(title.toStdString());
}

bool QXlnt::createSheets(const QStringList& sheetsTitle)
{
    for (auto&& sheetTitle : sheetsTitle) {
        try {
            if (_sheetMap.contains(sheetTitle)) {
                return true;
            }

            xlnt::worksheet ws;
            if (_wb.sheet_count() == 0) {
                ws = _wb.create_sheet();
            } else if (_wb.contains(sheetTitle.toStdString())) {
                ws = _wb.sheet_by_title(sheetTitle.toStdString());
            } else {
                ws = _wb.create_sheet();
            }

            ws.title(sheetTitle.toStdString());
            _sheetMap[sheetTitle] = ws;

            return true;
        } catch (const std::exception& e) {
            emit errored(QString("创建/选择工作表失败: %1").arg(e.what()));
            return false;
        }
    }
    return true;
}

bool QXlnt::removeSheet(const QString& sheetTitle)
{
    try {
        if (!_wb.contains(sheetTitle.toStdString())) {
            emit errored("工作表不存在");
            return false;
        }

        _wb.remove_sheet(_wb.sheet_by_title(sheetTitle.toStdString()));
        _sheetMap.remove(sheetTitle);

        return true;
    } catch (const std::exception& e) {
        emit errored(QString("删除工作表失败: %1").arg(e.what()));
        return false;
    }
    return true;
}

bool QXlnt::setCell(const QString& sheetTitle, int row, int column, const QVariant& value)
{
    try {
        auto&& opt = getSheet(sheetTitle);
        if (!opt.has_value()) {
            return false;
        }
        auto&& ws = opt.value();
        auto cell = ws.cell(xlnt::cell_reference(column, row));
        setXlntCellValue(cell, value);
        cell.alignment(xlnt::alignment()
                .wrap(true)
                .horizontal(xlnt::horizontal_alignment::center)
                .vertical(xlnt::vertical_alignment::center));
    } catch (const std::exception& e) {
        emit errored(QString("设置单元格失败: %1").arg(e.what()));
        return false;
    }
    return true;
}

bool QXlnt::setDatas(QString sheetTitle, QList<QList<QVariant>> datas, int startRow, int startColumn)
{
    try {
        auto sheetOpt = getSheet(sheetTitle);
        if (!sheetOpt) {
            return false;
        }

        auto& ws = sheetOpt.value();

        for (int i = 0; i < datas.size(); ++i) {
            for (int j = 0; j < datas[i].size(); ++j) {
                this->setCell(sheetTitle, startRow + i, startColumn + j, datas[i][j]);
            }
        }
        return true;
    } catch (const std::exception& e) {
        emit errored(QString("设置数据失败: %1").arg(e.what()));
        return false;
    }
}

bool QXlnt::setCellBorder(const QString& sheetTitle, int row, int column, BorderSides sides)
{
    try {
        auto opt = getSheet(sheetTitle);
        if (!opt.has_value()) {
            return false;
        }
        auto ws = opt.value();
        auto cell = ws.cell(xlnt::cell_reference(column, row));

        xlnt::border b;
        auto prop = xlnt::border::border_property().style(xlnt::border_style::thin);
        if (sides.testFlag(BorderLeft)) {
            b.side(xlnt::border_side::start, prop);
        }
        if (sides.testFlag(BorderRight)) {
            b.side(xlnt::border_side::end, prop);
        }
        if (sides.testFlag(BorderTop)) {
            b.side(xlnt::border_side::top, prop);
        }
        if (sides.testFlag(BorderBottom)) {
            b.side(xlnt::border_side::bottom, prop);
        }

        cell.border(b);
        return true;
    } catch (const std::exception& e) {
        emit errored(QString("设置单元格边框失败: %1").arg(e.what()));
        return false;
    }
}

bool QXlnt::setRangeBorder(const QString& sheetTitle, int startRow, int startColumn, int endRow, int endColumn, BorderSides sides)
{
    try {
        auto opt = getSheet(sheetTitle);
        if (!opt.has_value()) {
            return false;
        }

        for (int r = startRow; r <= endRow; ++r) {
            for (int c = startColumn; c <= endColumn; ++c) {
                setCellBorder(sheetTitle, r, c, sides);
            }
        }
        return true;
    } catch (const std::exception& e) {
        emit errored(QString("设置区域边框失败: %1").arg(e.what()));
        return false;
    }
}

bool QXlnt::mergeCells(const QString& sheetTitle, int startRow, int startColumn, int endRow, int endColumn)
{
    try {
        auto opt = getSheet(sheetTitle);
        if (!opt.has_value()) {
            return false;
        }
        auto ws = opt.value();

        xlnt::cell_reference start(startColumn, startRow);
        xlnt::cell_reference end(endColumn, endRow);
        xlnt::range_reference range(start, end);
        ws.merge_cells(range);
        return true;
    } catch (const std::exception& e) {
        emit errored(QString("合并单元格失败: %1").arg(e.what()));
        return false;
    }
}

QList<QVariantList> QXlnt::readSheet(const QString& sheetTitle) const
{
    QList<QVariantList> result;

    try {
        auto&& opt = this->getSheet(sheetTitle);
        if (!opt.has_value()) {
            return result;
        }

        auto&& ws = opt.value();
        for (auto row : ws.rows(false)) {
            QVariantList rowData;
            for (auto cell : row) {
                rowData.append(getXlntCellValue(cell));
            }
            result.append(rowData);
        }

    } catch (const std::exception& e) {
        qDebug() << "读取工作表失败:" << e.what();
    }

    return result;
}

bool QXlnt::convertToText(const QString& excelPath, const QString& txtPath, const QString& delimiter, bool includeTrailingSeparator)
{
    try {
        if (!load(excelPath)) {
            emit errored("无法加载 Excel 文件");
            return false;
        }

        QFile file(txtPath);
        if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
            emit errored("无法创建 TXT 文件");
            return false;
        }

        QTextStream out(&file);
        for (auto&& sheetName : _sheetMap.keys()) {
            auto&& data = readSheet(sheetName);
            for (const auto& row : data) {
                for (int i = 0; i < row.size(); ++i) {
                    out << row[i].toString();

                    if (i < row.size() - 1 || includeTrailingSeparator) {
                        out << delimiter;
                    }
                }
                out << "\n";
            }
        }
        file.close();

        return true;

    } catch (const std::exception& e) {
        emit errored(QString("转换失败: %1").arg(e.what()));
        return false;
    }
}

std::optional<xlnt::worksheet> QXlnt::getSheet(const QString& sheetTitle) const
{
    if (hasSheet(sheetTitle)) {
        return _wb.sheet_by_title(sheetTitle.toStdString());
    }
    return std::nullopt;
}

QVariant QXlnt::getXlntCellValue(const xlnt::cell& cell) const
{
    if (!cell.has_value()) {
        return QVariant();
    }

    switch (cell.data_type()) {
    case xlnt::cell::type::number:
        return cell.value<double>();

    case xlnt::cell::type::boolean:
        return cell.value<bool>();

    case xlnt::cell::type::shared_string:
    case xlnt::cell::type::inline_string:
    case xlnt::cell::type::formula_string:
        return QString::fromStdString(cell.value<std::string>());

    default:
        return QString::fromStdString(cell.to_string());
    }
}

void QXlnt::setXlntCellValue(xlnt::cell& cell, const QVariant& value)
{
    switch (value.typeId()) {
    case QMetaType::Int:
    case QMetaType::LongLong:
        cell.value(value.toLongLong());
        break;

    case QMetaType::Double:
    case QMetaType::Float:
        cell.value(value.toDouble());
        break;

    case QMetaType::Bool:
        cell.value(value.toBool());
        break;

    case QMetaType::QString:
    default:
        cell.value(value.toString().toStdString());
        break;
    }
}

bool QXlnt::hasSheet(const QString& sheetTitle) const
{
    return _wb.contains(sheetTitle.toStdString());
}
