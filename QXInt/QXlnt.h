#pragma once

#include "qxlnt_global.h"

#include <QMap>
#include <QObject>

#include <xlnt/xlnt.hpp>
class QXINT_EXPORT QXlnt {
public:
    QXlnt();

    void setSheetsTitle(const QStringList& sheetsTitle = {});

    /**
     * @brief 设置表单 单元格内容
     * @param sheetTitle 表单标题
     * @param row
     * @param column
     * @param value
     */
    void setCell(QString sheetTitle, int column, int row, QVariant value);

    /**
     * @brief 设置表单 表头
     * @param sheetTitle 表单标题
     * @param column
     */
    void setHeaders(QString sheetTitle, int column);

    /**
     * @brief 保存excel文件
     */
    void save(QString path);

    /**
     * @brief 设置excel文件标题
     * @param title
     */
    void setTitle(QString title);

    /**
     * @brief 设置表单内容
     * @param sheetTitle
     * @param datas
     */
    void setDatas(QString sheetTitle, QList<QList<QVariant>> datas);

    /**
     * @brief 获取当前表单的总行数
     * @param sheetTitle
     * @return
     */
    size_t currentRowLength(QString sheetTitle);

    /**
     * @brief 去读excel文件
     * @param path
     * @return
     */
    QList<QVariantList> readExcel(QString path);

    /**
     * @brief 转换excel文件为txt文件
     * @param path excel文件路径
     * @param splitter 每个单元格之间的分隔符
     * @param suffix txt文件后缀
     * @param s 行尾是否加分隔符
     * @return txt文件路径
     */
    QString ConvertExcel2Txt(QString path, QString splitter, QString suffix, bool s = true);

private:
    xlnt::workbook _wb;

    /**
     * @brief <表格名称,表格>
     */
    QMap<QString, xlnt::worksheet> _sheetMap {};
};
