#pragma once

#include "qxlnt_global.h"

#include <QMap>
#include <QObject>

#include <xlnt/xlnt.hpp>
class QXINT_EXPORT QXlnt {
public:
    QXlnt();

    QXlnt(const QStringList& sheetsTitle);

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

private:
    xlnt::workbook _wb;

    /**
     * @brief <表格名称,表格>
     */
    QMap<QString, xlnt::worksheet> _sheetMap {};
};
