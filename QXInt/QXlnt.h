#pragma once

#include "qxlnt_global.h"

#include <QMap>
#include <QObject>

#include <xlnt/xlnt.hpp>
class QXINT_EXPORT QXlnt : public QObject {
    Q_OBJECT
public:
    explicit QXlnt(QObject* parent = nullptr);

    /**
     * @brief 保存 Excel 文件
     * @param path 文件路径
     * @return 成功返回 true，失败返回 false
     */
    bool load(QString path);

    /**
     * @brief 保存excel文件
     */
    bool save(QString path);

    /**
     * @brief 设置excel文件标题
     * @param title
     */
    void setTitle(QString title);

    /**
     * @brief 创建
     * @param sheetsTitle
     * @return
     */
    bool createSheets(const QStringList& sheetsTitle = {});

    /**
     * @brief 删除表单
     * @param sheetTitle
     * @return
     */
    bool removeSheet(const QString& sheetTitle);

    /**
     * @brief 设置表单 单元格内容
     * @param sheetTitle 表单标题
     * @param row
     * @param column
     * @param value
     */
    bool setCell(const QString& sheetTitle, int row, int column, const QVariant& value);

    /**
     * @brief 设置表单内容
     * @param sheetTitle
     * @param datas
     */
    bool setDatas(QString sheetTitle, QList<QList<QVariant>> datas, int startRow = 0, int startColumn = 0);

    /**
     * @brief 读取表单内容
     * @param sheetTitle
     * @return
     */
    QList<QVariantList> readSheet(const QString& sheetTitle) const;

    /**
     * @brief 转换为文本文件
     * @param excelPath
     * @param txtPath
     * @param delimiter 分隔符
     * @param includeTrailingSeparator
     * @return
     */
    bool convertToText(const QString& excelPath, const QString& txtPath, const QString& delimiter, bool includeTrailingSeparator);
signals :
    /**
     * @brief 错误信号
     */
    void errored(const QString& error);

private:
    /**
     * @brief 获取 xlnt 表单
     * @param sheetTitle
     * @return
     */
    std::optional<xlnt::worksheet> getSheet(const QString& sheetTitle) const;

    /**
     * @brief 获取 xlnt 单元格内容
     * @param cell
     * @return
     */
    QVariant getXlntCellValue(const xlnt::cell& cell) const;

    /**
     * @brief 设置 xlnt 单元格内容
     * @param cell
     * @param value
     */
    void setXlntCellValue(xlnt::cell& cell, const QVariant& value);

    /**
     * @brief xlnt 表单是否存在
     * @param sheetTitle
     * @return
     */
    bool hasSheet(const QString& sheetTitle) const;

private:
    xlnt::workbook _wb;

    /**
     * @brief <表格名称,表格>
     */
    QMap<QString, xlnt::worksheet> _sheetMap {};
};
