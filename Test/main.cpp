#include <QVariant>
#include <QtCore/QCoreApplication>

#include "QXlnt.h"
#include <xlnt/xlnt.hpp>
int main(int argc, char* argv[])
{
    QCoreApplication a(argc, argv);
    QXlnt xlnt;
    xlnt.ConvertExcel2Txt("C:\\Users\\18435\\Desktop\\1.xls", "  ", "in1");
    return a.exec();
}
