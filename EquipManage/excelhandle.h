#ifndef EXCELHANDLE_H
#define EXCELHANDLE_H

#include <QList>
#include <QVariant>
#include <QAxObject>


class ExcelHandle
{
private:
    QAxObject *mExcel;
    QAxObject *mWorkbooks;
    QAxObject *mWorkbook;
    QAxObject * mWorksheets;
    QAxObject * mWorksheet;

public:
    ExcelHandle(QString &filePath);
    ~ExcelHandle();

    bool readSheetAllCells(int sheetNum, QList<QList<QVariant>>& data);
    void readRowsAndColsOfSheet(int sheetNum, int &startRow, int &rows, int &startCol, int &cols);
    void saveAs(QString &filePath);     //另存为
    void setCellBackground(int sheetNum, int row, int col);       //设置单元格背景色
    void save();    //保存
    void setWorkSheet(int sheetNum);    //设置目前操作的sheet


};

#endif // EXCELHANDLE_H
