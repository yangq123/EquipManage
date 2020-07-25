#include "excelhandle.h"
#include <QDebug>
#include <QDir>
#include <QColor>

ExcelHandle::ExcelHandle(QString &filePath)
{
    mExcel = new QAxObject("Excel.Application");
    mExcel->dynamicCall("SetVisible(bool)", false);
    mWorkbooks = mExcel->querySubObject("WorkBooks");
    mWorkbook = mWorkbooks->querySubObject("Open(QString &)", filePath);
    mWorksheets = mWorkbook->querySubObject("WorkSheets");
}

ExcelHandle::~ExcelHandle()
{
    mWorkbook->dynamicCall("Close (Boolean)", false);  //关闭文件
    mExcel->dynamicCall("Quit(void)");  //退出
    if(mExcel != NULL)   {delete mExcel;mExcel=NULL;}
//    if(mWorkbooks != NULL)   {delete mWorkbooks;mWorkbooks=NULL;}
//    if(mWorkbook != NULL)    {delete mWorkbook;mWorkbook=NULL;}
//    if(mWorksheets != NULL)   {delete mWorksheets;mWorksheets=NULL;}
//    if(mWorksheet != NULL)   {delete mWorksheet;mWorksheet=NULL;}


}

//读取某个sheet所有单元格的值，存入data
 bool ExcelHandle::readSheetAllCells(int sheetNum, QList<QList<QVariant>>& data)
{
    //worksheet = workbook->querySubObject("WorkSheets(int)", 1); // 获取工作表集合的工作表1， 即sheet1
    mWorksheet = mWorksheets->querySubObject("Item(int)", sheetNum);

    //获取整个sheet的值
    QVariant var;
    if (mWorksheet != NULL && ! mWorksheet->isNull())
    {
        QAxObject *usedRange = mWorksheet->querySubObject("UsedRange");
        if(NULL == usedRange || usedRange->isNull())
        {
            return false;
        }
        var = usedRange->dynamicCall("Value");      //var值是一个table，其实是QList<QList<QVariant>>
        delete usedRange;
    }
    else
        return false;

    //将var转换为QList<QList<QVariant>>类型
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
        return false;
    }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        data.push_back(rowData);
    }
    return true;
}


  void ExcelHandle::readRowsAndColsOfSheet(int sheetNum, int &startRow, int &rowNums, int &startCol, int &colNums)
  {
      mWorksheet = mWorksheets->querySubObject("Item(int)", sheetNum);

      //表格范围和行数列数
      QAxObject * usedrange = mWorksheet->querySubObject("UsedRange");

      QAxObject * rows = usedrange->querySubObject("Rows");
      rowNums = rows->property("Count").toInt();
      qDebug() << QString("行数为: %1").arg(QString::number(rowNums));

      QAxObject * columns = usedrange->querySubObject("Columns");
      colNums = columns->property("Count").toInt();
      qDebug() << QString("列数为: %1").arg(QString::number(colNums));

      startRow = rows->property("Row").toInt();
      qDebug() << QString("起始行为: %1").arg(QString::number(startRow));

      startCol = columns->property("Column").toInt();
      qDebug() << QString("起始列为: %1").arg(QString::number(startCol));

  }


  void ExcelHandle::saveAs(QString &filePath)
  {
      QVariant res=mWorkbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));
      if(!res.isValid())
      {
          qDebug()<<"saveAs ERROR!";
      }
  }

  void ExcelHandle::setWorkSheet(int sheetNum)
  {
      mWorksheet = mWorksheets->querySubObject("Item(int)", sheetNum);
  }

  void ExcelHandle::setCellBackground(int sheetNum, int row, int col)
  {
      //mWorksheet = mWorksheets->querySubObject("Item(int)", sheetNum);


      //测试
//      QAxObject* usedrange = mWorksheet->querySubObject("UsedRange"); // sheet范围
//          int intRowStart = usedrange->property("Row").toInt(); // 起始行数   为1
//          int intColStart = usedrange->property("Column").toInt();  // 起始列数 为1


//          QAxObject *rows, *columns;
//          rows = usedrange->querySubObject("Rows");  // 行
//          columns = usedrange->querySubObject("Columns");  // 列


//          int intRow = rows->property("Count").toInt(); // 行数
//          int intCol = columns->property("Count").toInt();  // 列数
//           qDebug()<<"intRowStart:"<<intRowStart<<"\t intColStart"<<intColStart;
//          qDebug()<<"intRow"<<intRow<<"\t intCol"<<intCol;







      QAxObject* cell = mWorksheet->querySubObject("Cells(int, int)", row, col);
//      QString cellStr=cell->dynamicCall("Value2()").toString();
//      qDebug() << row << col << cellStr;

      QAxObject* interior = cell->querySubObject("Interior");
      bool r=interior->setProperty("Color", QColor(255, 0, 0));   //设置单元格背景色（红色）
//      qDebug()<<"set color res = "<<r;
  }

  void ExcelHandle::save()
  {
      mWorkbook->dynamicCall("Save()");  //保存文件
  }








