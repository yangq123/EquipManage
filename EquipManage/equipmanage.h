#ifndef EQUIPMANAGE_H
#define EQUIPMANAGE_H

#include <QMainWindow>
#include "excelhandle.h"
#include <QSet>
#include <QStandardItemModel>
#include <QtSql>
#include <QProgressDialog>
#include "cmparedbdialog.h"
#include "libxl.h"
using namespace libxl;

QT_BEGIN_NAMESPACE
namespace Ui { class EquipManage; }
QT_END_NAMESPACE

/******错误类型******/
enum CellErrorType  {
    nullError,      //空错误
    ruleError,      //规则错误
    matchError      //匹配错误
};

/******字段错误信息******/
struct ErrorInfo
{
    int num;        //字段序号
    QString fieldName;      //字段名称
    CellErrorType errType;      //字段错误类型
    QString errReason;      //错误原因
    QString recomContent;   //推荐填写内容


    ErrorInfo(int n, QString name, CellErrorType type)
    {
        num=n;
        fieldName=name;
        errType=type;
    }
    ErrorInfo(int n, QString name, CellErrorType type, QString errRe)
    {
        num=n;
        fieldName=name;
        errType=type;
        errReason=errRe;
    }
    ErrorInfo(int n, QString name, CellErrorType type, QString errRe, QString content)
    {
        num=n;
        fieldName=name;
        errType=type;
        errReason=errRe;
        recomContent=content;
    }

};

/******厂家、型号类******/
class ProducerTypeInfo
{
public:
    QString producer;       //厂家
    QString type;           //型号

public:
    ProducerTypeInfo(QString p, QString t)
    {
        producer=p;
        type=t;
    }

//    friend bool operator ==(const ProducerTypeInfo &pt1, const ProducerTypeInfo &pt2);

    bool operator == (const ProducerTypeInfo &pt) const        //重载==运算符
    {
        if(producer==pt.producer && type==pt.type) return true;
        else
            return false;
    }
};

uint qHash(const ProducerTypeInfo &pt);

//厂家型号相同的记录分组信息
struct GroupOfProType
{
    QPair<QString, QString> proAndType;     //厂家型号
    QList<QSet<int>> rownumSet;     //厂家型号关联字段相同的行号集合
    QList<int> diffFild;        //集合间不同字段的字段序号List
};


class EquipManage : public QMainWindow
{
    Q_OBJECT
private:
//    ExcelHandle *mExcelHandle;      //Excel连接
    QString mFilePath;       //样本库路径
    QString mFileExportPath;        //样本库导出路径
    QStandardItemModel *mTableModel;    //模型数据
    QStandardItemModel *mTableModelOfMajor;     //显示多数条目的model
    QStandardItemModel *mTableModelOfWrongItem; //显示格式错误、与对比库中不符的条目信息
    QSqlDatabase mDuanDB;        //对比库连接
    QProgressDialog *mProgress;      //进度条
    int mValueOfProgress;
    cmpareDbDialog *mCmpDbDialog;       //对比库对话框


    int mRowsOfSheet;
    int mColsOfSheet;
    int mStartRowOfSheet;
    int mStartColOfSheet;
    QList<QList<QVariant>> mExcelSheetData;     //excel某个sheet的所有单元格数据
    //mWrongItems,mCorrectItems,mManuWrongItems互斥，且并集为mExcelSheetData
    QMap<int, QList<ErrorInfo>> mWrongItems;    //sheet中错误条目信息,每个键值对表示某个条目，int表示行号，QList<ErrorInfo>表示某行出错字段信息
    QList<int> mCorrectItems;    //sheet中正确条目行号
    QList<int> mManuWrongItems;      //人工判断为错误的条目
    QMap<int, ProducerTypeInfo> mWriteDbItems;        //待写入对比库的条目<行号,厂家型号信息>

    QList<int> mRecogProduAndTypeItems;      //基本规则正确的条目，待对厂家和型号信息进行检查
    QMap<int, ProducerTypeInfo> mNoRecOfDbItems;    //对比库中无记录条目，int为序号，ProducerTypeInfo为厂家型号信息
    QSet<ProducerTypeInfo> mUniqNoRecOfDbItems;     //对比库中无记录条目的厂家型号信息
    //mFewOfMinorItems 和 mNotExistSameFldsOfMinorItems 互斥，且并集为mMinorItems
    QList<int> mMinorItems;      //样本库中厂家、型号相同数目较少的条目，以及尽管数目多但不满足存在2/3以上相关属性相同，待人工判断
    QMap<int, int> mFewOfMinorItems;        //样本库中厂家、型号相同数目较少的条目,key:行号，value：与该行的厂家型号相同的记录数目
    QMap<int, int> mNotExistSameFldsOfMinorItems;       //尽管数目多但不满足存在2/3以上相关属性相同,key:行号，value：与该行的厂家型号相同的记录数目

    //QMap<int, QList<int>> mRedFidOfMinorItems; //mMinorItems中相互之间不同的字段号，int：行号，QList<int>:字段序号
    QMap<QPair<QString,QString>, GroupOfProType> mRedFidOfMinorItems;//记录mMinorItems中的集合信息，以及标红信息,QPair:厂家型号，GroupOfProType:集合信息

    QMap<int,QList<int>> mMinorOfMajorItems;      //样本库中厂家、型号相同数目较多的条目中存在的少数相关属性不同条目，int序号，QList<int>条目中与mMajorOfMajorItems字段不同的序号，待人工判断
    QList<QList<int>> mSetOfMinorOfMajorItems;  //mMinorOfMajorItems中厂家型号关联字段相同的行号的集合
    QMap<int,QString> corPro;    //mMinorOfMajorItems中条目的正确概率(该条目数/总条目数)
    QMap<int, ProducerTypeInfo> mMajorOfMajorItems;      //样本库中厂家、型号相同数目较多的条目中大部分相同的条目，作为疑似正确条目，待mMinorOfMajorItems判断后以修正
    QList<int> mColOfProAndType;        //与厂家、型号信息相关联的字段序号


    void judgeDuanluqi();       //判断断路器
    void judgeFieldsOfDuanluqi(int row, QList<ErrorInfo>& wrongFieldList, QMap<QString, int> &fieldNameMap);     //判断断路器各个字段
    void queryDbOfDuanluqi();       //查询对比库，对有记录的条目进行字段正误识别，无记录的条目进行记录
    void establishCompaDBOfDuanluqi();      //建立/完善断路器对比库
    void initColNumAboutProAndType();       //初始化与厂家和型号信息相关的字段序号
    void showExcelData();       //显示数据
    void writeDataToDB();       //写入对比库
    void exportExcelData();     //导出Excel
    QString convertToColName(int col);    //数字转换为excel列名称
    void convertFieldNameToInt(QMap<QString, int> &fieldMap);   //将字段名称转换为整型
    void clear();
    GroupOfProType establishGroupOfProType(ProducerTypeInfo &proType, QMap<int, QList<QString>> &fieldsMap);   //针对同一厂家型号记录进行分组比较

    void readAllCellsByXlsx(QString path);

    //通过LibXl读取excel（支持xlsx，xls）
    void readAllCellsByLibXL(QString path);
    //通过LibXl复制excel（支持xlsx，xls）
    void copyExcelDataByLibXL(Book* srcBook, Book* dstBook, const wchar_t *pathDes);
    //通过LibXl导出excel（支持xlsx，xls）
    void exportExcelDataByLibXL();
    //显示excel格式错误数据和与对比库中不符的数据条目信息
    void showWrongItemExcelData();

public:
    EquipManage(QWidget *parent = nullptr);
    ~EquipManage();

private slots:
    void on_importAction_triggered();

    void on_exportAction_triggered();

    void on_aboutAction_triggered();
    void tableViewCliked(const QModelIndex & index);    //显示mMinorOfMajorItems对应的多数(2/3)条目,作为与错误条目的正确参考


    void on_cmpDbClickedAction_triggered();

private:
    Ui::EquipManage *ui;
};
#endif // EQUIPMANAGE_H
