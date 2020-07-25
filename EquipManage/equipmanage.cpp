#include "equipmanage.h"
#include "ui_equipmanage.h"
#include <QFileDialog>
#include <QDebug>
#include <QMessageBox>
#include "xlsxdocument.h"
#include "xlsxconditionalformatting.h"
#include "qwaiting.h"
#include <QPair>
#include "xlsxworksheet.h"
#include "xlsxworkbook.h"

#include <iostream>
#include <string>
using namespace std;


const int SHEET_NUM = 1;
const wchar_t * NAME_LIBXL = L"Halil Kural";
const wchar_t * KEY_LIBXL = L"windows-2723210a07c4e90162b26966a8jcdboe";
#define XLSX 0
#define LIBXL 1

EquipManage::EquipManage(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::EquipManage)
{
    ui->setupUi(this);
    ui->textLabel->setText("");
    ui->textLabel1->setText("");
    ui->textLabel_wrongItem->setText("");
    setWindowState(Qt::WindowMaximized);
    mTableModel = new QStandardItemModel();
    mTableModelOfMajor = new QStandardItemModel();
    mTableModelOfWrongItem =  new QStandardItemModel();

    connect(ui->tableView, SIGNAL(clicked(QModelIndex)), this, SLOT(tableViewCliked(QModelIndex)));
}

EquipManage::~EquipManage()
{
    delete ui;
}

void EquipManage::initColNumAboutProAndType()
{
    int eddyIndex=mExcelSheetData[0].indexOf("额定电压");
    mColOfProAndType.push_back(eddyIndex);

    int eddlIndex=mExcelSheetData[0].indexOf("额定电流");
    mColOfProAndType.push_back(eddlIndex);

    int edjyspIndex=mExcelSheetData[0].indexOf("额定绝缘水平");
    mColOfProAndType.push_back(edjyspIndex);

    int eddldlkdcsIndex=mExcelSheetData[0].indexOf("额定短路电流开断次数");
    mColOfProAndType.push_back(eddldlkdcsIndex);

    int eddlkddlIndex=mExcelSheetData[0].indexOf("额定短路开断电流");
    mColOfProAndType.push_back(eddlkddlIndex);

    int eddlghdlIndex=mExcelSheetData[0].indexOf("额定短路关合电流");
    mColOfProAndType.push_back(eddlghdlIndex);

    int dwddlIndex=mExcelSheetData[0].indexOf("动稳定电流");
    mColOfProAndType.push_back(dwddlIndex);

    int rwddlIndex=mExcelSheetData[0].indexOf("热稳定电流");
    mColOfProAndType.push_back(rwddlIndex);

    int eddlcxsjIndex=mExcelSheetData[0].indexOf("额定短路持续时间");
    mColOfProAndType.push_back(eddlcxsjIndex);

    int dkslIndex=mExcelSheetData[0].indexOf("断口数量");
    mColOfProAndType.push_back(dkslIndex);

    int tgpdjlIndex=mExcelSheetData[0].indexOf("套管爬电距离");
    mColOfProAndType.push_back(tgpdjlIndex);

    int tgghjlIndex=mExcelSheetData[0].indexOf("套管干弧距离");
    mColOfProAndType.push_back(tgghjlIndex);

    int jxsmIndex=mExcelSheetData[0].indexOf("机械寿命");
    mColOfProAndType.push_back(jxsmIndex);

    int syhjIndex=mExcelSheetData[0].indexOf("使用环境");
    mColOfProAndType.push_back(syhjIndex);

    int jgxsIndex=mExcelSheetData[0].indexOf("结构型式");
    mColOfProAndType.push_back(jgxsIndex);

    int mhjzIndex=mExcelSheetData[0].indexOf("灭弧介质");
    mColOfProAndType.push_back(mhjzIndex);

    int czjgxsIndex=mExcelSheetData[0].indexOf("操作机构型式");
    mColOfProAndType.push_back(czjgxsIndex);


}

//将字段名称转换为整型
void EquipManage::convertFieldNameToInt(QMap<QString, int> &fieldMap)
{

    /******
     * fieldNameList中已有字段顺序不能改变！！！
     * 新加入字段只能从list最后加入！！！
     * 用于judgeFieldsOfDuanluqi函数中switch case操作
     *
    *****/
    const QStringList fieldNameList = {"所属地市", "设备名称", "运行编号", "设备编码", "所属电站", "间隔单元", "运维单位",
                                       "维护班组", "电压等级", "设备状态", "相数", "相别", "投运日期", "组合设备类型", "是否农网",
                                       "使用环境", "专业分类", "型号", "结构型式", "操作机构型式", "灭弧介质", "生产厂家", "出厂日期",
                                       "额定电压", "额定电流", "额定频率", "额定绝缘水平", "额定短路电流开断次数", "额定短路开断电流",
                                       "额定短路关合电流", "动稳定电流", "热稳定电流", "额定短路持续时间", "断口数量", "套管爬电距离",
                                       "套管干弧距离", "机械寿命", "合闸时间", "分闸时间", "合分时间", "资产性质", "资产单位",
                                       "设备增加方式", "登记时间", "实物ID", "最近投运日期", "组合设备名称"};
    /******
     * QMap中字段名称与索引号对应关系不能改变！！！
     * 用于judgeFieldsOfDuanluqi函数中switch case操作
    *****/
    for(int i=1; i<=fieldNameList.size(); i++)
    {
        fieldMap.insert(fieldNameList[i-1], i);
    }

}

void EquipManage::judgeDuanluqi()
{
    QMap<QString, int> fieldMap;    //存储字段名称与整型索引对应关系
    convertFieldNameToInt(fieldMap);

    //基本规则判断
    for(int i=1; i<mRowsOfSheet; i++)
    {
        QList<ErrorInfo> wrongFieldList;    //出错字段list
        judgeFieldsOfDuanluqi(i, wrongFieldList, fieldMap);
        if(!wrongFieldList.isEmpty())
        {
            mWrongItems.insert(i, wrongFieldList);      //基本规则错误，记录错误条目
            wrongFieldList.clear();
        }
        else
        {
            mRecogProduAndTypeItems.append(i);      //待对厂家和型号信息进行检查
        }
        QCoreApplication::processEvents();
        mValueOfProgress++;
        mProgress->setValue(mValueOfProgress);
    }


    //初始化与厂家和型号信息相关的字段序号
    initColNumAboutProAndType();

    QCoreApplication::processEvents();
    mValueOfProgress++;
    mProgress->setValue(mValueOfProgress);

    //针对基本规则正确的条目查对比库
    queryDbOfDuanluqi();

    QCoreApplication::processEvents();
    mValueOfProgress++;
    mProgress->setValue(mValueOfProgress);

    //通过样本库建立对比库，待人工判断后确认写入对比库条目
    establishCompaDBOfDuanluqi();

}

GroupOfProType EquipManage::establishGroupOfProType(ProducerTypeInfo &proType, QMap<int, QList<QString>> &fieldsMap)
{
    QMap<int, QList<QString>>::const_iterator iter=fieldsMap.constBegin();
    QList<int> allRows = fieldsMap.keys();
    QList<int> foundRows;
    QList<QSet<int>> rownumSet;     //分组集合
    QList<int> diffFild;        //集合间不同字段的字段序号List

    if(fieldsMap.size() == 1)
    {
        rownumSet.append(QSet<int>(allRows.begin(), allRows.end()));
        diffFild.append(mColOfProAndType);
    }
    else
    {
        //建立行号分组集合
        while(iter != fieldsMap.constEnd())
        {
            QList<QString> fildValue = iter.value();
            QList<int> rows = fieldsMap.keys(fildValue);
            if(!foundRows.contains(rows[0]))
            {
                foundRows.append(rows);
                rownumSet.append(QSet<int>(rows.begin(), rows.end()));
            }
            iter++;
        }
        if(rownumSet.size() == 1 && fieldsMap.size() < 3)   //针对记录数小于3但是字段都相同的情况
        {
            diffFild.append(mColOfProAndType);
        }
        else
        {
            //建立集合间不同字段序号List
            for(int i=0; i<mColOfProAndType.size(); i++)
            {
                QSet<QString> fildStr;
                for(int j=0; j<rownumSet.size(); j++)
                {
                    int num = *(rownumSet[j].begin());
                    QString fldstr = fieldsMap[num][i];
                    fildStr.insert(fldstr);
                }
                if(fildStr.size() != 1)
                {
                    diffFild.append(mColOfProAndType[i]);
                }
            }
        }
    }

    GroupOfProType group;
    QPair<QString, QString> proAndType(proType.producer, proType.type);
    group.proAndType = proAndType;
    group.rownumSet = rownumSet;
    group.diffFild = diffFild;

    return group;
}



void EquipManage::establishCompaDBOfDuanluqi()
{
    QSet<ProducerTypeInfo>::const_iterator i=mUniqNoRecOfDbItems.constBegin();
    while(i!=mUniqNoRecOfDbItems.constEnd())
    {
        ProducerTypeInfo ptinfo=*i;
        QPair<QString, QString> protype(ptinfo.producer, ptinfo.type);
        QList<int> indexList=mNoRecOfDbItems.keys(ptinfo);      //具有相同厂家、型号信息的行号
        QMap<int, QList<QString>> fieldsOfProAndTypeMap;    //具有相同厂家、型号信息的各个条目的与厂家、型号相关的字段信息(int行号，QList<QString>字段list)
        QList<QString> fieldsOfProAndTypeList;      //每一行字段list

        //获取具有相同厂家和型号的关联字段信息
        for(int i1=0; i1<indexList.size(); i1++)
        {
            for(int j=0; j<mColOfProAndType.size(); j++)
            {
                fieldsOfProAndTypeList.push_back(mExcelSheetData[indexList[i1]][mColOfProAndType[j]].toString());
            }
            fieldsOfProAndTypeMap.insert(indexList[i1], fieldsOfProAndTypeList);
            fieldsOfProAndTypeList.clear();
        }

        //样本库中厂家、型号相同数目较少的条目,待人工判断
        if(indexList.size() < 3)
        {
            GroupOfProType group = establishGroupOfProType(ptinfo, fieldsOfProAndTypeMap);
            mRedFidOfMinorItems.insert(protype, group);
            for(int k=0; k<indexList.size(); k++)
            {
                mMinorItems.push_back(indexList[k]);
                mFewOfMinorItems.insert(indexList[k], indexList.size());        //by yang, 2020.7.18
            }
            i++;
            continue;
        }



        //找出2/3以上相同的条目
        int majorFlag=0;    //是否存在2/3以上部分相同标志位
        QList<int> allRows = fieldsOfProAndTypeMap.keys();  //具有相同厂家、型号信息的所有行号
        QMap<int, QList<QString>>::const_iterator i2=fieldsOfProAndTypeMap.constBegin();
        while(i2!=fieldsOfProAndTypeMap.constEnd())
        {
            QList<QString> value=i2.value();
            QList<int> majorRows=fieldsOfProAndTypeMap.keys(value);  //具有与value的所有字段相同的所有行号
            if(static_cast<double>(majorRows.count())/fieldsOfProAndTypeMap.size() > 0.5)
            {
                majorFlag=1;
                for(int j2=0; j2<allRows.size(); j2++)
                {
                    if(majorRows.contains(allRows[j2]))        //属于2/3部分
                    {
                        mMajorOfMajorItems.insert(allRows[j2], ptinfo);      //存储行号和厂家、型号信息
                    }
                    //不属于2/3部分
                    else
                    {
                        QList<QString> fieldList=fieldsOfProAndTypeMap[allRows[j2]];
                        QList<int> diffField;
                        for(int j3=0; j3<value.size(); j3++)        //与2/3部分逐个字段比较
                        {
                            if(fieldList[j3] != value[j3])
                            {
                                diffField.push_back(mColOfProAndType[j3]);
                            }
                        }
                        mMinorOfMajorItems.insert(allRows[j2], diffField);      //记录1/3部分的疑似错误字段，待人工判断

                        //完善mSetOfMinorOfMajorItems
                        QList<int> sameNo = fieldsOfProAndTypeMap.keys(fieldList);  //与该条目的字段相同的条目行号
                        int flag=0;
                        for(int i=0; i<mSetOfMinorOfMajorItems.size(); i++)
                        {
                            if(mSetOfMinorOfMajorItems[i].contains(sameNo[0]))
                            {
                                flag=1;
                                break;
                            }
                        }
                        if(flag == 0)
                        {
                            mSetOfMinorOfMajorItems.append(sameNo);
                        }
                        int count = sameNo.size();   //与该条目的字段相同的条目数目
                        QString correctPro=QString::number(count)+"/"+QString::number(fieldsOfProAndTypeMap.size()); //该条目的正确概率
                        //double correctPro=(static_cast<double>(count))/fieldsOfProAndTypeMap.size();    //该条目的正确概率
                        corPro.insert(allRows[j2], correctPro);
                    }
                }
                break;
            }
            i2++;
        }
        //不存在2/3以上部分相同
        if(majorFlag == 0)
        {
            GroupOfProType group = establishGroupOfProType(ptinfo, fieldsOfProAndTypeMap);
            mRedFidOfMinorItems.insert(protype, group);
            mMinorItems.append(allRows);       //待人工判断

            //by yang, 2020.7.18
            for(int k=0; k<allRows.size(); k++)
            {
                mNotExistSameFldsOfMinorItems.insert(allRows[k], allRows.size());
            }
        }
        i++;
    }
}


void EquipManage::queryDbOfDuanluqi()
{   
    static int openCount=0;
    //QString dbFileName = QFileDialog::getOpenFileName(this, "", ".", "*.db3 *.db");

    QString dbFileName=QDir::currentPath()+"/equipManage.db";
    mDuanDB = QSqlDatabase::addDatabase("QSQLITE", QString("DuanDB%1").arg(++openCount));
    mDuanDB.setDatabaseName(dbFileName);
    if(!mDuanDB.open())
    {
        mDuanDB = QSqlDatabase();
        QSqlDatabase::removeDatabase(QString("ZhcData%1").arg(openCount));
        QMessageBox::critical(0, QObject::tr("DataBase Error"), mDuanDB.lastError().text());
        return;
    }

    QList<ErrorInfo> wrongFieldList;    //出错字段list
    for(int i=0; i<mRecogProduAndTypeItems.size(); i++)
    {
        int proIndex=mExcelSheetData[0].indexOf("生产厂家");
        QString producerStr=mExcelSheetData[mRecogProduAndTypeItems[i]][proIndex].toString();
        int typeIndex=mExcelSheetData[0].indexOf("型号");
        QString typeStr=mExcelSheetData[mRecogProduAndTypeItems[i]][typeIndex].toString();

        QString sqlStr="select eddy,eddl,edjysp,eddldlkdcs,eddlkddl,eddlghdl,dwddl,rwddl,eddlcxsj,dksl,tgpdjl,tgghjl,jxsm,"
                       "czjgxs, syhj, jgxs, mhjz from compareLib_DuanLuQi where sccj='"+producerStr+"' and xh='"+typeStr+"'";
//        qDebug()<<producerStr;
//        qDebug()<<sqlStr;
        QSqlQuery query(mDuanDB);
        if(!query.exec(sqlStr))
        {
            qDebug()<<query.lastError();
        }
        if(query.next())
        {
            QSqlRecord rec = query.record();
            int eddyNo = rec.indexOf("eddy");
            QString eddyStr = query.value(eddyNo).toString();
            int eddyIndex=mExcelSheetData[0].indexOf("额定电压");
            if(eddyStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][eddyIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = eddyStr;
                ErrorInfo errInfo(eddyIndex,"额定电压", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int eddlNo = rec.indexOf("eddl");
            QString eddlStr = query.value(eddlNo).toString();
            int eddlIndex=mExcelSheetData[0].indexOf("额定电流");
            if(eddlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][eddlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = eddlStr;
                ErrorInfo errInfo(eddlIndex,"额定电流", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int edjyspNo = rec.indexOf("edjysp");
            QString edjyspStr = query.value(edjyspNo).toString();
            int edjyspIndex=mExcelSheetData[0].indexOf("额定绝缘水平");
            if(edjyspStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][edjyspIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = edjyspStr;
                ErrorInfo errInfo(edjyspIndex,"额定绝缘水平", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int eddldlkdcsNo = rec.indexOf("eddldlkdcs");
            QString eddldlkdcsStr = query.value(eddldlkdcsNo).toString();
            int eddldlkdcsIndex=mExcelSheetData[0].indexOf("额定短路电流开断次数");
            if(eddldlkdcsStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][eddldlkdcsIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = eddldlkdcsStr;
                ErrorInfo errInfo(eddldlkdcsIndex,"额定短路电流开断次数", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int eddlkddlNo = rec.indexOf("eddlkddl");
            QString eddlkddlStr = query.value(eddlkddlNo).toString();
            int eddlkddlIndex=mExcelSheetData[0].indexOf("额定短路开断电流");
            if(eddlkddlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][eddlkddlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = eddlkddlStr;
                ErrorInfo errInfo(eddlkddlIndex,"额定短路开断电流", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int eddlghdlNo = rec.indexOf("eddlghdl");
            QString eddlghdlStr = query.value(eddlghdlNo).toString();
            int eddlghdlIndex=mExcelSheetData[0].indexOf("额定短路关合电流");
            if(eddlghdlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][eddlghdlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = eddlghdlStr;
                ErrorInfo errInfo(eddlghdlIndex,"额定短路关合电流", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int dwddlNo = rec.indexOf("dwddl");
            QString dwddlStr = query.value(dwddlNo).toString();
            int dwddlIndex=mExcelSheetData[0].indexOf("动稳定电流");
            if(dwddlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][dwddlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = dwddlStr;
                ErrorInfo errInfo(dwddlIndex,"动稳定电流", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int rwddlNo = rec.indexOf("rwddl");
            QString rwddlStr = query.value(rwddlNo).toString();
            int rwddlIndex=mExcelSheetData[0].indexOf("热稳定电流");
            if(rwddlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][rwddlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = rwddlStr;
                ErrorInfo errInfo(rwddlIndex,"热稳定电流", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int eddlcxsjNo = rec.indexOf("eddlcxsj");
            QString eddlcxsjStr = query.value(eddlcxsjNo).toString();
            int eddlcxsjIndex=mExcelSheetData[0].indexOf("额定短路持续时间");
            if(eddlcxsjStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][eddlcxsjIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = eddlcxsjStr;
                ErrorInfo errInfo(eddlcxsjIndex,"额定短路持续时间", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int dkslNo = rec.indexOf("dksl");
            QString dkslStr = query.value(dkslNo).toString();
            int dkslIndex=mExcelSheetData[0].indexOf("断口数量");
            if(dkslStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][dkslIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = dkslStr;
                ErrorInfo errInfo(dkslIndex,"断口数量", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int tgpdjlNo = rec.indexOf("tgpdjl");
            QString tgpdjlStr = query.value(tgpdjlNo).toString();
            int tgpdjlIndex=mExcelSheetData[0].indexOf("套管爬电距离");
            if(tgpdjlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][tgpdjlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = tgpdjlStr;
                ErrorInfo errInfo(tgpdjlIndex,"套管爬电距离", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int tgghjlNo = rec.indexOf("tgghjl");
            QString tgghjlStr = query.value(tgghjlNo).toString();
            int tgghjlIndex=mExcelSheetData[0].indexOf("套管干弧距离");
            if(tgghjlStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][tgghjlIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = tgghjlStr;
                ErrorInfo errInfo(tgghjlIndex,"套管干弧距离", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int jxsmNo = rec.indexOf("jxsm");
            QString jxsmStr = query.value(jxsmNo).toString();
            int jxsmIndex=mExcelSheetData[0].indexOf("机械寿命");
            if(jxsmStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][jxsmIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = jxsmStr;
                ErrorInfo errInfo(jxsmIndex,"机械寿命", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int czjgxsNo = rec.indexOf("czjgxs");
            QString czjgxsStr = query.value(czjgxsNo).toString();
            int czjgxsIndex=mExcelSheetData[0].indexOf("操作机构型式");
            if(czjgxsStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][czjgxsIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = czjgxsStr;
                ErrorInfo errInfo(czjgxsIndex,"操作机构型式", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int syhjNo = rec.indexOf("syhj");
            QString syhjStr = query.value(syhjNo).toString();
            int syhjIndex=mExcelSheetData[0].indexOf("使用环境");
            if(syhjStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][syhjIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = syhjStr;
                ErrorInfo errInfo(syhjIndex,"使用环境", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int jgxsNo = rec.indexOf("jgxs");
            QString jgxsStr = query.value(jgxsNo).toString();
            int jgxsIndex=mExcelSheetData[0].indexOf("结构型式");
            if(jgxsStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][jgxsIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = jgxsStr;
                ErrorInfo errInfo(jgxsIndex,"结构型式", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            int mhjzNo = rec.indexOf("mhjz");
            QString mhjzStr = query.value(mhjzNo).toString();
            int mhjzIndex=mExcelSheetData[0].indexOf("灭弧介质");
            if(mhjzStr!=mExcelSheetData[mRecogProduAndTypeItems[i]][mhjzIndex].toString())
            {
                QString errRes = "与对比库不符";
                QString cont = mhjzStr;
                ErrorInfo errInfo(mhjzIndex,"灭弧介质", matchError, errRes, cont);
                wrongFieldList.push_back(errInfo);
            }

            if(!wrongFieldList.isEmpty())
            {
                mWrongItems.insert(mRecogProduAndTypeItems[i], wrongFieldList); //错误条目
                wrongFieldList.clear();
            }
            else
            {
                mCorrectItems.push_back(mRecogProduAndTypeItems[i]);    //正确条目
            }
        }
        else        //对比库中无记录
        {
            ProducerTypeInfo noRecItem(producerStr, typeStr);
            mNoRecOfDbItems.insert(mRecogProduAndTypeItems[i], noRecItem);
            if(!mUniqNoRecOfDbItems.contains(noRecItem))
            {
                mUniqNoRecOfDbItems.insert(noRecItem);
            }
        }
    }
}

void EquipManage::tableViewCliked(const QModelIndex & index)
{
//    int no=index.sibling(index.row(), 0).data().toInt();    //该记录的行号(序号)
    int no=index.sibling(index.row(), 2).data().toInt();    //该记录的行号(序号)
    if(mMinorOfMajorItems.contains(no))
    {
        //标记显示绿色的字段
        QList<int> greenList;

        int aa = mTableModel->columnCount();
        qDebug()<<"aa = "<<aa;

        for(int i=1; i<mTableModel->columnCount(); i++)
        {
//            QColor color=mTableModel->item(index.row(), i)->foreground().color();
            QColor color=mTableModel->item(index.row(), i)->foreground().color();
            if(color == QColor(255, 0, 0))
            {
                greenList.push_back(i);
            }
        }

//        QString id = index.sibling(index.row(), 0).data().toString();
        QString id = index.sibling(index.row(), 2).data().toString();
        ui->textLabel1->setText(id+"号记录可参考的正确记录如下：");
        //设置颜色
        QPalette pa;
        pa.setColor(QPalette::WindowText,Qt::red);
        ui->textLabel1->setPalette(pa);


        int sccjIndex = mExcelSheetData[0].indexOf("生产厂家");
        QString sccjStr= mExcelSheetData[id.toInt()][sccjIndex].toString();
        int xhIndex = mExcelSheetData[0].indexOf("型号");
        QString xhStr= mExcelSheetData[id.toInt()][xhIndex].toString();

        ProducerTypeInfo proAndXh(sccjStr, xhStr);
        int majorKey=mMajorOfMajorItems.key(proAndXh);  //多数条目的行号

        mTableModelOfMajor->setRowCount(1);
        mTableModelOfMajor->setColumnCount(mColsOfSheet+2);        //增加正确概率和用户判断是否错误两列
        QStringList headerList;
        for(int i=0; i<mExcelSheetData[0].size(); i++)
        {
            //去掉表头中的换行符，避免冻结列时，表头高度不一致
            QString headerStr = mExcelSheetData[0][i].toString().remove(QString("\n"));
            headerList.append(headerStr);
        }
        headerList.removeFirst();
        headerList.push_front("序号");
//        headerList.push_back("综合判断及正确概率");
        headerList.push_front("综合判断及正确概率");
        headerList.push_back("是否正确");
        mTableModelOfMajor->setHorizontalHeaderLabels(headerList);

        QStandardItem *item;
        for(int j=0; j<mColsOfSheet; j++)
        {
            item = new QStandardItem(mExcelSheetData[majorKey][j].toString());
//            if(greenList.contains(j))
            if(greenList.contains(j+2))
            {
                item->setForeground(QBrush(QColor(0, 255, 0)));     //设置绿色
                //字体加粗
                item->setFont(QFont("Times", 10, QFont::Black));
            }
//            mTableModelOfMajor->setItem(0,j,item);
            mTableModelOfMajor->setItem(0,j+1,item);
        }

        //添加概率
        int majorCount = mMajorOfMajorItems.keys(proAndXh).size();
        int allCount = mNoRecOfDbItems.keys(proAndXh).size();
        QString corProStr = QString("正确率：")+QString::number(majorCount)+"/"+QString::number(allCount);
        item = new QStandardItem(corProStr);
//        mTableModelOfMajor->setItem(0,mColsOfSheet,item);
        mTableModelOfMajor->setItem(0,0,item);
        //添加是否错误判断
        item = new QStandardItem();
        item->setCheckable(true);
        item->setCheckState(Qt::Unchecked);
        mTableModelOfMajor->setItem(0,mColsOfSheet+1,item);

        ui->tableView1->setModel(mTableModelOfMajor);
    }
    else
    {
//        QString idNum = index.sibling(index.row(), 0).data().toString();
        QString idNum = index.sibling(index.row(), 2).data().toString();
        ui->textLabel1->setText(idNum+"号记录无可参考的正确记录！");
        //设置颜色
        QPalette pa;
        pa.setColor(QPalette::WindowText,Qt::red);
        ui->textLabel1->setPalette(pa);
        mTableModelOfMajor->clear();
    }


}






void EquipManage::judgeFieldsOfDuanluqi(int row, QList<ErrorInfo>& wrongFieldList, QMap<QString, int> &fieldNameMap)
{
    QRegExp rx;     //正则表达式
    QString capStr;     //提取字符串

    QString yxbh;   //运行编号
    QString tyrq;   //投运日期
    QString ccrq;   //出厂日期
    QString zjtyrq; //最近投运日期
    QString zhsblx; //组合设备类型
    QString zhsbmc; //组合设备名称
    QString jgxs;   //结构形式
    QString mhjz;   //灭弧介质
    QString eddlkddl;   //额定短路开断电流
    QString eddlghdl;   //额定短路关合电流
    QString dwddl;      //动稳定电流
    QString rwddl;      //热稳定电流
    QString dydj;       //电压等级
    QString eddlcxsj;   //额定短路持续时间
    QString dksl;       //断口数量
    QString tgpdjl;     //套管爬电距离
    QString tgghjl;     //套管干弧距离
    QString jxsm;       //机械寿命
    QString hzsj;       //合闸时间
    QString fzsj;       //分闸时间
    QString hfsj;       //合分时间


    for(int col=1; col<mColsOfSheet; col++)
    {
        QString fieldName=mExcelSheetData.at(0).at(col).toString();
        //qDebug()<<fieldName;

        switch (fieldNameMap[fieldName])
        {
        case 1://"所属地市"
        case 7://"运维单位"
        case 8://"维护班组"
        case 5://"所属电站"
        case 18://"型号"
        case 22://"生产厂家"
        case 41://"资产性质"
        case 42://"资产单位"
        case 45://"实物ID"
        {
            rx.setPattern("^\\s*(\\w*)");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1 || rx.cap(1)=="")
            {
                QString errRe="空错误";
                QString cont="不能为空";
                ErrorInfo wrongField(col, fieldName, nullError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 2://"设备名称"
        {
            rx.setPattern("^\\s*((\\d+)\\s*(开关|断路器))\\s*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="运行编号+设备类型";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                if(yxbh.isEmpty())
                    yxbh=rx.cap(2);
                else if( yxbh != rx.cap(2))
                {
                    QString errRe="与运行编号关联错误";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 3://"运行编号"
        {
            QString yxbhStr=mExcelSheetData.at(row).at(col).toString();
            if(!yxbhStr.isEmpty())
            {
                rx.setPattern("^\\s*(\\d+)\\s*$");
                int res=rx.indexIn(yxbhStr);
                if(res == -1)
                {
                    QString errRe="格式错误";
                    QString cont="数字";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                else
                {
                    if(yxbh.isEmpty())
                        yxbh=rx.cap(1);
                    else if( yxbh != rx.cap(1))
                    {
                        QString errRe="与设备名称关联错误";
                        ErrorInfo wrongField(col, fieldName, matchError,errRe);   //出错字段
                        wrongFieldList.push_back(wrongField);
                    }
                }
            }
            break;
        }
        case 6://"间隔单元"
        {
            rx.setPattern("^\\s*\\d+(kv|kV|Kv|KV)\\s*\\S*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="电压等级+空格+设备编号";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 9://"电压等级"
        {
            rx.setPattern("^\\s*交流\\s*(1000|750|500|330|220|110|66|35|20|10|6)\\s*(kv|kV|Kv|KV)\\s*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="交流+数字+kV";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
                dydj=rx.cap(1);
            break;
        }
        case 10://"设备状态"
        {
            rx.setPattern("^\\s*(在运|未投运|退役|现场留用|库存备用|待报废|报废)\\s*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="在运,未投运,退役,现场留用,库存备用,待报废,报废";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 11://"相数"
        {
            rx.setPattern("^\\s*三相\\s*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="三相";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 12://"相别"
        {
            rx.setPattern("^\\s*ABC\\s*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="ABC";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 13://"投运日期"
        {
            rx.setPattern("^((((1[6-9]|[2-9]\\d)\\d{2})/(0?[13578]|1[02])/(0?[1-9]|[12]\\d|3[01]))|"
                          "(((1[6-9]|[2-9]\\d)\\d{2})/(0?[13456789]|1[012])/(0?[1-9]|[12]\\d|30))|"
                          "(((1[6-9]|[2-9]\\d)\\d{2})/0?2/(0?[1-9]|1\\d|2[0-8]))|(((1[6-9]|[2-9]\\d)"
                          "(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))/0?2/29/))$");
            QString date=mExcelSheetData.at(row).at(col).toDate().toString("yyyy/MM/dd");
            int res=rx.indexIn(date);
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="20**/*/**";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                tyrq=rx.cap(0);
                if(!zjtyrq.isEmpty() && zjtyrq<tyrq)
                {
                    QString errRe="与最近投运日期关联错误";
                    QString cont="最近投运日期大于等于投运日期";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                if(!ccrq.isEmpty() && ccrq>tyrq)
                {
                    QString errRe="与出厂日期关联错误";
                    QString cont="投运日期必须在出厂日期之后";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 46://“最近投运日期"
        {
            rx.setPattern("(^\\s*$)|(^((((1[6-9]|[2-9]\\d)\\d{2})/(0?[13578]|1[02])/(0?[1-9]|[12]\\d|3[01]))|"
                          "(((1[6-9]|[2-9]\\d)\\d{2})/(0?[13456789]|1[012])/(0?[1-9]|[12]\\d|30))|"
                          "(((1[6-9]|[2-9]\\d)\\d{2})/0?2/(0?[1-9]|1\\d|2[0-8]))|(((1[6-9]|[2-9]\\d)"
                          "(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))/0?2/29/))$)");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toDate().toString("yyyy/MM/dd"));
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="20**/*/**";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                zjtyrq=rx.cap(0);
                QString tmp=zjtyrq.remove(" ");
                if(!tmp.isEmpty() && !tyrq.isEmpty() && tyrq>zjtyrq)
                {
                    QString errRe="与投运日期关联错误";
                    QString cont="最近投运日期大于等于投运日期";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                if(!tmp.isEmpty() && !ccrq.isEmpty() && ccrq>zjtyrq)
                {
                    QString errRe="与出厂日期关联错误";
                    QString cont="最近投运日期必须在出厂日期之后";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 23://"出厂日期"
        {
            rx.setPattern("^((((1[6-9]|[2-9]\\d)\\d{2})/(0?[13578]|1[02])/(0?[1-9]|[12]\\d|3[01]))|"
                          "(((1[6-9]|[2-9]\\d)\\d{2})/(0?[13456789]|1[012])/(0?[1-9]|[12]\\d|30))|"
                          "(((1[6-9]|[2-9]\\d)\\d{2})/0?2/(0?[1-9]|1\\d|2[0-8]))|(((1[6-9]|[2-9]\\d)"
                          "(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))/0?2/29/))$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toDate().toString("yyyy/MM/dd"));
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="20**/*/**";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                ccrq=rx.cap(0);
                if(!tyrq.isEmpty() && tyrq<ccrq)
                {
                    QString errRe="与投运日期关联错误";
                    QString cont="投运日期必须在出厂日期之后";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                if(!zjtyrq.isEmpty() && ccrq>zjtyrq)
                {
                    QString errRe="与最近投运日期关联错误";
                    QString cont="最近投运日期必须在出厂日期之后";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 14://"组合设备类型"
        {
            rx.setPattern("^\\s*(否|开关柜|组合电器|电力电容器)\\s*$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="否,开关柜,组合电器,电力电容器";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                zhsblx=rx.cap(0);
                if(!zhsbmc.isNull())
                {
                    if(zhsblx == "否" && !zhsbmc.isEmpty())
                    {
                        QString errRe="与组合设备类型关联错误";
                        int index=mExcelSheetData[0].indexOf("组合设备名称");
                        ErrorInfo wrongField(index, "组合设备名称", matchError,errRe);   //出错字段
                        wrongFieldList.push_back(wrongField);
                    }
                    else if(zhsbmc.isEmpty())
                    {
                        QString errRe="与组合设备类型关联错误";
                        int index=mExcelSheetData[0].indexOf("组合设备名称");
                        ErrorInfo wrongField(index, "组合设备名称", matchError,errRe);   //出错字段
                        wrongFieldList.push_back(wrongField);
                    }
                }
                if(!jgxs.isEmpty())
                {
                    if((zhsblx=="开关柜" && jgxs!="其他") || (zhsblx=="组合电器" && jgxs!="GIS" && jgxs!="HGIS")
                            || (zhsblx=="否" && jgxs!="瓷柱式" && jgxs!="罐式"))
                    {
                        QString errRe="与组合设备类型关联错误";
                        int index=mExcelSheetData[0].indexOf("结构型式");
                        ErrorInfo wrongField(index, "结构型式", matchError,errRe);   //出错字段
                        wrongFieldList.push_back(wrongField);
                    }
                }
            }
            break;
        }
        case 47://"组合设备名称"
        {
            zhsbmc=mExcelSheetData.at(row).at(col).toString();
            if(!zhsblx.isEmpty())
            {
                if(zhsblx=="否" && !zhsbmc.isEmpty())
                {
                    QString errRe="与组合设备类型关联错误";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                else if(zhsblx!="否" && zhsbmc.isEmpty())
                {
                    QString errRe="与组合设备类型关联错误";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 15://"是否农网"
        {
            if(mExcelSheetData.at(row).at(col).toString() != "否")
            {
                QString errRe="格式错误";
                QString cont="否";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 16://"使用环境"
        {
            QString syhj=mExcelSheetData.at(row).at(col).toString();
            if(syhj != "户内式" && syhj!="户外式")
            {
                QString errRe="格式错误";
                QString cont="户内式,户外式";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 17://"专业分类"
        {
            if(mExcelSheetData.at(row).at(col).toString() != "变电")
            {
                QString errRe="格式错误";
                QString cont="变电";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 19://"结构型式"
        {
            rx.setPattern("^(其他|GIS|HGIS|瓷柱式|罐式|PASS|COMPASS)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="其他,GIS,HGIS,瓷柱式,罐式,PASS,COMPASS";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                jgxs=rx.cap(0);
                if(!zhsblx.isEmpty())
                {
                    if((zhsblx=="开关柜" && jgxs!="其他") || (zhsblx=="组合电器" && jgxs!="GIS" && jgxs!="HGIS")
                            || (zhsblx=="否" && jgxs!="瓷柱式" && jgxs!="罐式"))
                    {
                        QString errRe="与组合设备类型关联错误";
                        ErrorInfo wrongField(col, fieldName, matchError,errRe);   //出错字段
                        wrongFieldList.push_back(wrongField);
                    }
                }
                if(mhjz=="多油" && jgxs!="罐式")
                {
                    QString errRe="与灭弧介质关联错误";
                    QString cont="罐式";
                    ErrorInfo wrongField(col, fieldName, matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 20://"操作机构型式"
        {
            rx.setPattern("^(弹簧|电磁|液压|液簧|气动|气动弹簧|其他)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="弹簧,电磁,液压,液簧,气动,气动弹簧,其他";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 21://"灭弧介质"
        {
            rx.setPattern("^(SF6|真空|多油|少油|空气)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="SF6,真空,多油,少油,空气";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                mhjz=rx.cap(0);
                if(!jgxs.isEmpty() && mhjz=="多油" && jgxs!="罐式")
                {
                    QString errRe="与灭弧介质关联错误";
                    QString cont="罐式";
                    int index=mExcelSheetData[0].indexOf("结构型式");
                    ErrorInfo wrongField(index, "结构型式", matchError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 24://"额定电压"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 25://"额定电流"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 26://"额定频率"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                QString edplStr = rx.cap(0);
                double edpl=edplStr.toDouble();
                double diff = edpl-50;
                if(abs(diff) > 0.001)
                {
                    QString errRe="格式错误";
                    QString cont="50";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 27://"额定绝缘水平"
        {
            rx.setPattern("^(\\d+/\\d+)|(\\d+/\\d+/\\d+)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 28://"额定短路电流开断次数"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        case 29://"额定短路开断电流"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                eddlkddl=rx.cap(0);
            }
            break;
        }
        case 30://"额定短路关合电流"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
                eddlghdl=rx.cap(0);
            break;
        }
        case 31://"动稳定电流"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
                dwddl=rx.cap(0);
            break;
        }
        case 32://"热稳定电流"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
                rwddl=rx.cap(0);
            break;
        }
        case 33://"额定短路持续时间"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="2,3,4";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                QString eddlcxsjStr=rx.cap(0);
                double eddlcxsjD = eddlcxsjStr.toDouble();
                eddlcxsj=QString::number(static_cast<int>(eddlcxsjD));
                if(eddlcxsj != "2" && eddlcxsj != "3" && eddlcxsj != "4")
                {
                    QString errRe="格式错误";
                    QString cont="2,3,4";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
            }
            break;
        }
        case 34://"断口数量"
        {
            rx.setPattern("^(2|1)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="1,2";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
                dksl=rx.cap(0);
            break;
        }
        case 35://"套管爬电距离"
        {
            QString tmp_tgpdjl=mExcelSheetData.at(row).at(col).toString();
            if(!tmp_tgpdjl.isEmpty())
            {
                rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
                int res=rx.indexIn(tmp_tgpdjl);
                if(res == -1)
                {
                    QString errRe="格式错误";
                    QString cont="数字";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                else
                {
                    QString tgpdjlStr=rx.cap(0);
                    double tgpdjlD = tgpdjlStr.toDouble();
                    if(static_cast<int>(tgpdjlD) != 0)
                        tgpdjl = QString::number(static_cast<int>(tgpdjlD));
                }
            }
            break;
        }
        case 36://"套管干弧距离"
        {
            QString tmp_tgghjl=mExcelSheetData.at(row).at(col).toString();
            if(!tmp_tgghjl.isEmpty())
            {
                rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
                int res=rx.indexIn(tmp_tgghjl);
                if(res == -1)
                {
                    QString errRe="格式错误";
                    QString cont="数字";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                else
                {
                    QString tgghjlStr=rx.cap(0);
                    double tgghjlD = tgghjlStr.toDouble();
                    if(static_cast<int>(tgghjlD) != 0)
                        tgghjl = QString::number(static_cast<int>(tgghjlD));
                }
            }
            break;
        }
        case 37://"机械寿命"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="数字";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                QString jxsmStr=rx.cap(0);
                double jxsmD = jxsmStr.toDouble();
                jxsm = QString::number(static_cast<int>(jxsmD));
            }
            break;
        }
        case 38://"合闸时间"
        {
            QString tmpHzsj=mExcelSheetData.at(row).at(col).toString();
            if(!tmpHzsj.isEmpty())
            {
                rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
                int res=rx.indexIn(tmpHzsj);
                if(res == -1)
                {
                    QString errRe="格式错误";
                    QString cont="数字：20～150";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                else
                {
                    hzsj=rx.cap(0);
                    if(hzsj.toInt() < 20 || hzsj.toInt() > 150)
                    {
                        QString errRe="格式错误";
                        QString cont="数字：20～150";
                        ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                        wrongFieldList.push_back(wrongField);
                        hzsj.clear();
                    }
                }
            }
            break;
        }
        case 39://"分闸时间"
        {
            QString tmpFzsj=mExcelSheetData.at(row).at(col).toString();
            if(!tmpFzsj.isEmpty())
            {
                rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
                int res=rx.indexIn(tmpFzsj);
                if(res == -1)
                {
                    QString errRe="格式错误";
                    QString cont="参考值：10～60";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                }
                else
                {
                    fzsj=rx.cap(0);
                    if(fzsj.toInt() < 10 || fzsj.toInt() > 60)
                    {
                        QString errRe="格式错误";
                        QString cont="参考值：10～60";
                        ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                        wrongFieldList.push_back(wrongField);
                        fzsj.clear();
                    }
                }
            }
            break;
        }
        case 40://"合分时间"
        {
            rx.setPattern("^((0\\.\\d+)|([1-9]\\d*.?[0-9]*)|0)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="参考值：0～100";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            else
            {
                hfsj=rx.cap(0);
                if(hfsj.toInt() <= 0 || hfsj.toInt() > 100)
                {
                    QString errRe="格式错误";
                    QString cont="参考值：0～100";
                    ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                    wrongFieldList.push_back(wrongField);
                    hfsj.clear();
                }
            }
            break;
        }
        case 43://"设备增加方式"
        {
            rx.setPattern("^(设备增加-基本建设|设备增加-技术改造|设备增加-零星购置|设备增加-无偿调入)$");
            int res=rx.indexIn(mExcelSheetData.at(row).at(col).toString());
            if(res == -1)
            {
                QString errRe="格式错误";
                QString cont="设备增加-基本建设,设备增加-技术改造,设备增加-零星购置,设备增加-无偿调入";
                ErrorInfo wrongField(col, fieldName, ruleError,errRe,cont);   //出错字段
                wrongFieldList.push_back(wrongField);
            }
            break;
        }
        default:
            break;
    }
    }

    /******额定短路开断电流、额定短路关合电流、动稳定电流、热稳定电流关系判断******/
    if(!eddlghdl.isEmpty() && !dwddl.isEmpty() && eddlghdl!=dwddl)
    {
        int index1=mExcelSheetData[0].indexOf("额定短路关合电流");
        int index2=mExcelSheetData[0].indexOf("动稳定电流");
        QString errRe="与动稳定电流关联错误";
        ErrorInfo wrongField1(index1, "额定短路关合电流", matchError,errRe);   //出错字段
        wrongFieldList.push_back(wrongField1);
        QString errRe1="与额定短路关合电流关联错误";
        ErrorInfo wrongField2(index2, "动稳定电流", matchError,errRe1);   //出错字段
        wrongFieldList.push_back(wrongField2);
    }
    if(!eddlkddl.isEmpty() && !rwddl.isEmpty() && eddlkddl!=rwddl)
    {
        int index1=mExcelSheetData[0].indexOf("额定短路开断电流");
        int index2=mExcelSheetData[0].indexOf("热稳定电流");
        QString errRe="与热稳定电流关联错误";
        ErrorInfo wrongField1(index1, "额定短路开断电流", matchError,errRe);   //出错字段
        wrongFieldList.push_back(wrongField1);
        QString errRe1="与额定短路开断电流关联错误";
        ErrorInfo wrongField2(index2, "热稳定电流", matchError,errRe1);   //出错字段
        wrongFieldList.push_back(wrongField2);
    }
    if(!dwddl.isEmpty() && !rwddl.isEmpty())
    {
        double dwdData=dwddl.toDouble();
        double rwdData=rwddl.toDouble();
        if(dwdData < 2.4*rwdData || dwdData > 2.6*rwdData)
        {
            int index1=mExcelSheetData[0].indexOf("动稳定电流");
            int index2=mExcelSheetData[0].indexOf("热稳定电流");
            QString errRe="与热稳定电流关联错误";
            ErrorInfo wrongField1(index1, "动稳定电流", matchError,errRe);   //出错字段
            wrongFieldList.push_back(wrongField1);
            QString errRe1="与动稳定电流关联错误";
            ErrorInfo wrongField2(index2, "热稳定电流", matchError,errRe1);   //出错字段
            wrongFieldList.push_back(wrongField2);
        }
    }
    /******额定短路持续时间判断******/
    if(!dydj.isEmpty() && !eddlcxsj.isEmpty())
    {
        if((dydj.toInt() == 500 && eddlcxsj != "2" && eddlcxsj != "3") ||
                (dydj.toInt() >= 66 && dydj.toInt() <= 220 && eddlcxsj != "3" && eddlcxsj != "4") ||
                (dydj.toInt() <= 35 && eddlcxsj != "4"))
        {
            int index1=mExcelSheetData[0].indexOf("额定短路持续时间");
            QString errRe="与电压等级关联错误";
            QString cont;
            if(dydj.toInt() == 500)
            {
                cont="2或3";
            }
            else if(dydj.toInt() >= 66 && dydj.toInt() <= 220)
            {
                cont="3或4";
            }
            else if(dydj.toInt() <= 35)
            {
                cont="4";
            }
            ErrorInfo wrongField1(index1, "额定短路持续时间", matchError,errRe,cont);   //出错字段
            wrongFieldList.push_back(wrongField1);
        }
    }
    /******断口数量关系判断******/
    if(!dydj.isEmpty() && !dksl.isEmpty())
    {
        if((dydj.toInt() == 500 && dksl != "2") || (dydj.toInt() == 220 && dksl != "1"  && dksl != "2")
                || (dydj.toInt() <= 110 && dksl != "1"))
        {
            int index1=mExcelSheetData[0].indexOf("断口数量");
            QString errRe="与电压等级关联错误";
            QString cont;
            if(dydj.toInt() == 500)
            {
                cont="2";
            }
            else if(dydj.toInt() == 220)
            {
                cont="1或2";
            }
            else if(dydj.toInt() <= 110)
            {
                cont="1";
            }
            ErrorInfo wrongField1(index1, "断口数量", matchError,errRe,cont);   //出错字段
            wrongFieldList.push_back(wrongField1);
        }
    }
    /******套管爬电距离判断******/
    if(!dydj.isEmpty() && !tgpdjl.isEmpty())
    {
        if((dydj.toInt() == 500 && (tgpdjl.toInt() < 13500 || tgpdjl.toInt() > 18000)) ||
                (dydj.toInt() == 220 && (tgpdjl.toInt() < 5000 || tgpdjl.toInt() > 8000)) ||
                (dydj.toInt() == 110 && (tgpdjl.toInt() < 2500 || tgpdjl.toInt() > 4800)) ||
                (dydj.toInt() == 66 && (tgpdjl.toInt() < 2000 || tgpdjl.toInt() > 2500)) ||
                (dydj.toInt() == 35 && (tgpdjl.toInt() < 800 || tgpdjl.toInt() > 2500)) ||
                ((dydj.toInt() == 10 || dydj.toInt() == 6) && (tgpdjl.toInt() < 200 || tgpdjl.toInt() > 400)))
        {
            int index1=mExcelSheetData[0].indexOf("套管爬电距离");
            QString errRe="与电压等级关联错误";
            QString cont;
            if(dydj.toInt() == 500)
            {
                cont="13500～18000";
            }
            else if(dydj.toInt() == 220)
            {
                cont="5000～8000";
            }
            else if(dydj.toInt() == 110)
            {
                cont="2500～4800";
            }
            else if(dydj.toInt() == 66)
            {
                cont="2000～2500";
            }
            else if(dydj.toInt() == 35)
            {
                cont="800～2500";
            }
            else if(dydj.toInt() == 10 || dydj.toInt() == 6)
            {
                cont="200～400";
            }
            ErrorInfo wrongField1(index1, "套管爬电距离", matchError,errRe,cont);   //出错字段
            wrongFieldList.push_back(wrongField1);
        }
    }
    /******套管干弧距离判断******/
    if(!dydj.isEmpty() && !tgghjl.isEmpty())
    {
        if((dydj.toInt() == 500 && (tgghjl.toInt() < 3500 || tgghjl.toInt() > 4000)) ||
                (dydj.toInt() == 220 && (tgghjl.toInt() < 1800 || tgghjl.toInt() > 2200)) ||
                (dydj.toInt() == 110 && (tgghjl.toInt() < 900 || tgghjl.toInt() > 1200)) ||
                ((dydj.toInt() == 66 || dydj.toInt() == 35) && (tgghjl.toInt() < 300 || tgghjl.toInt() > 600)) ||
                ((dydj.toInt() == 10 || dydj.toInt() == 6) && (tgghjl.toInt() < 180 || tgghjl.toInt() > 280)))
        {
            int index1=mExcelSheetData[0].indexOf("套管干弧距离");
            QString errRe="与电压等级关联错误";
            QString cont;
            if(dydj.toInt() == 500)
            {
                cont="3500～4000";
            }
            else if(dydj.toInt() == 220)
            {
                cont="1800～2200";
            }
            else if(dydj.toInt() == 110)
            {
                cont="900～1200";
            }
            else if(dydj.toInt() == 66 || dydj.toInt() == 35)
            {
                cont="300～600";
            }
            else if(dydj.toInt() == 10 || dydj.toInt() == 6)
            {
                cont="180～280";
            }
            ErrorInfo wrongField1(index1, "套管干弧距离", matchError,errRe,cont);   //出错字段
            wrongFieldList.push_back(wrongField1);
        }
    }
    /******机械寿命判断******/
    if(!mhjz.isEmpty() && !jxsm.isEmpty())
    {
        if((mhjz == "真空" && jxsm.toInt() != 5000 && jxsm.toInt() != 10000 && jxsm.toInt() != 20000 && jxsm.toInt() != 30000)
                || (mhjz == "SF6" && jxsm.toInt() != 2000 && jxsm.toInt() != 3000 && jxsm.toInt() != 5000 && jxsm.toInt() != 6000 && jxsm.toInt() != 10000)
                || (mhjz == "其他" && jxsm.toInt() != 2000 && jxsm.toInt() != 3000 && jxsm.toInt() != 5000))
        {
            int index1=mExcelSheetData[0].indexOf("机械寿命");
            QString errRe="与灭弧介质关联错误";
            QString cont;
            if(mhjz == "真空")
            {
                cont="5000,10000,20000,30000";
            }
            else if(mhjz == "SF6")
            {
                cont="2000,3000,5000,6000,10000";
            }
            else if(mhjz == "其他")
            {
                cont="2000,3000,5000";
            }
            ErrorInfo wrongField1(index1, "机械寿命", matchError,errRe,cont);   //出错字段
            wrongFieldList.push_back(wrongField1);
        }
    }
    /******合闸时间、分闸时间判断******/
    if(!hzsj.isEmpty() && !fzsj.isEmpty() && hzsj.toInt() < fzsj.toInt())
    {
        int index1=mExcelSheetData[0].indexOf("合闸时间");
        int index2=mExcelSheetData[0].indexOf("分闸时间");
        QString errRe="与分闸时间关联错误";
        QString cont="合闸时间大于分闸时间";
        ErrorInfo wrongField1(index1, "合闸时间", matchError,errRe,cont);   //出错字段
        wrongFieldList.push_back(wrongField1);
        QString errRe1="与合闸时间关联错误";
        QString cont1="合闸时间大于分闸时间";
        ErrorInfo wrongField2(index2, "分闸时间", matchError,errRe1,cont1);   //出错字段
        wrongFieldList.push_back(wrongField2);
    }
    /******合分时间判断******/
    if(!dydj.isEmpty() && !hfsj.isEmpty())
    {
        if((dydj.toInt() == 550 && (hfsj.toInt() <= 0 || hfsj.toInt() > 50)) ||
                ((dydj.toInt() == 220 || dydj.toInt() == 110) && (hfsj.toInt() <= 0 || hfsj.toInt() > 60)) ||
                (dydj.toInt() <= 66 && (hfsj.toInt() <= 0 || hfsj.toInt() > 100)))
        {
            int index1=mExcelSheetData[0].indexOf("合分时间");
            QString errRe="与电压等级关联错误";
            QString cont;
            if(dydj.toInt() == 500)
            {
                cont="0～50";
            }
            else if(dydj.toInt() == 220 || dydj.toInt() == 110)
            {
                cont="0～60";
            }
            else if(dydj.toInt() <= 66)
            {
                cont="0～100";
            }
            ErrorInfo wrongField1(index1, "合分时间", matchError,errRe,cont);   //出错字段
            wrongFieldList.push_back(wrongField1);
        }
    }
}

//by yang 2020.7.9
void EquipManage::showWrongItemExcelData()
{
    ui->textLabel_wrongItem->setText(QString::number(mWrongItems.size())+QString("条记录存在格式错误，或与对比库不符："));
    //设置文本颜色
    QPalette pa;
    pa.setColor(QPalette::WindowText,Qt::red);
    ui->textLabel_wrongItem->setPalette(pa);

    mTableModelOfWrongItem->clear();
    mTableModelOfWrongItem->setRowCount(mWrongItems.size());
    mTableModelOfWrongItem->setColumnCount(mColsOfSheet);
    QStringList headerList;
    for(int i=0; i<mExcelSheetData[0].size(); i++)
    {
        headerList.append(mExcelSheetData[0][i].toString());
    }
    headerList.removeFirst();
    headerList.push_front("序号");
    mTableModelOfWrongItem->setHorizontalHeaderLabels(headerList);

    QStandardItem *item;

    QMap<int, QList<ErrorInfo>>::const_iterator iter = mWrongItems.constBegin();
    int i = 0;
    while (iter != mWrongItems.constEnd())
    {
        int row = iter.key();       //行号
        QList<ErrorInfo> errorInfo = iter.value();
        QList<int> errorCol;        //该行中错误的字段列序号
        for(int k=0; k<errorInfo.size(); k++)
        {
            errorCol.append(errorInfo.at(k).num);
        }
        for(int j=0; j<mColsOfSheet; j++)
        {
            if(errorCol.contains(j))        //该字段为错误字段
            {
                int indexErr = errorCol.indexOf(j);
                QString errFldStr = mExcelSheetData[row][j].toString()+
                        "("+errorInfo[indexErr].errReason+";"+errorInfo[indexErr].recomContent+")";
                item = new QStandardItem(errFldStr);
                mTableModelOfWrongItem->setItem(i,j,item);

                //设置字体颜色
                mTableModelOfWrongItem->item(i, j)->setForeground(QBrush(QColor(255, 0, 0)));
                //字体加粗
                mTableModelOfWrongItem->item(i, j)->setFont(QFont("Times", 10, QFont::Black));
            }
            else            //正确字段
            {
                item = new QStandardItem(mExcelSheetData[row][j].toString());
                mTableModelOfWrongItem->setItem(i,j,item);
            }
        }
        i++;
        iter++;
    }
    ui->tableView_wrongItem->setModel(mTableModelOfWrongItem);
}


void EquipManage::showExcelData()
{
    int indexPro = mExcelSheetData[0].indexOf("生产厂家");
    int indexType = mExcelSheetData[0].indexOf("型号");
    QList<int> presentData;     //存储需要显示的行号
    QList<QList<int>> groupData;     //存储显示内容presentData中相同厂家、型号的组号，int：组号，QList<int>：该组中的行号

    //mMinorItems：mRedFidOfMinorItems集合中只展示1条，并标红互不相同字段
    QPair<QString, QString> proType;    //存储厂家和型号信息
    QMap<QPair<QString,QString>, GroupOfProType>::const_iterator iter;
    for(iter=mRedFidOfMinorItems.constBegin(); iter != mRedFidOfMinorItems.constEnd(); iter++)
    {
        QList<int> numList;     //组中的行号list
        QList<QSet<int>> rownumSet = iter.value().rownumSet;
        for(int i=0; i<rownumSet.size(); i++)
        {
            presentData.append(*rownumSet[i].begin());

            //by yang, 2020.7.18
            numList.append(*rownumSet[i].begin());
        }
        //by yang, 2020.7.18
        groupData.append(numList);
    }

//    presentData.append(mMinorItems);
//    QList<int> minorOfMajorItemsData = mMinorOfMajorItems.keys();
    //minorOfMajorItemsData厂家型号关联字段相同的只显示1条
    QList<int> uniqMinorOfMajorItemsData;
    for(int i=0; i<mSetOfMinorOfMajorItems.size(); i++)
    {
        uniqMinorOfMajorItemsData.append(mSetOfMinorOfMajorItems[i][0]);
    }
    presentData.append(uniqMinorOfMajorItemsData);

    // by yang 2020.7.19
    //对uniqMinorOfMajorItemsData中的行号进行相同厂家和型号分组
    QPair<QString,QString> producerAndtype;
    QList<int> rowGro;      //相同厂家和型号的行号集合
    for(int i=0; i<uniqMinorOfMajorItemsData.size(); i++)
    {
        int rowNum = uniqMinorOfMajorItemsData[i];
        QString pro = mExcelSheetData[rowNum][indexPro].toString();
        QString type = mExcelSheetData[rowNum][indexType].toString();
        QPair<QString,QString> proTypeTmp(pro, type);
        if(producerAndtype == proTypeTmp)
        {
            rowGro.append(rowNum);
        }
        else
        {
            producerAndtype = proTypeTmp;
            if(!rowGro.isEmpty())
            {
                groupData.append(rowGro);
            }
            rowGro.clear();
            rowGro.append(rowNum);
        }
        //将最后一组数据加入
        if(i == uniqMinorOfMajorItemsData.size()-1)
        {
            groupData.append(rowGro);
        }
    }



//    std::sort(presentData.begin(), presentData.end());

    mTableModel->clear();
    mTableModel->setRowCount(presentData.size());
    mTableModel->setColumnCount(mColsOfSheet+3);        //增加组号、错误概率、用户判断是否错误两列
    QStringList headerList;
    for(int i=0; i<mExcelSheetData[0].size(); i++)
    {
        //去掉表头中的换行符，避免冻结列时，表头高度不一致
        QString headerStr = mExcelSheetData[0][i].toString().remove(QString("\n"));
        headerList.append(headerStr);
    }
    headerList.removeFirst();
    headerList.push_front("序号");

//    headerList.push_back("综合判断及正确概率");
    headerList.push_front("综合判断及正确概率");
    headerList.push_front("组号");            //by yang 2020.7.19

    headerList.push_back("是否正确");
    mTableModel->setHorizontalHeaderLabels(headerList);

//    //样本库中厂家、型号相同数目较少的条目，以及尽管数目多但不满足存在2/3以上相关属性相同的条目中，
//    //与厂家型号关联字段相同的条目的条目数    //by yang  2020.7.8
//    QMap<int, int> minorItemsSize;      //key：int表示行号，value：int表示与行号的厂家型号关联字段相同的条目的条目数

    //设置tableView中的组号       by yang, 2020.7.19
    int currRow = 0;
    for(int k=0; k<groupData.size(); k++)
    {
        ui->tableView->setSpan(currRow, 0, groupData[k].size(), 1);
        QStandardItem* item = new QStandardItem(QString::number(k+1));
        mTableModel->setItem(currRow, 0, item);
        currRow += groupData[k].size();
    }

    QStandardItem *item;
    for(int i=0; i<presentData.size(); i++)
    {
        for(int j=0; j<mColsOfSheet; j++)
        {
            item = new QStandardItem(mExcelSheetData[presentData[i]][j].toString());
//            mTableModel->setItem(i,j,item);
            mTableModel->setItem(i,j+2,item);
        }
        if(mMinorItems.contains(presentData[i]))
        {
            QString pro = mExcelSheetData[presentData[i]][indexPro].toString();
            QString type = mExcelSheetData[presentData[i]][indexType].toString();
            QPair<QString,QString> protype(pro, type);
            QList<int> diffFild = mRedFidOfMinorItems[protype].diffFild;

//            //by yang 2020.7.8
//            QList<QSet<int>> numSetList = mRedFidOfMinorItems[protype].rownumSet;
//            int numSize = 0;        //与presentData[i]行号，相同的所有行数
//            for (int ii = 0; ii < numSetList.size(); ++ii)
//            {
//                if(numSetList.at(ii).contains(presentData[i]))
//                {
//                    numSize = numSetList.at(ii).size();
//                    break;
//                }
//            }
//            minorItemsSize.insert(presentData[i], numSize);

            //设置可能出错字段为红色字体
            for(int k=0; k<diffFild.size(); k++)
            {
//                mTableModel->item(i, diffFild[k])->setForeground(QBrush(QColor(255, 0, 0)));
                mTableModel->item(i, diffFild[k]+2)->setForeground(QBrush(QColor(255, 0, 0)));
                //字体加粗
//                mTableModel->item(i, diffFild[k])->setFont(QFont("Times", 10, QFont::Black));
                mTableModel->item(i, diffFild[k]+2)->setFont(QFont("Times", 10, QFont::Black));
            }
        }
        else if(mMinorOfMajorItems.contains(presentData[i]))
        {
            QList<int> col=mMinorOfMajorItems[presentData[i]];
            //设置可能出错字段为红色字体
            for(int k=0; k<col.size(); k++)
            {
//                mTableModel->item(i, col[k])->setForeground(QBrush(QColor(255, 0, 0)));
                mTableModel->item(i, col[k]+2)->setForeground(QBrush(QColor(255, 0, 0)));
                //字体加粗
//                mTableModel->item(i, col[k])->setFont(QFont("Times", 10, QFont::Black));
                mTableModel->item(i, col[k]+2)->setFont(QFont("Times", 10, QFont::Black));
            }
        }

        //添加概率
        if(corPro.contains(presentData[i]))
        {
            QString correctPro = QString("正确率：") + corPro[presentData[i]];
            item = new QStandardItem(correctPro);
        }
        else
        {
//            //by yang 2020.7.8
//            int numSize = minorItemsSize.value(presentData[i]);
//            if(numSize < 3)
//            {
//                item = new QStandardItem(QString("样本少(")+QString::number(numSize)+QString(")，请单独核查主要参数！"));
//            }
//            else
//            {
//                item = new QStandardItem(QString("样本少(")+QString::number(numSize)+QString(")，请对比核查不一致参数！"));
//            }

            //by yang 2020.7.18
            if(mFewOfMinorItems.contains(presentData[i]))
            {
                int sampleNum = mFewOfMinorItems.value(presentData[i]);
                item = new QStandardItem(QString("样本少(")+QString::number(sampleNum)+QString(")，请单独核查主要参数！"));
            }
            else if(mNotExistSameFldsOfMinorItems.contains(presentData[i]))
            {
                int sampleNum = mNotExistSameFldsOfMinorItems.value(presentData[i]);
                item = new QStandardItem(QString("样本少(")+QString::number(sampleNum)+QString(")，请对比核查不一致参数！"));
            }
            else
            {
                qDebug()<<"ERROR: mFewOfMinorItems And mNotExistSameFldsOfMinorItems No Record!";
            }
        }
//        mTableModel->setItem(i,mColsOfSheet,item);
        mTableModel->setItem(i,1,item);
        //添加是否错误判断
        item = new QStandardItem();
        item->setCheckable(true);
        item->setCheckState(Qt::Unchecked);
        mTableModel->setItem(i,mColsOfSheet+2,item);
    }
    ui->tableView->setModel(mTableModel);

    //显示格式错误、与对比库不同的条目信息    by yang 2020.7.10
    showWrongItemExcelData();
}

void EquipManage::clear()
{
    mExcelSheetData.clear();
    mWrongItems.clear();
    mCorrectItems.clear();
    mManuWrongItems.clear();

    mWriteDbItems.clear();
    mRecogProduAndTypeItems.clear();
    mNoRecOfDbItems.clear();
    mUniqNoRecOfDbItems.clear();
    mMinorItems.clear();
    mFewOfMinorItems.clear();
    mNotExistSameFldsOfMinorItems.clear();
    mRedFidOfMinorItems.clear();
    mMinorOfMajorItems.clear();
    mSetOfMinorOfMajorItems.clear();
    corPro.clear();
    mMajorOfMajorItems.clear();
    mColOfProAndType.clear();
}

//通过LibXl读取excel（支持xlsx，xls）
void EquipManage::readAllCellsByLibXL(QString path)
{
    wstring pathW = path.toStdWString();
    const wchar_t *pathWc = pathW.c_str();

    Book * book = NULL;
    QRegExp rx;     //正则表达式
    rx.setPattern(".xlsx$");
    int res=rx.indexIn(path);
    if(res != -1)       //xlsx 文件
    {
        book = xlCreateXMLBook();     //打开 xlsx 文件
    }
    else        //xls文件
    {
        book = xlCreateBook();        //打开 xls 文件
    }
    book->setKey(NAME_LIBXL, KEY_LIBXL);

    bool  xmload = book->load(pathWc);
    if(xmload == false)
    {
        qDebug()<<book->errorMessage();
        QMessageBox::critical(0, QObject::tr("Open Xlsx Error"), QString(book->errorMessage()));
        return;
    }
    Sheet * sheet = book->getSheet(0);

    if(sheet)
    {
        int rowfirst = sheet->firstRow();
        int rowlast = sheet->lastRow();
        int colfirst = sheet->firstCol();
        int collast = sheet->lastCol();

    //    wcout << L"数据开始行 ：" << rowfirst << endl;
    //    wcout << L"数据结束行 ：" << rowlast << endl;
    //    wcout << L"数据开始列 ：" << colfirst << endl;
    //    wcout << L"数据结束列 ：" << collast << endl;

        cout << rowfirst << endl;
        cout << rowlast << endl;
        cout << colfirst << endl;
        cout << collast << endl;

        mRowsOfSheet = rowlast - rowfirst;      //行数
        mColsOfSheet = collast - colfirst;      //列数
        mStartRowOfSheet = rowfirst;            //启始行
        mStartColOfSheet = colfirst;            //启始列

        for (int i = rowfirst; i < rowlast; ++i)
        {
            QList<QVariant> valueList;
            for (int j = colfirst; j < collast; j++)
            {
//                QString value;
//                //处理日期格式
//                if(sheet->isDate(i, j))
//                {
//                    int year, month, day;
//                    book->dateUnpack(sheet->readNum(i, j), &year, &month, &day);
//                    value = QDateTime(QDate(year, month, day)).toString("yyyy/MM/dd");
//                }
//                else
//                {
//                    const wchar_t *valueWC = sheet->readStr(i, j);
//                    if(valueWC)
//                    {
//                        value = QString::fromWCharArray(valueWC);

//                    }
//                    else
//                    {
//                        value = QString::number(sheet->readNum(i, j));
//    //                    qDebug()<<book->errorMessage();
//                    }

//                }
//                valueList.append(value);


                QString value;
                CellType cellType = sheet->cellType(i, j);
                if(sheet->isFormula(i, j))
                {
                   const wchar_t* s = sheet->readFormula(i, j);
                   std::wcout << "("<<i<<","<<j<<")"<< (s ? s : L"null") << " [formula]";
                }
                else
                {
                   switch(cellType)
                   {
                      case CELLTYPE_EMPTY: std::wcout << "("<<i<<","<<j<<")"<< "[empty]"; break;
                      case CELLTYPE_NUMBER:
                      {
                       if(sheet->isDate(i, j))      //处理日期格式
                       {
                           int year, month, day;
                           book->dateUnpack(sheet->readNum(i, j), &year, &month, &day);
                           value = QDateTime(QDate(year, month, day)).toString("yyyy/MM/dd");
                       }
                       else
                       {
                           value = QString::number(sheet->readNum(i, j));
                       }
                       break;
                      }
                      case CELLTYPE_STRING:
                      {
                       const wchar_t *valueWC = sheet->readStr(i, j);
                       value = QString::fromWCharArray(valueWC);
                       break;
                      }
                      case CELLTYPE_BOOLEAN:
                      {
                         bool b = sheet->readBool(i, j);
                         std::wcout << "("<<i<<","<<j<<")"<< (b ? "true" : "false") << " [boolean]";
                         break;
                      }
                      case CELLTYPE_BLANK: std::wcout << "("<<i<<","<<j<<")"<< "[blank]"; break;
                      case CELLTYPE_ERROR: std::wcout << "("<<i<<","<<j<<")"<< "[error]"; break;
                   }
                }
                valueList.append(value);
            }
            mExcelSheetData.append(valueList);
        }
    }
    else
    {
        qDebug()<<book->errorMessage();
        return;
    }
    book->release();
}

//通过QT xlsx读取exel（只支持xlsx，不支持xls）
void EquipManage::readAllCellsByXlsx(QString path)
{
    QXlsx::Document xlsx(path);
    QXlsx::Workbook *workBook = xlsx.workbook();
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    mRowsOfSheet=workSheet->dimension().rowCount(); //行数
    mColsOfSheet=workSheet->dimension().columnCount(); //列数
    mStartRowOfSheet=workSheet->dimension().firstRow(); //启始行
    mStartColOfSheet=workSheet->dimension().firstColumn();  //启始列
    for (int i = 1; i <= mRowsOfSheet; i++)
    {
        QList<QVariant> valueList;
         for (int j = 1; j <= mColsOfSheet; j++)
         {
             QVariant value;
             QXlsx::Cell *cell = workSheet->cellAt(i, j);
             if(cell != NULL)
             {
                 if (cell->isDateTime())//日期
                 {
                     if (cell->dateTime().date().year()==1899) continue;
                     value = cell->dateTime().toString("yyyy/MM/dd");
                 }
                 else
                 {
                     value = cell->value();
                 }
             }
             valueList.append(value);
         }
         mExcelSheetData.append(valueList);
    }
}


void EquipManage::on_importAction_triggered()
{
    clear();

    mValueOfProgress=0;
    mFilePath = QFileDialog::getOpenFileName(this, "", ".", "*.xls *.xlsx");
    if(mFilePath == NULL)   return;

//    QWaiting *pQwait = new QWaiting(this);
//    pQwait->show();

    QCoreApplication::processEvents();

#if XLSX
    readAllCellsByXlsx(mFilePath);
#elif LIBXL
    readAllCellsByLibXL(mFilePath);
#else

    ExcelHandle *excelHandle = new ExcelHandle(mFilePath);
    excelHandle->readRowsAndColsOfSheet(SHEET_NUM, mStartRowOfSheet, mRowsOfSheet, mStartColOfSheet, mColsOfSheet); 
#endif
//    pQwait->close();
    //等待进度条
    QCoreApplication::processEvents();
    QApplication::setOverrideCursor(Qt::WaitCursor);//设置鼠标为等待状态
    mProgress = new QProgressDialog(this);
    mProgress->setFixedSize(300,100);
    mProgress->setCancelButton(NULL);
    mProgress->setWindowTitle("提示");
    mProgress->setLabelText("正在进行识别...");
    mProgress->setRange(0, mRowsOfSheet+8);//设置范围
    mProgress->setModal(true);//设置为模态对话框
    mProgress->show();
    mValueOfProgress++;
    mProgress->setValue(mValueOfProgress);
    QCoreApplication::processEvents();

#ifndef XLSX
    excelHandle->readSheetAllCells(SHEET_NUM, mExcelSheetData);
    delete excelHandle;
#endif

//    QMessageBox::information(this, "提示", "导入完成",
//                             QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);




    QCoreApplication::processEvents();
    mValueOfProgress++;
    mProgress->setValue(mValueOfProgress);

    judgeDuanluqi();

    QCoreApplication::processEvents();
    mValueOfProgress++;
    mProgress->setValue(mValueOfProgress);

    showExcelData();

    QCoreApplication::processEvents();
    mValueOfProgress++;
    mProgress->setValue(mValueOfProgress);
    mProgress->close();
    delete mProgress;
    QApplication::restoreOverrideCursor();


    int errorNum = mTableModel->rowCount();
    QString text = QString::number(errorNum)+"条记录可能存在错误，请人工判断！";
    ui->textLabel->setText(text);
//    //设置字号
//    QFont ft;
//    ft.setPointSize(20);
//    ui->textLabel->setFont(ft);
    //设置颜色
    QPalette pa;
    pa.setColor(QPalette::WindowText,Qt::red);
    ui->textLabel->setPalette(pa);


}

void EquipManage::on_exportAction_triggered()
{
    QString fileName = QFileDialog::getSaveFileName(this, "", ".", "*.xls *.xlsx");
    if(fileName.isNull()) return;

    mFileExportPath = fileName;     //获取导出样本库路径

    QApplication::setOverrideCursor(Qt::WaitCursor);//设置鼠标为等待状态
    mProgress = new QProgressDialog(this);
    mProgress->setFixedSize(300,100);
    mProgress->setCancelButton(NULL);
    mProgress->setWindowTitle("提示");
    mProgress->setLabelText("正在导出样本库...");
    mProgress->setRange(0, mRowsOfSheet+6);//设置范围
    mProgress->setModal(true);//设置为模态对话框
    mProgress->show();
    mValueOfProgress = 1;
    mProgress->setValue(mValueOfProgress++);
    QCoreApplication::processEvents();


    //获取人工判断结果并处理
    for(int i=0; i<mTableModel->rowCount(); i++)
    {
        if(mTableModel->item(i, mColsOfSheet+1)->checkState() == Qt::Unchecked)     //人工判断为错误
        {
            //加入人工判断为错误list
//            QModelIndex indexM = mTableModel->index(i, 0);
            QModelIndex indexM = mTableModel->index(i, 2);
            mManuWrongItems.push_back(indexM.data().toInt());
        }
        else        //人工判断为正确
        {
//            int baseNum= mTableModel->index(i, 0).data().toInt();   //人工判断正确的行号
            int baseNum= mTableModel->index(i, 2).data().toInt();   //人工判断正确的行号
            //判断该条目的厂家、型号信息是否在mMajorOfMajorItems中
            int proNum = mExcelSheetData[0].indexOf("生产厂家");
            int typeNum = mExcelSheetData[0].indexOf("型号");
//            QString pro = mTableModel->index(i, proNum).data().toString();
//            QString ty = mTableModel->index(i, typeNum).data().toString();
            QString pro = mTableModel->index(i, proNum+2).data().toString();        //by yang,2020.7.19
            QString ty = mTableModel->index(i, typeNum+2).data().toString();

            ProducerTypeInfo proType(pro, ty);          //厂家型号信息
            QList<int> keysMajor = mMajorOfMajorItems.keys(proType);    //与人工判断为正确的条目具有相同厂家型号信息的条目行号
            if(!keysMajor.isEmpty())
            {
                for(int j1=0; j1<keysMajor.size(); j1++)
                {
                    QList<ErrorInfo> errorList;
                    for(int j2=0; j2<mColOfProAndType.size(); j2++)
                    {
                        QString base=mExcelSheetData[baseNum][mColOfProAndType[j2]].toString();
                        QString conStr=mExcelSheetData[keysMajor[j1]][mColOfProAndType[j2]].toString();
                        if(base != conStr)  //字段出错
                        {
                            QString fieName = mExcelSheetData[0][mColOfProAndType[j2]].toString();
                            ErrorInfo errInfo(mColOfProAndType[j2], fieName, matchError);
                            errorList.push_back(errInfo);
                        }
                    }
                    mWrongItems.insert(keysMajor[j1], errorList);
                    //mMajorOfMajorItems删除keysMajor对应的条目
                    mMajorOfMajorItems.remove(keysMajor[j1]);
                }
            }
            //加入待写入对比库list
            if(mWriteDbItems.keys(proType).isEmpty())
            {
                mWriteDbItems.insert(baseNum, proType);
            }
        }
        QCoreApplication::processEvents();
    }

    mProgress->setValue(mValueOfProgress++);

    //mMajorOfMajorItems加入待写入对比库List
    QMap<int, ProducerTypeInfo>::const_iterator i = mMajorOfMajorItems.constBegin();
    while(i != mMajorOfMajorItems.constEnd())
    {
        ProducerTypeInfo proAndTy = i.value();
        if(mWriteDbItems.keys(proAndTy).isEmpty())
        {
            mWriteDbItems.insert(i.key(), proAndTy);
        }
        i++;
        QCoreApplication::processEvents();
    }

    mProgress->setValue(mValueOfProgress++);

    //写入对比库
    writeDataToDB();

    QCoreApplication::processEvents();
    mProgress->setValue(mValueOfProgress++);

    //导出excel
//    exportExcelData();
    exportExcelDataByLibXL();


    qDebug()<<"6666";



    QApplication::restoreOverrideCursor();
    mProgress->close();
    delete mProgress;

    qDebug()<<"7777";

    ui->textLabel->setText("");

    qDebug()<<"8888";
}


void EquipManage::writeDataToDB()
{
//    //获取数据库字段名称
//    QSqlQuery query1(mDuanDB);
//    QString strTableNmae = "compareLib_DuanLuQi";
//    QString str = "PRAGMA table_info(" + strTableNmae + ")";
//    query1.prepare(str);
//    if (query1.exec())
//    {
//        while (query1.next())
//        {
//            qDebug() << QString(QString("字段数:%1     字段名:%2     字段类型:%3")).
//                        arg(query1.value(0).toString()).arg(query1.value(1).toString()).arg(query1.value(2).toString());
//        }
//    }else{
//        qDebug() << query1.lastError();
//    }


    //数据库中字段顺序
    QList<QPair<QString,QString>> fieldNameList;
    fieldNameList.push_back(qMakePair(QString("id"), QString("ID")));
    fieldNameList.push_back(qMakePair(QString("ssds"), QString("所属地市")));
    fieldNameList.push_back(qMakePair(QString("sbmc"), QString("设备名称")));
    fieldNameList.push_back(qMakePair(QString("yxbh"), QString("运行编号")));
    fieldNameList.push_back(qMakePair(QString("ssdz"), QString("所属电站")));
    fieldNameList.push_back(qMakePair(QString("jgdy"), QString("间隔单元")));
    fieldNameList.push_back(qMakePair(QString("ywdw"), QString("运维单位")));
    fieldNameList.push_back(qMakePair(QString("whbz"), QString("维护班组")));
    fieldNameList.push_back(qMakePair(QString("dydj"), QString("电压等级")));
    fieldNameList.push_back(qMakePair(QString("sbzt"), QString("设备状态")));
    fieldNameList.push_back(qMakePair(QString("xs"), QString("相数")));
    fieldNameList.push_back(qMakePair(QString("xb"), QString("相别")));
    fieldNameList.push_back(qMakePair(QString("tyrq"), QString("投运日期")));
    fieldNameList.push_back(qMakePair(QString("zhsblx"), QString("组合设备类型")));
    fieldNameList.push_back(qMakePair(QString("sfnw"), QString("是否农网")));
    fieldNameList.push_back(qMakePair(QString("syhj"), QString("使用环境")));
    fieldNameList.push_back(qMakePair(QString("zyfl"), QString("专业分类")));
    fieldNameList.push_back(qMakePair(QString("xh"), QString("型号")));
    fieldNameList.push_back(qMakePair(QString("jgxs"), QString("结构型式")));
    fieldNameList.push_back(qMakePair(QString("czjgxs"), QString("操作机构型式")));
    fieldNameList.push_back(qMakePair(QString("mhjz"), QString("灭弧介质")));
    fieldNameList.push_back(qMakePair(QString("sccj"), QString("生产厂家")));
    fieldNameList.push_back(qMakePair(QString("ccrq"), QString("出厂日期")));
    fieldNameList.push_back(qMakePair(QString("eddy"), QString("额定电压")));
    fieldNameList.push_back(qMakePair(QString("eddl"), QString("额定电流")));
    fieldNameList.push_back(qMakePair(QString("edpl"), QString("额定频率")));
    fieldNameList.push_back(qMakePair(QString("edjysp"), QString("额定绝缘水平")));
    fieldNameList.push_back(qMakePair(QString("eddldlkdcs"), QString("额定短路电流开断次数")));
    fieldNameList.push_back(qMakePair(QString("eddlkddl"), QString("额定短路开断电流")));
    fieldNameList.push_back(qMakePair(QString("eddlghdl"), QString("额定短路关合电流")));
    fieldNameList.push_back(qMakePair(QString("dwddl"), QString("动稳定电流")));
    fieldNameList.push_back(qMakePair(QString("rwddl"), QString("热稳定电流")));
    fieldNameList.push_back(qMakePair(QString("eddlcxsj"), QString("额定短路持续时间")));
    fieldNameList.push_back(qMakePair(QString("dksl"), QString("断口数量")));
    fieldNameList.push_back(qMakePair(QString("tgpdjl"), QString("套管爬电距离")));
    fieldNameList.push_back(qMakePair(QString("tgghjl"), QString("套管干弧距离")));
    fieldNameList.push_back(qMakePair(QString("jxsm"), QString("机械寿命")));
    fieldNameList.push_back(qMakePair(QString("hzsj"), QString("合闸时间")));
    fieldNameList.push_back(qMakePair(QString("fzsj"), QString("分闸时间")));
    fieldNameList.push_back(qMakePair(QString("hfsj"), QString("合分时间")));
    fieldNameList.push_back(qMakePair(QString("zcxz"), QString("资产性质")));
    fieldNameList.push_back(qMakePair(QString("zcdw"), QString("资产单位")));
    fieldNameList.push_back(qMakePair(QString("sbzjfs"), QString("设备增加方式")));
    fieldNameList.push_back(qMakePair(QString("djsj"), QString("登记时间")));

    QMap<int, ProducerTypeInfo>::const_iterator i = mWriteDbItems.constBegin();
    while(i != mWriteDbItems.constEnd())
    {
        int num=i.key();
        int countRec=0;
        ProducerTypeInfo proType = i.value();
        QString sqlStr1="select count(*) from compareLib_DuanLuQi where sccj='"+proType.producer+"' and xh='"+proType.type+"'";
        QSqlQuery query(mDuanDB);
        query.exec(sqlStr1);
        while(query.next())
        {
            countRec=query.value(0).toInt();
        }
        if(countRec == 0)
        {
            QDateTime dateTime(QDateTime::currentDateTime());
            QString dateStr = dateTime.toString("yyyyMMddhhmmsszzz");
            QString sqlStr2="insert into compareLib_DuanLuQi values(:id,:ssds,:sbmc,:yxbh,:ssdz,:jgdy,:ywdw,:whbz,:dydj,:sbzt,:xs,"
                            ":xb,:tyrq,:zhsblx,:sfnw,:syhj,:zyfl,:xh,:jgxs,:czjgxs,:mhjz,:sccj,:ccrq,:eddy,:eddl,:edpl,:edjysp,"
                            ":eddldlkdcs,:eddlkddl,:eddlghdl,:dwddl,:rwddl,:eddlcxsj,:dksl,:tgpdjl,:tgghjl,:jxsm,:hzsj,:fzsj,:hfsj,"
                            ":zcxz,:zcdw,:sbzjfs,:djsj)";

            QSqlQuery query2(mDuanDB);
            query2.prepare(sqlStr2);
            query2.bindValue(":id", dateStr);
            for(int j=1; j<fieldNameList.size(); j++)
            {
                int index=mExcelSheetData[0].indexOf(fieldNameList[j].second);
                QString data = mExcelSheetData[num][index].toString();
                QString fieldNum = ":"+fieldNameList[j].first;
                query2.bindValue(fieldNum, data);
            }
            if(!query2.exec())
            {
                qDebug()<<query2.lastError();
                qDebug()<<"insert DB ERROR!";
            }
        }
        i++;
    }
}

void EquipManage::copyExcelDataByLibXL(Book* srcBook, Book* dstBook, const wchar_t *pathDes)
{
    Sheet* srcSheet = srcBook->getSheet(0);

    Sheet* dstSheet = dstBook->addSheet(L"my");

    // setting column widths
    for(int col = srcSheet->firstCol(); col < srcSheet->lastCol(); ++col)
    {
        dstSheet->setCol(col, col, srcSheet->colWidth(col), 0, srcSheet->colHidden(col));
    }

    std::map<Format*, Format*> formats;

    for(int row = srcSheet->firstRow(); row < srcSheet->lastRow(); ++row)
    {
        // setting row heights
        dstSheet->setRow(row, srcSheet->rowHeight(row), 0, srcSheet->rowHidden(row));

        for(int col = srcSheet->firstCol(); col < srcSheet->lastCol(); ++col)
        {
            // copying merging blocks
            int rowFirst, rowLast, colFirst, colLast;
            if(srcSheet->getMerge(row, col, &rowFirst, &rowLast, &colFirst, &colLast))
            {
                dstSheet->setMerge(rowFirst, rowLast, colFirst, colLast);
            }

            // copying formats
            Format *srcFormat, *dstFormat;

            srcFormat = srcSheet->cellFormat(row, col);
            if(!srcFormat) continue;

            // checking formats
            if(formats.count(srcFormat) == 0)
            {
                // format is not found, creating it in the output file
                dstFormat = dstBook->addFormat(srcFormat);
                formats[srcFormat] = dstFormat;
            }
            else
            {
                // format was already created
                dstFormat = formats[srcFormat];
            }

            // copying cell's values
            CellType ct = srcSheet->cellType(row, col);
            switch(ct)
            {
                case CELLTYPE_NUMBER:
                {
                    double value = srcSheet->readNum(row, col, &srcFormat);
                    dstSheet->writeNum(row, col, value, dstFormat);
                    break;
                }
                case CELLTYPE_BOOLEAN:
                {
                    qDebug()<<"("<<row<<","<<col<<"): BOOL";
                    bool value = srcSheet->readBool(row, col, &srcFormat);
                    dstSheet->writeBool(row, col, value, dstFormat);
                    break;
                }
                case CELLTYPE_STRING:
                {
                    const wchar_t* s = srcSheet->readStr(row, col, &srcFormat);
                    dstSheet->writeStr(row, col, s, dstFormat);
                    break;
                }
                case CELLTYPE_BLANK:
                {
                    qDebug()<<"("<<row<<","<<col<<"): BLANK";
                    srcSheet->readBlank(row, col, &srcFormat);
                    dstSheet->writeBlank(row, col, dstFormat);
                    break;
                }
            }
        }
    }

    dstBook->save(pathDes);

}




void EquipManage::exportExcelDataByLibXL()
{
    /*************复制原excel文件**************/
    //  若原excel与导出excel类型相同（同为xls或xlsx），则直接另存为
    //  若二者类型不同，则需复制原excel到导出excel

    wstring pathW = mFilePath.toStdWString();
    const wchar_t *pathWc = pathW.c_str();      //样本库路径
    wstring pathExW = mFileExportPath.toStdWString();
    const wchar_t *pathExWc = pathExW.c_str();  //导出库路径
    Book* bookSrc = NULL;
    Book* bookDes = NULL;

    QRegExp rx;     //正则表达式
    rx.setPattern(".xlsx$");
    int res1 = rx.indexIn(mFilePath);
    int res2 = rx.indexIn(mFileExportPath);
    if(res1 != -1)      //原excel为xlsx
    {
        bookSrc = xlCreateXMLBook();
        bookSrc->setKey(NAME_LIBXL, KEY_LIBXL);
        if(!bookSrc->load(pathWc))
        {
            qDebug()<<"bookSrc load xlsx ERROR: "<<bookSrc->errorMessage();
            return;
        }

        if(res2 != -1)      //导出excel为xlsx
        {
            if(!bookSrc->save(pathExWc))
            {
                qDebug()<<"bookSrc save xlsx ERROR: "<<bookSrc->errorMessage();
                return;
            }
            bookDes = xlCreateXMLBook();
            bookDes->setKey(NAME_LIBXL, KEY_LIBXL);
            if(!bookDes->load(pathExWc))
            {
                qDebug()<<"bookDes load xlsx ERROR: "<<bookSrc->errorMessage();
                return;
            }
        }
        else                //导出excel为xls
        {
            bookDes = xlCreateBook();
            bookDes->setKey(NAME_LIBXL, KEY_LIBXL);

            copyExcelDataByLibXL(bookSrc, bookDes, pathExWc);
        }
    }
    else                //原excel为xls
    {
        bookSrc = xlCreateBook();
        bookSrc->setKey(NAME_LIBXL, KEY_LIBXL);
        if(!bookSrc->load(pathWc))
        {
            qDebug()<<"bookSrc load xls ERROR: "<<bookSrc->errorMessage();
            return;
        }

        if(res2 == -1)      //导出excel为xls
        {
            if(!bookSrc->save(pathExWc))
            {
                qDebug()<<"bookSrc save xls ERROR: "<<bookSrc->errorMessage();
                return;
            }
            bookDes = xlCreateBook();
            bookDes->setKey(NAME_LIBXL, KEY_LIBXL);
            if(!bookDes->load(pathExWc))
            {
                qDebug()<<"bookDes load xls ERROR: "<<bookSrc->errorMessage();
                return;
            }
        }
        else                //导出excel为xlsx
        {
            bookDes = xlCreateXMLBook();
            bookDes->setKey(NAME_LIBXL, KEY_LIBXL);

            copyExcelDataByLibXL(bookSrc, bookDes, pathExWc);
        }
    }
    if(bookSrc != NULL) bookSrc->release();

    Sheet* desSheet = bookDes->getSheet(0);
    Format *cellFormat = NULL;

    QCoreApplication::processEvents();

    /******写入excel相关错误信息******/
//    QXlsx::Document xlsx(filePth);

    //mWrongItems
    QMap<int, QList<ErrorInfo>>::ConstIterator i=mWrongItems.constBegin();
    while(i != mWrongItems.constEnd())
    {
        int num=i.key();
        //将错误行号的序号标红
        CellType rowNumCell = desSheet->cellType(num, 0);
        switch(rowNumCell)
        {
            case CELLTYPE_NUMBER:
            {
                double v = desSheet->readNum(num, 0, &cellFormat);
                QString cellStr = QString::number(v);
                wstring cellStrW = cellStr.toStdWString();
                const wchar_t *cellStrWc = cellStrW.c_str();
                Format *desFormat = bookDes->addFormat(cellFormat);
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, cellStrWc, desFormat);
                break;
            }
            case CELLTYPE_BOOLEAN:
            {
                qDebug()<<"ERROR: BOOL!";
                bool v = desSheet->readBool(num, 0, &cellFormat);
                desSheet->writeBool(num, 0, v, cellFormat);
                break;
            }
            case CELLTYPE_STRING:
            {
                const wchar_t* s = desSheet->readStr(num, 0, &cellFormat);
                Format *desFormat = bookDes->addFormat(cellFormat);
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, s, desFormat);
                break;
            }
            case CELLTYPE_BLANK:
            {
                desSheet->readBlank(num, 0, &cellFormat);
                Format *desFormat = bookDes->addFormat(cellFormat);
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, QString("").toStdWString().c_str(), desFormat);
                break;
            }
            case CELLTYPE_EMPTY:
            {
                Format *desFormat = bookDes->addFormat();
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, QString("").toStdWString().c_str(), desFormat);
                break;
            }
            case CELLTYPE_ERROR:
            {
                qDebug()<<"ERROR: CELLTYPE_ERROR!";
                break;
            }
        }


        QList<ErrorInfo> value=i.value();
        for(int j=0; j<value.size(); j++)
        {
            CellType ct = desSheet->cellType(num, value[j].num);
            switch(ct)
            {
                case CELLTYPE_NUMBER:
                {
                    double v = desSheet->readNum(num, value[j].num, &cellFormat);
                    QString cellStr = QString::number(v)+"("+value[j].errReason+";"+value[j].recomContent+")";
                    wstring cellStrW = cellStr.toStdWString();
                    const wchar_t *cellStrWc = cellStrW.c_str();
                    Format *desFormat = bookDes->addFormat(cellFormat);
                    desFormat->setFillPattern(FILLPATTERN_SOLID);
                    desFormat->setPatternForegroundColor(COLOR_RED);
                    desSheet->writeStr(num, value[j].num, cellStrWc, desFormat);
 //                    desSheet->writeNum(num, value[j].num, v, cellFormat);
                    break;
                }
                case CELLTYPE_BOOLEAN:
                {
                    qDebug()<<"ERROR: BOOL!";
                    bool v = desSheet->readBool(num, value[j].num, &cellFormat);
                    desSheet->writeBool(num, value[j].num, v, cellFormat);
                    break;
                }
                case CELLTYPE_STRING:
                {
                    const wchar_t* s = desSheet->readStr(num, value[j].num, &cellFormat);
                    QString errStr = "("+value[j].errReason+";"+value[j].recomContent+")";
                    wstring errStrW = errStr.toStdWString();
                    const wchar_t *errStrWc = errStrW.c_str();
                    wchar_t cellStr[3000] = {0};
                    wcscpy (cellStr, s);
                    wcscat (cellStr, errStrWc);

//                    if(wcslen(cellStr) > 200)
//                        qDebug()<<"("<<num<<","<<value[j].num<<")"<<wcslen(cellStr);
                    Format *desFormat = bookDes->addFormat(cellFormat);
                    desFormat->setFillPattern(FILLPATTERN_SOLID);
                    desFormat->setPatternForegroundColor(COLOR_RED);
                    desSheet->writeStr(num, value[j].num, cellStr, desFormat);
                    break;
                }
                case CELLTYPE_BLANK:
                {
                    desSheet->readBlank(num, value[j].num, &cellFormat);
                    QString errStr = "("+value[j].errReason+";"+value[j].recomContent+")";
                    wstring errStrW = errStr.toStdWString();
                    const wchar_t *errStrWc = errStrW.c_str();
                    Format *desFormat = bookDes->addFormat(cellFormat);
                    desFormat->setFillPattern(FILLPATTERN_SOLID);
                    desFormat->setPatternForegroundColor(COLOR_RED);
                    desSheet->writeStr(num, value[j].num, errStrWc, desFormat);
 //                    desSheet->writeBlank(num, value[j].num, cellFormat);
                    break;
                }
                case CELLTYPE_EMPTY:
                {
                    QString errStr = "("+value[j].errReason+";"+value[j].recomContent+")";
                    wstring errStrW = errStr.toStdWString();
                    const wchar_t *errStrWc = errStrW.c_str();
                    Format *desFormat = bookDes->addFormat();
                    desFormat->setFillPattern(FILLPATTERN_SOLID);
                    desFormat->setPatternForegroundColor(COLOR_RED);
                    desSheet->writeStr(num, value[j].num, errStrWc, desFormat);
                    break;
                }
                case CELLTYPE_ERROR:
                {
                    qDebug()<<"ERROR: CELLTYPE_ERROR!";
                    break;
                }
            }
        }
        i++;
        QCoreApplication::processEvents();
        mProgress->setValue(mValueOfProgress++);
    }

    //mManuWrongItems
    for(int k=0; k<mManuWrongItems.size(); k++)
    {
        //将错误行号的序号标红
        int num = mManuWrongItems[k];
        CellType rowNumCell = desSheet->cellType(num, 0);
        switch(rowNumCell)
        {
            case CELLTYPE_NUMBER:
            {
                double v = desSheet->readNum(num, 0, &cellFormat);
                QString cellStr = QString::number(v);
                wstring cellStrW = cellStr.toStdWString();
                const wchar_t *cellStrWc = cellStrW.c_str();
                Format *desFormat = bookDes->addFormat(cellFormat);
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, cellStrWc, desFormat);
                break;
            }
            case CELLTYPE_BOOLEAN:
            {
                qDebug()<<"ERROR: BOOL!";
                bool v = desSheet->readBool(num, 0, &cellFormat);
                desSheet->writeBool(num, 0, v, cellFormat);
                break;
            }
            case CELLTYPE_STRING:
            {
                const wchar_t* s = desSheet->readStr(num, 0, &cellFormat);
                Format *desFormat = bookDes->addFormat(cellFormat);
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, s, desFormat);
                break;
            }
            case CELLTYPE_BLANK:
            {
                desSheet->readBlank(num, 0, &cellFormat);
                Format *desFormat = bookDes->addFormat(cellFormat);
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, QString("").toStdWString().c_str(), desFormat);
                break;
            }
            case CELLTYPE_EMPTY:
            {
                Format *desFormat = bookDes->addFormat();
                desFormat->setFillPattern(FILLPATTERN_SOLID);
                desFormat->setPatternForegroundColor(COLOR_RED);
                desSheet->writeStr(num, 0, QString("").toStdWString().c_str(), desFormat);
                break;
            }
            case CELLTYPE_ERROR:
            {
                qDebug()<<"ERROR: CELLTYPE_ERROR!";
                break;
            }
        }


        //若mManuWrongItems为mMinorOfMajorItems中条目
        //只在excel中标记与mMajorOfMajorItems不同的字段
        if(mMinorOfMajorItems.contains(mManuWrongItems[k]))
        {
            //根据厂家、型号信息查找正确的行号
            int sccjIndex = mExcelSheetData[0].indexOf("生产厂家");
            QString sccjStr= mExcelSheetData[mManuWrongItems[k]][sccjIndex].toString();
            int xhIndex = mExcelSheetData[0].indexOf("型号");
            QString xhStr= mExcelSheetData[mManuWrongItems[k]][xhIndex].toString();
            ProducerTypeInfo proAndXh(sccjStr, xhStr);
            int majorKey=mMajorOfMajorItems.key(proAndXh);  //多数条目的行号

            QList<int> col = mMinorOfMajorItems[mManuWrongItems[k]];
            for(int k1=0; k1<col.size(); k1++)
            {

                CellType ct = desSheet->cellType(mManuWrongItems[k], col[k1]);
                switch(ct)
                {
                    case CELLTYPE_NUMBER:
                    {
                        double v = desSheet->readNum(mManuWrongItems[k], col[k1], &cellFormat);
                        QString fieldCorValue = mExcelSheetData[majorKey][col[k1]].toString();
                        QString cellStr = QString::number(v)+"(与对比库不符;"+fieldCorValue+")";
                        wstring cellStrW = cellStr.toStdWString();
                        const wchar_t *cellStrWc = cellStrW.c_str();
                        Format *desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeStr(mManuWrongItems[k], col[k1], cellStrWc, desFormat);
    //                    desSheet->writeNum(num, value[j].num, v, cellFormat);
                        break;
                    }
                    case CELLTYPE_BOOLEAN:
                    {
                        qDebug()<<"ERROR: BOOL!";
                        bool v = desSheet->readBool(mManuWrongItems[k], col[k1], &cellFormat);
                        desSheet->writeBool(mManuWrongItems[k], col[k1], v, cellFormat);
                        break;
                    }
                    case CELLTYPE_STRING:
                    {
                        const wchar_t* s = desSheet->readStr(mManuWrongItems[k], col[k1], &cellFormat);
                        QString fieldCorValue = mExcelSheetData[majorKey][col[k1]].toString();
                        QString errStr = "(与对比库不符;"+fieldCorValue+")";
                        wstring errStrW = errStr.toStdWString();
                        const wchar_t *errStrWc = errStrW.c_str();
                        wchar_t cellStr[300] = {0};
                        wcscpy (cellStr, s);
                        wcscat (cellStr, errStrWc);
                        Format *desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeStr(mManuWrongItems[k], col[k1], cellStr, desFormat);
                        break;
                    }
                    case CELLTYPE_BLANK:
                    {
                        desSheet->readBlank(mManuWrongItems[k], col[k1], &cellFormat);
                        QString fieldCorValue = mExcelSheetData[majorKey][col[k1]].toString();
                        QString errStr = "(与对比库不符;"+fieldCorValue+")";
                        wstring errStrW = errStr.toStdWString();
                        const wchar_t *errStrWc = errStrW.c_str();
                        Format *desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeStr(mManuWrongItems[k], col[k1], errStrWc, desFormat);
    //                    desSheet->writeBlank(num, value[j].num, cellFormat);
                        break;
                    }
                    case CELLTYPE_EMPTY:
                    {
                        QString fieldCorValue = mExcelSheetData[majorKey][col[k1]].toString();
                        QString errStr = "(与对比库不符;"+fieldCorValue+")";
                        wstring errStrW = errStr.toStdWString();
                        const wchar_t *errStrWc = errStrW.c_str();
                        Format *desFormat = bookDes->addFormat();
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeStr(mManuWrongItems[k], col[k1], errStrWc, desFormat);
                        break;
                    }
                    case CELLTYPE_ERROR:
                    {
                        qDebug()<<"ERROR: CELLTYPE_ERROR!";
                        break;
                    }
                }
            }
        }
        //mManuWrongItems不是mMinorOfMajorItems中条目
        //将该记录所有字段都标记
        else
        {
            for(int k1=1; k1<=mColsOfSheet; k1++)
            {
                Format *cellFormat = NULL;
                Format *desFormat = NULL;
 //                cellFormat = desSheet->cellFormat(mManuWrongItems[k], k1-1);


                CellType ct = desSheet->cellType(mManuWrongItems[k], k1-1);
                switch(ct)
                {
                    case CELLTYPE_NUMBER:
                    {
                        double value = desSheet->readNum(mManuWrongItems[k], k1-1, &cellFormat);
                        desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeNum(mManuWrongItems[k], k1-1, value, desFormat);
                        break;
                    }
                    case CELLTYPE_BOOLEAN:
                    {
                        bool value = desSheet->readBool(mManuWrongItems[k], k1-1, &cellFormat);
                        desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeBool(mManuWrongItems[k], k1-1, value, desFormat);
                        break;
                    }
                    case CELLTYPE_STRING:
                    {
                        const wchar_t* s = desSheet->readStr(mManuWrongItems[k], k1-1, &cellFormat);
                        desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeStr(mManuWrongItems[k], k1-1, s, desFormat);
                        break;
                    }
                    case CELLTYPE_BLANK:
                    {
                        desSheet->readBlank(mManuWrongItems[k], k1-1, &cellFormat);
                        desFormat = bookDes->addFormat(cellFormat);
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeBlank(mManuWrongItems[k], k1-1, desFormat);
                        break;
                    }
                    case CELLTYPE_EMPTY:
                    {
                        desFormat = bookDes->addFormat();
                        desFormat->setFillPattern(FILLPATTERN_SOLID);
                        desFormat->setPatternForegroundColor(COLOR_RED);
                        desSheet->writeStr(mManuWrongItems[k], k1-1, L"", desFormat);
                        break;
                    }
                    case CELLTYPE_ERROR:
                    {
                        qDebug()<<"ERROR: CELLTYPE_ERROR!";
                        break;
                    }
                }
            }
        }

        QCoreApplication::processEvents();
        mProgress->setValue(mValueOfProgress++);
    }

    bookDes->save(pathExWc);
    QCoreApplication::processEvents();

}


void EquipManage::exportExcelData()
{
    //若Excel需重新写入
//    QString rangeStr = convertToColName(mColsOfSheet);
//    rangeStr += mRowsOfSheet;
//    rangeStr = "A1:"+rangeStr;
//    qDebug()<<"rangStr: "<<rangeStr;

    //直接将原Excel另存为
//    ExcelHandle *excelHandle1 = new ExcelHandle(mFilePath);
//    QString filePth=QDir::currentPath()+"/test.xlsx";
//    excelHandle1->saveAs(filePth);
//    delete excelHandle1;

    //直接将原Excel另存为
    QXlsx::Document xlsx1(mFilePath);
    QString filePth=QDir::currentPath()+"/test.xlsx";
    xlsx1.saveAs(filePth);

    QCoreApplication::processEvents();

    /******用第三方库Qt xlsx操作excel******/
    QXlsx::Document xlsx(filePth);

    //mWrongItems
    QMap<int, QList<ErrorInfo>>::ConstIterator i=mWrongItems.constBegin();
    while(i != mWrongItems.constEnd())
    {
        int num=i.key();
        QList<ErrorInfo> value=i.value();
        for(int j=0; j<value.size(); j++)
        {
            QXlsx::Cell *cell = xlsx.cellAt(num+1, value[j].num+1);
            QXlsx::Format format;
            if(cell == NULL)
            {
                //QString cellValue = "(空错误;不能为空)";
                QString cellValue="("+value[j].errReason+";"+value[j].recomContent+")";
                format.setPatternBackgroundColor(QColor(255,0,0));
                xlsx.write(num+1, value[j].num+1, cellValue, format);
            }
            else
            {
                format=cell->format();
                format.setPatternBackgroundColor(QColor(255,0,0));
                QVariant v=cell->value();
                QString cellValue=v.toString()+"("+value[j].errReason+";"+value[j].recomContent+")";
                xlsx.write(num+1, value[j].num+1, cellValue, format);
            }
        }
        i++;
        QCoreApplication::processEvents();
        mProgress->setValue(mValueOfProgress++);
    }
    //mManuWrongItems
    for(int k=0; k<mManuWrongItems.size(); k++)
    {
        //若mManuWrongItems为mMinorOfMajorItems中条目
        //只在excel中标记与mMajorOfMajorItems不同的字段
        if(mMinorOfMajorItems.contains(mManuWrongItems[k]))
        {
            //根据厂家、型号信息查找正确的行号
            int sccjIndex = mExcelSheetData[0].indexOf("生产厂家");
            QString sccjStr= mExcelSheetData[mManuWrongItems[k]][sccjIndex].toString();
            int xhIndex = mExcelSheetData[0].indexOf("型号");
            QString xhStr= mExcelSheetData[mManuWrongItems[k]][xhIndex].toString();
            ProducerTypeInfo proAndXh(sccjStr, xhStr);
            int majorKey=mMajorOfMajorItems.key(proAndXh);  //多数条目的行号

            QList<int> col = mMinorOfMajorItems[mManuWrongItems[k]];
            for(int k1=0; k1<col.size(); k1++)
            {
                QXlsx::Cell *cell = xlsx.cellAt(mManuWrongItems[k]+1, col[k1]+1);
                QXlsx::Format format;
                if(cell == NULL)
                {
                    format.setPatternBackgroundColor(QColor(255,0,0));
                    QString fieldCorValue = mExcelSheetData[majorKey][col[k1]].toString();
                    QString cellValue="(与对比库不符;"+fieldCorValue+")";
                    xlsx.write(mManuWrongItems[k]+1, col[k1]+1, cellValue, format);
                }
                else
                {
                    format=cell->format();
                    format.setPatternBackgroundColor(QColor(255,0,0));
                    QVariant v=cell->value();
                    QString fieldCorValue = mExcelSheetData[majorKey][col[k1]].toString();
                    QString cellValue=v.toString()+"(与对比库不符;"+fieldCorValue+")";
                    xlsx.write(mManuWrongItems[k]+1, col[k1]+1, cellValue, format);
                }
            }
        }
        //mManuWrongItems不是mMinorOfMajorItems中条目
        //将该记录所有字段都标记
        else
        {
            for(int k1=1; k1<=mColsOfSheet; k1++)
            {
                QXlsx::Cell *cell = xlsx.cellAt(mManuWrongItems[k]+1, k1);
                QXlsx::Format format;
                if(cell == NULL)
                {
                    format.setPatternBackgroundColor(QColor(255,0,0));
                    xlsx.write(mManuWrongItems[k]+1, k1, "", format);
                }
                else
                {
                    format=cell->format();
                    format.setPatternBackgroundColor(QColor(255,0,0));
                    QVariant v=cell->value();
                    xlsx.write(mManuWrongItems[k]+1, k1, v, format);
                }
            }

        }





        QCoreApplication::processEvents();
        mProgress->setValue(mValueOfProgress++);
    }

    xlsx.save();
    QCoreApplication::processEvents();





/*
    //标记错误信息
    ExcelHandle *excelHandle = new ExcelHandle(filePth);
    excelHandle->setWorkSheet(SHEET_NUM);
    //mWrongItems
    QMap<int, QList<ErrorInfo>>::ConstIterator i=mWrongItems.constBegin();
    while(i != mWrongItems.constEnd())
    {
        int num=i.key();
        QList<ErrorInfo> value=i.value();
        for(int j=0; j<value.size(); j++)
        {
            excelHandle->setCellBackground(SHEET_NUM, num+1, value[j].num+1);
        }
        i++;
    }
    //mManuWrongItems
    for(int k=0; k<mManuWrongItems.size(); k++)
    {
        for(int k1=1; k1<=mColsOfSheet; k1++)
        {
            excelHandle->setCellBackground(SHEET_NUM, mManuWrongItems[k]+1, k1);
        }
    }
    excelHandle->save();
    delete excelHandle;
*/







}

QString EquipManage::convertToColName(int col)
{
    QString res;
    if(col <= 0) return "";
    int quotients=col/26;
    int remainder=col%26;
    if(quotients >=1 && remainder == 0)     //针对是26倍数数字的特殊处理
    {
        quotients--;
        remainder = 26;
    }
    res += convertToColName(quotients);
    QChar ch = remainder + 0x40;
    res += ch;
    return res;
}


void EquipManage::on_aboutAction_triggered()
{

}



void EquipManage::on_cmpDbClickedAction_triggered()
{
    mCmpDbDialog = new cmpareDbDialog(this);
    mCmpDbDialog->setModal(false);
    mCmpDbDialog->show();
}
