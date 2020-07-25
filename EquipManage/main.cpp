#include "equipmanage.h"

#include <QApplication>


//bool operator ==(const ProducerTypeInfo &pt1, const ProducerTypeInfo &pt2)        //重载==运算符
//{
//    if(pt1.producer==pt2.producer && pt1.type==pt2.type) return true;
//    else
//        return false;
//}

uint qHash(const ProducerTypeInfo &pt)
{
    QString pro=pt.producer;
    QString ty=pt.type;
    QChar *cPro=pro.data();
    QChar *cTy=ty.data();
    int value=0;
    while(!cPro->isNull())
    {
        value+=cPro->unicode();
        ++cPro;
    }
    while(!cTy->isNull())
    {
        value+=cTy->unicode();
        ++cTy;
    }
    return value;
}

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    EquipManage w;
    w.show();
    return a.exec();
}
