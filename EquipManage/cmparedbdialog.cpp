#include "cmparedbdialog.h"
#include "ui_cmparedbdialog.h"
#include <QMessageBox>
#include <QDebug>
#include <QSqlError>

cmpareDbDialog::cmpareDbDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::cmpareDbDialog)
{
    ui->setupUi(this);
    showCompareDb();
}

cmpareDbDialog::~cmpareDbDialog()
{
    delete ui;
}

void cmpareDbDialog::showCompareDb()
{
    //操作数据库
    {
        QString dbName = "equipManage.db";
        QString dbPath = QString("%1/%2").arg(QCoreApplication::applicationDirPath()).arg(dbName);

        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "equipManage.db");
        db.setDatabaseName(dbPath);
        if(!db.open())
        {
            QMessageBox::warning(this, tr("警告"), tr("数据库打开失败!"));
            qDebug()<<db.lastError();
            return;
        }

        QSqlQueryModel *queryModelComDb = new QSqlQueryModel(ui->tableView_cmpDb);
        queryModelComDb->setQuery("select * from compareLib_DuanLuQi", db);
        ui->tableView_cmpDb->setModel(queryModelComDb);
        ui->tableView_cmpDb->resizeColumnsToContents();
        ui->tableView_cmpDb->horizontalHeader()->setStretchLastSection(true);
        db.close();
    }
    QSqlDatabase::removeDatabase("equipManage.db");
}

