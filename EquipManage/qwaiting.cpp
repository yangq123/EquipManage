#include "qwaiting.h"
#include "ui_qwaiting.h"

QWaiting::QWaiting(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::QWaiting)
{
    ui->setupUi(this);

    //设置透明度
    //this->setWindowOpacity(0.8);

    //取消对话框标题
    //setWindowFlags(Qt::Dialog | Qt::FramelessWindowHint);//设置为对话框风格，并且去掉边框
    //setWindowModality(Qt::WindowModal);//设置为模式对话框，同时在构造该对话框时要设置父窗口
    //ui->label->setStyleSheet("background-color: transparent;");
    mMovie = new QMovie(":/images/res/loading.gif");
    mMovie->setCacheMode(QMovie::CacheAll);
    ui->label->setMovie(mMovie);

    //ui->label->setScaledContents(true);
    mMovie->start();
}

QWaiting::~QWaiting()
{
    mMovie->stop();
    delete ui;
}
