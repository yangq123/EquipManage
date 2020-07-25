#ifndef QWAITING_H
#define QWAITING_H

#include <QDialog>
#include <QMovie>

namespace Ui {
class QWaiting;
}

class QWaiting : public QDialog
{
    Q_OBJECT

public:
    explicit QWaiting(QWidget *parent = nullptr);
    ~QWaiting();

private:
    QMovie *mMovie;
    Ui::QWaiting *ui;
};

#endif // QWAITING_H
