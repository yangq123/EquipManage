#ifndef CMPAREDBDIALOG_H
#define CMPAREDBDIALOG_H

#include <QDialog>
#include <QSqlQueryModel>

namespace Ui {
class cmpareDbDialog;
}

class cmpareDbDialog : public QDialog
{
    Q_OBJECT

public:
    explicit cmpareDbDialog(QWidget *parent = nullptr);
    ~cmpareDbDialog();

private:
    Ui::cmpareDbDialog *ui;
    void showCompareDb();
};

#endif // CMPAREDBDIALOG_H
