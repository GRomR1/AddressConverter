#ifndef FINDBID_H
#define FINDBID_H

#include <QWidget>
#include <QMap>
#include <QMultiMap>
#include <QVector>
#include <QDebug>
#include <QString>
#include <QStringList>
#include <QFileDialog>
#include <QTextCodec>
#include <QMessageBox>
#include <QTableWidget>
#include <QTime>
#include <QDateTime>
#include <assert.h>

#include "worker.h"

namespace Ui {
class FindBid;
}

typedef QVector<QStringList> Vector;

class FindBid : public QWidget
{
    Q_OBJECT

public:
    explicit FindBid(QWidget *parent = 0);
    ~FindBid();

signals:

public slots:
    void appendRow(QStringList rowList);

private slots:
    void on_pushButtonOpenBase_clicked();
    void on_pushButtonOpenIn_clicked();
    void on_pushButtonClear_clicked();
    void on_pushButtonFind_clicked();
    void on_pushButtonStop_clicked();
    void on_pushButtonSave_clicked();

private:
    Ui::FindBid *ui;
    Vector _vectorBase;
    QMultiMap<QString, QStringList> _mapBase;
    QStringList _headBase;
    QMap<QString, int> _mapBaseHead;
    Vector _vectorIn;
    QStringList _headIn;
    QMap<QString, int> _mapInHead;
    Vector _result;
    Worker *_worker;
    QThread _thread;

    void openIn(QString str);
    void openBase(QString str);
    bool openFromCsv(QString filename, Vector *data, QStringList &head, int countRow);
    bool saveToCsv(QString fname);
};

#endif // FINDBID_H
