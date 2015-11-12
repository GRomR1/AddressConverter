#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QString>
#include <QStringList>
#include <QDebug>
#include <QFileDialog>
#include <QTextCodec>
#include <QMessageBox>
#include <QTableWidget>
#include <QTime>
#include <QDateTime>

#include <QAxBase>
#include <QAxObject>
#include <assert.h>

const int CountRow = 1000; //макс. количество строк базы данных которое возможно открыть

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_pushButtonSave_clicked();
    void on_pushButtonOpen_clicked();
    void on_pushButtonClear_clicked();
    void on_pushButtonConvert_clicked();
    void aboutTriggered(bool checked);
    void debugError(int,QString,QString,QString);

    void on_pushButtonOpenBase_clicked();

    /**
     * @brief
     * Работа со строкой базы данных
     * @param row Список записей строки БД
     */
    void workWitkRow(QStringList &row);

    /**
     * @brief
     * Работа со строкой исходных данных
     * @param row Список записей строки
     */
    void workWitkRowIn(QStringList &row, bool isOneColumnAddr,
                       int colStr, int colB, int colK,
                       int colE, int colAdd);

private:
    Ui::MainWindow *ui;
    QVector<QStringList> _vectorBase;
    QVector<QStringList> _vectorIn;
    QList<QVariant> _listData;
    QString _filenameSave;
    QString _stringResultForCsv;

    void appendText(QString text);
    bool openFromCsv(QString filename, QTableWidget *tbl, int countRow);
    bool saveToCsv(QString fname, QTableWidget *tbl);
    bool saveToCsv(QString fname, QStringList head, QVector<QStringList> &vect);
    bool saveToXls();
    void clearResultData();
};

#endif // MAINWINDOW_H
