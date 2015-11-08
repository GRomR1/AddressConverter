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

#include <QAxBase>
#include <QAxObject>

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

private:
    Ui::MainWindow *ui;
    QVector<QStringList> _vector;
    QList<QVariant> _listData;
    QString _filenameSave;
    QString _stringResultForCsv;

    void appendText(QString text);
    bool openFromCsv(QString filename, QTableWidget *tbl, int countRow);
    bool saveToCsv(QString fname, QTableWidget *tbl);
    bool saveToXls();
    void clearResultData();
};

#endif // MAINWINDOW_H
