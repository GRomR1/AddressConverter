#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QString>
#include <QStringList>
#include <QDebug>
#include <QFileDialog>
#include <QTextCodec>
#include <QMessageBox>

#include <QAxBase>
#include <QAxObject>

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

private:
    Ui::MainWindow *ui;
    QVector<QStringList> _vector;
    QList<QVariant> _listData;
    QString _filenameSave;
    QString _stringResultForCsv;

    void appendText(QString text);
    bool saveToCsv();
    bool saveToXls();
    void clearResultData();
};

#endif // MAINWINDOW_H
