#include "mainwindow.h"
#include "findbid.h"
#include <QApplication>
//#include <QTimer>
//#include <QDebug>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
//    MainWindow w;
//    w.show();

    FindBid fb;
    fb.show();

//    QTimer::singleShot(100, &a, SLOT(quit()));
    return a.exec();
}
