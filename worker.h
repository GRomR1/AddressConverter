#ifndef WORKER_H
#define WORKER_H

#include <QObject>
#include <QMutex>
#include <QDebug>
#include <QString>
#include <QThread>
#include <QTime>
#include <QDateTime>
#include <QStringList>
#include <QList>

#include <string>

typedef QVector<QStringList> Vector;

class Worker : public QObject
{
    Q_OBJECT
public:
    explicit Worker(QObject *parent = 0);
    void initialize(
            QMultiMap<QString, QStringList> &mapBase,
            QMap<QString, int> &mapBaseHead,
            Vector &vectIn,
            QMap<QString, int> &mapInHead);
    ~Worker()
    {
        qDebug() << "Worker destructor";
    }

signals:
    void addRow(QStringList row);
    void finished();

public slots:
    void process();
    void stop();

private:
    QMultiMap<QString, QStringList> *_mapBase;
    QMap<QString, int> *_mapBaseHead;
    Vector *_vectorIn;
    QMap<QString, int> *_mapInHead;
    bool _stopped;

    size_t levDistance(const std::string &s1, const std::string &s2);
};

#endif // WORKER_H
