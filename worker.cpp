#include "worker.h"

Worker::Worker(QObject *parent) :
    QObject(parent)
{
    _mapBase=0;
    _mapBaseHead=0;
    _vectorIn=0;
    _mapInHead=0;
    _stopped=false;
}

void Worker::initialize(QMultiMap<QString, QStringList> &mapBase,
                        QMap<QString, int> &mapBaseHead,
                        Vector &vectIn,
                        QMap<QString, int> &mapInHead)
{
    qDebug() << "Worker init";
    _mapBase=&mapBase;
    _mapBaseHead=&mapBaseHead;
    _vectorIn=&vectIn;
    _mapInHead=&mapInHead;
}

void Worker::process()
{
    qDebug() << "Worker process" << this->thread()->currentThreadId();
    qDebug() << "Поиск в словаре БД";
    QList<QStringList> *listRightRows = new QList<QStringList>;
    for(int i=0; i<_vectorIn->size(); i++)
    {
        if(_stopped)
            break;
        bool founded=false;
        QStringList row=_vectorIn->at(i);
        QString ename=row.at(_mapInHead->value("E"));
        qDebug() << ename << row.at(0);
        if(!ename.isEmpty())
        {
            QList<QStringList> *listBaseChoosenEname=new QList<QStringList>(_mapBase->values(ename));
            for(int i=0; i<listBaseChoosenEname->size(); i++)
            {
                if(_stopped)
                    break;
//                TODO
//                QString street=listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET"));
//                if(street consist [\\s-.]+)
//                split :
//                QStringList listWordsOfStreet=.split(QRegExp("[]"));
//                for(list)
//                compare all word (find any part in map)
                if(levDistance(listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET")).toStdString(),
                               row.at(_mapInHead->value("STR")).toStdString()) <= 2)
                {
                    if(listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD"))
                            == row.at(_mapInHead->value("B")))
                    {

                        if(listBaseChoosenEname->at(i).at(_mapBaseHead->value("KORP"))
                                == row.at(_mapInHead->value("K")))
                        {

//                            qDebug() << "STR"
//                                     << listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET"))
//                                     << row.at(_mapInHead->value("STR"))
//                                     << endl
//                                     << "B"
//                                     << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD"))
//                                     << row.at(_mapInHead->value("B"))
//                                     << endl
//                                     << "K"
//                                     << listBaseChoosenEname->at(i).at(_mapBaseHead->value("KORP"))
//                                     << row.at(_mapInHead->value("K"));

//                            nDebugString++;
//                            if(nDebugString>10)
//                                assert(0);
                            QStringList findedRow;
                            findedRow << row.at(0)
                                      << row.at(_mapInHead->value("E"))
                                      << row.at(_mapInHead->value("STR"))
                                      << row.at(_mapInHead->value("B"))
                                      << row.at(_mapInHead->value("K"))
                                      << row.at(_mapInHead->value("ADD"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("ENAME"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("KORP"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("ADD"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD_ID"));
                            founded=true;
                            qDebug() << "f:" << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD_ID"));
                            emit addRow(findedRow);
                        }
                    }
                }
            }//end for2
            delete listBaseChoosenEname;
            if(!founded)
            {
                QStringList findedRow;
                findedRow << row.at(0)
                          << row.at(_mapInHead->value("E"))
                          << row.at(_mapInHead->value("STR"))
                          << row.at(_mapInHead->value("B"))
                          << row.at(_mapInHead->value("K"))
                          << row.at(_mapInHead->value("ADD"))
                          << "not found";
                qDebug() << "not f:" << row.at(_mapInHead->value("STR"));
                emit addRow(findedRow);
            }
        }
        else
        {
            //if enameIn empty
            QList<QStringList> *listBaseChoosenEname=new QList<QStringList>(_mapBase->values());
            for(int i=0; i<listBaseChoosenEname->size(); i++)
            {
                if(_stopped)
                    break;
                if(levDistance(listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET")).toStdString(),
                               row.at(_mapInHead->value("STR")).toStdString()) <= 2)
                {
                    if(listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD"))
                            == row.at(_mapInHead->value("B")))
                    {

                        if(listBaseChoosenEname->at(i).at(_mapBaseHead->value("KORP"))
                                == row.at(_mapInHead->value("K")))
                        {

//                            qDebug() << "STR"
//                                     << listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET"))
//                                     << row.at(_mapInHead->value("STR"))
//                                     << endl
//                                     << "B"
//                                     << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD"))
//                                     << row.at(_mapInHead->value("B"))
//                                     << endl
//                                     << "K"
//                                     << listBaseChoosenEname->at(i).at(_mapBaseHead->value("KORP"))
//                                     << row.at(_mapInHead->value("K"));

//                            nDebugString++;
//                            if(nDebugString>10)
//                                assert(0);
                            QStringList findedRow;
                            findedRow << row.at(0)
                                      << row.at(_mapInHead->value("E"))
                                      << row.at(_mapInHead->value("STR"))
                                      << row.at(_mapInHead->value("B"))
                                      << row.at(_mapInHead->value("K"))
                                      << row.at(_mapInHead->value("ADD"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("ENAME"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("STREET"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("KORP"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("ADD"))
                                      << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD_ID"));
                            founded=true;
                            qDebug() << "f:" << listBaseChoosenEname->at(i).at(_mapBaseHead->value("BLD_ID"));
                            emit addRow(findedRow);
                        }
                    }
                }
            }//end for2
            delete listBaseChoosenEname;
            if(!founded)
            {
                QStringList findedRow;
                findedRow << row.at(0)
                          << row.at(_mapInHead->value("E"))
                          << row.at(_mapInHead->value("STR"))
                          << row.at(_mapInHead->value("B"))
                          << row.at(_mapInHead->value("K"))
                          << row.at(_mapInHead->value("ADD"))
                          << "not found";
                qDebug() << "not f:" << row.at(_mapInHead->value("STR"));
                emit addRow(findedRow);
            }
        }//end else
    }//end for 1
    delete listRightRows;
    emit finished();
}


void Worker::stop()
{
    _stopped=true;
}

size_t Worker::levDistance(const std::string &s1, const std::string &s2)
{
  const size_t m(s1.size());
  const size_t n(s2.size());

  if( m==0 ) return n;
  if( n==0 ) return m;

  size_t *costs = new size_t[n + 1];

  for( size_t k=0; k<=n; k++ ) costs[k] = k;

  size_t i = 0;
  for ( std::string::const_iterator it1 = s1.begin(); it1 != s1.end(); ++it1, ++i )
  {
    costs[0] = i+1;
    size_t corner = i;

    size_t j = 0;
    for ( std::string::const_iterator it2 = s2.begin(); it2 != s2.end(); ++it2, ++j )
    {
      size_t upper = costs[j+1];
      if( *it1 == *it2 )
      {
          costs[j+1] = corner;
      }
      else
      {
        size_t t(upper<corner?upper:corner);
        costs[j+1] = (costs[j]<t?costs[j]:t)+1;
      }

      corner = upper;
    }
  }

  size_t result = costs[n];
  delete [] costs;

  return result;
}
