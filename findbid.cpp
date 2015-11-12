#include "findbid.h"
#include "ui_findbid.h"

//static int nDebugString = 0;

FindBid::FindBid(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::FindBid)
{
    ui->setupUi(this);
    _worker=NULL;
    openBase("corrAdd_All.csv");
    openIn("inData1.csv");
}

FindBid::~FindBid()
{
    if(_thread.isRunning())
        _worker->stop();
    _thread.quit();
    _thread.wait();
    delete ui;
}

void FindBid::on_pushButtonOpenBase_clicked()
{
    QString str =
            QFileDialog::getOpenFileName(this, trUtf8("Укажите файл базы данных"),
                                         "",
                                         tr("Excel (*.csv)"));
    if(str.isEmpty())
        return;
    openBase(str);
}

void FindBid::on_pushButtonOpenIn_clicked()
{
    QString str =
            QFileDialog::getOpenFileName(this, trUtf8("Укажите файл заказчика"),
                                         "",
                                         tr("Excel (*.csv)"));
    if(str.isEmpty())
        return;
    openIn(str);
}

void FindBid::on_pushButtonClear_clicked()
{
    _vectorIn.clear();
    _headBase.clear();
    _headIn.clear();
    _mapBase.clear();
    ui->pushButtonFind->setEnabled(false);
    ui->pushButtonSave->setEnabled(false);
}

void FindBid::openIn(QString str)
{
    if(openFromCsv(str, &_vectorIn, _headIn, -1))
    {
        for(int i=0; i<_headIn.size(); i++)
            _mapInHead.insert(_headIn.at(i), i);

        ui->pushButtonOpenBase->setEnabled(true);
    }
}

void FindBid::openBase(QString str)
{
    Vector vectorBase;
    vectorBase.reserve(87477);
    if(openFromCsv(str, &vectorBase, _headBase, -1))
    {
        for(int i=0; i<_headBase.size(); i++)
        {
            _mapBaseHead.insert(_headBase.at(i), i);
        }
        for(int i=0; i<vectorBase.size(); i++)
        {
            QString ename = vectorBase.at(i).at(_mapBaseHead.value("ENAME"));
            _mapBase.insert(ename, vectorBase.at(i));
        }

        if(!_vectorIn.isEmpty())
            ui->pushButtonFind->setEnabled(true);
    }
}

bool FindBid::openFromCsv(QString filename, Vector *vect, QStringList &head, int countRow=-1)
{
    qint64 n1 = QDateTime::currentMSecsSinceEpoch();
    QFile file1(filename);
    if (!file1.open(QIODevice::ReadOnly))
    {
        qDebug() << "Ошибка открытия для чтения";
        return false;
    }
    QTextCodec *defaultTextCodec = QTextCodec::codecForName("Windows-1251");
    QTextDecoder *decoder = new QTextDecoder(defaultTextCodec);
    QString *readedFileStr = new QString (decoder->toUnicode(file1.readAll()));
    delete decoder;
    file1.close();

    QStringList *rowList = new QStringList (readedFileStr->split("\n"));
    delete readedFileStr;
    head = rowList->at(0).split(";");//получение шапки
    vect->clear();
    for(int i=1; i<rowList->size(); i++)
    {
        QStringList row = rowList->at(i).split(";");
        if(!row.isEmpty())
        {
            vect->append(row);
            //оставливаем обработку если получено нужное количество строк
            if(countRow > 0 && vect->size() >= countRow)
                break;
        }
    }
    delete rowList;
    qint64 n2 = QDateTime::currentMSecsSinceEpoch();
    qDebug() << "Времени затрачено на чтение" << n2-n1 << QTime().addMSecs(n2-n1).toString("hh:mm:ss.zzz");
    qDebug() << "Всего строк подходящих по условию:" << vect->size();
    qDebug() << "Столбцов:" << head.size();
    return true;
}


void FindBid::on_pushButtonFind_clicked()
{
    qDebug() << "FindBid find" << this->thread()->currentThreadId();
    if(_thread.isRunning())
        return;
    ui->pushButtonSave->setEnabled(true);
    _worker=new Worker;
    _worker->initialize(
                _mapBase,
                _mapBaseHead,
                _vectorIn,
                _mapInHead);
    _worker->moveToThread(&_thread);
    connect(&_thread, SIGNAL(started()),
            _worker, SLOT(process()));
    connect(_worker, SIGNAL(finished()),
            &_thread, SLOT(quit()));
    connect(_worker, SIGNAL(finished()),
            _worker, SLOT(deleteLater()));
    connect(_worker, SIGNAL(addRow(QStringList)),
            this, SLOT(appendRow(QStringList)));
    _thread.start();

//    qDebug() << "Количество записей в словаре БД";
//    foreach (QString ename, _mapBase.uniqueKeys()) {
//        qDebug() << ename << _mapBase.values(ename).size();
//    }
}


void FindBid::appendRow(QStringList rowList)
{
    QTableWidget *tbl = ui->tableWidgetResult;
    int row = tbl->rowCount();
    tbl->insertRow(row);
    for(int i=0;i<rowList.size(); i++)
        tbl->setItem(row, i, new QTableWidgetItem(rowList.at(i)));
}

void FindBid::on_pushButtonStop_clicked()
{
    if(_worker!=NULL)
    {
        _worker->stop();
    }
}

void FindBid::on_pushButtonSave_clicked()
{
    QString str =
            QFileDialog::getSaveFileName(this,
                                         trUtf8("Укажите куда сохранить результаты"),
                                         "",
                                         "Excel (*.csv)");

    if(str.isEmpty())
        return;
    str.replace("/", "\\");

    if(saveToCsv(str))
        qDebug() << "success saved";
    else
        qDebug() << "error saved!";
}

bool FindBid::saveToCsv(QString fname)
{
    qDebug() << "FindBid saveToCsv" << fname;
    QTableWidget *tbl=ui->tableWidgetResult;
    QFile file1(fname);
    if (!file1.open(QIODevice::WriteOnly))
    {
        qDebug() << "Ошибка открытия файла для записи";
        return false;
    }
    int rows = tbl->rowCount();
    int cols = tbl->columnCount();
    QString stringResultForCsv;
    QStringList strListHeader;
    for(int col=0; col<cols; col++)
        strListHeader.append(tbl->horizontalHeaderItem(col)->text());
    stringResultForCsv.append(strListHeader.join(';'));
    stringResultForCsv+="\n";
//    qDebug() << stringResultForCsv;
    for(int row=0; row<rows; row++)
    {
        for(int col=0; col<cols; col++)
        {
            QTableWidgetItem *ptwi = tbl->item(row, col);
            if(ptwi)
                stringResultForCsv+=ptwi->text();
            if(col<cols-1)
                stringResultForCsv+=';';
        }
        stringResultForCsv+="\n";
    }
    QTextCodec *codec = QTextCodec::codecForName("Windows-1251");
    QByteArray ba = codec->fromUnicode(stringResultForCsv.toUtf8());
    file1.write(ba);
    file1.close();
    return true;
}
