#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    _filenameSave="result.csv";
    ui->plainTextEditTo->hide();
    ui->plainTextEditFrom->hide();

   QAction *ab1 = ui->menuBar->addAction(trUtf8("О программе"));
   connect(ab1, SIGNAL(triggered(bool)),
           this, SLOT(aboutTriggered(bool)));
   ui->radioButtonXls->setToolTip(trUtf8("Медленней"));
   ui->radioButtonCsv->setToolTip(trUtf8("Быстрее"));
   _vector.reserve(89073);
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::appendText(QString text)
{
    QString str = ui->plainTextEditTo->toPlainText();
    if(!str.isEmpty())
        str.append("\n\n***\n");
    ui->plainTextEditTo->setPlainText(str+text);
}

bool MainWindow::openFromCsv(QString filename, QTableWidget *tbl)
{
    QFile file1(str);
    if (!file1.open(QIODevice::ReadOnly))
    {
        qDebug() << "Ошибка открытия для чтения";
        return;
    }
    QTextCodec *defaultTextCodec = QTextCodec::codecForName("Windows-1251");
    QTextDecoder *decoder = new QTextDecoder(defaultTextCodec);
    QString *readedFileStr = new QString (decoder->toUnicode(file1.readAll()));
    delete decoder;
    file1.close();

    QStringList *rowList = new QStringList (readedFileStr->split("\n"));
    delete readedFileStr;
    QString header = rowList->at(0);
    QStringList head = header.split(";");
    int cols=head.size();
    QVector<QStringList> vect;
    for(int i=1; i<rowList->size(); i++)
    {
//        qDebug() << rowList.at(i);
        QStringList row = rowList->at(i).split(";");
        workWitkRow(row); // **** DO IT Rishat ***
        if(!row.isEmpty())
            vect.append(row);
    }
    delete rowList;
    qDebug() << "Всего адресов из СПб в базе: " << vect->size();
    qDebug() << "Столбцов: " << cols;

    //заполняем таблицу
    QTableWidgetItem *ptwi = 0;
    tbl->setColumnCount(cols);
    //получение шапки
    for(int i=0; i<cols; i++)
    {
        head[i].remove("\"");
        head[i] = head.at(i).trimmed();
    }
    tbl->setHorizontalHeaderLabels(head);

    //заполнение таблицы TableWidget
    int rows = vect->size();
    for(int row=0; row<rows; row++)
    {
        tbl->insertRow(tbl->rowCount());
        for(int col=0; col<cols; col++)
        {
            ptwi = new QTableWidgetItem(vect.at(row).at(col));
            tbl->setItem(row, col, ptwi);
        }
    }
    return true;
}

bool MainWindow::saveToCsv(QString fname, QTableWidget *tbl)
{
    QFile file1(fname);
    if (!file1.open(QIODevice::WriteOnly))
    {
        qDebug() << "Ошибка открытия файла для записи";
        return false;
    }
    QTableWidgetItem *ptwi = 0;
    int rows = tbl->rowCount();
    int cols = tbl->columnCount();
    QString stringResultForCsv;
    QStringList strListHeader;
    for(int col=0; col<cols; col++)
    {
        strListHeader.append(tbl->horizontalHeaderItem(col)->text());
    }
    stringResultForCsv.append(strListHeader.join(';'));
    stringResultForCsv+="\n";
    for(int row=0; row<rows; row++)
    {
        for(int col=0; col<cols; col++)
        {
            QTableWidgetItem *ptwi = tbl->item(row, col);
            stringResultForCsv+=ptwi->text()+';';
        }
        stringResultForCsv+="\n";
    }
    QTextCodec *codec = QTextCodec::codecForName("Windows-1251");
    QByteArray ba = codec->fromUnicode(stringResultForCsv.toUtf8());
    file1.write(ba);
    file1.close();
    return true;
}

bool MainWindow::saveToCsv()
{
    QFile file1(_filenameSave);
    if (!file1.open(QIODevice::WriteOnly))
    {
        qDebug() << "Ошибка открытия файла для записи";
        return false;
    }
    QTextCodec *codec = QTextCodec::codecForName("Windows-1251");
    QByteArray ba = codec->fromUnicode(_stringResultForCsv.toUtf8());
    file1.write(ba);
    file1.close();
    return true;
}


void MainWindow::on_pushButtonSave_clicked()
{
    QString filter;
    if(ui->radioButtonCsv->isChecked())
        filter="Excel (*.csv)";
    else if(ui->radioButtonXls->isChecked())
        filter="Excel (*.xls *.xlsx)";
    QString str =
            QFileDialog::getSaveFileName(this,
                                         trUtf8("Укажите куда сохранить результаты"),
                                         "",
                                         filter);

    if(str.isEmpty())
        return;

    _filenameSave = str;
    _filenameSave.replace("/", "\\");

    if(ui->radioButtonCsv->isChecked())
    {
        if(saveToCsv())
            ui->statusBar->showMessage("Успешно сохранено", 5000);
        else
            ui->statusBar->showMessage("Ошибка сохранения", 5000);

    }
    else if(ui->radioButtonXls->isChecked())
    {
        if(saveToXls())
            ui->statusBar->showMessage("Успешно сохранено", 5000);
        else
            ui->statusBar->showMessage("Ошибка сохранения", 5000);
    }
}

void MainWindow::on_pushButtonOpen_clicked()
{
    ui->lineEditOpen->setEnabled(true);
    QString str =
            QFileDialog::getOpenFileName(this, trUtf8("Укажите исходный файл"),
                                         "",
                                         tr("Excel (*.xls *.xlsx)"));
    if(str.isEmpty())
        return;
    ui->lineEditOpen->setText(str);

//    QFile file1(str);
//    if (!file1.open(QIODevice::ReadOnly))
//    {
//        qDebug() << "Ошибка открытия для чтения";
//        return;
//    }
//    QTextCodec *defaultTextCodec = QTextCodec::codecForName("Windows-1251");
//    QTextDecoder *decoder = new QTextDecoder(defaultTextCodec);
//    QString str3 = decoder->toUnicode(file1.readAll());
//    delete decoder;
//    file1.close();

    QAxObject *excel = new QAxObject("Excel.Application",this); // получаем указатель на Excel
    if(excel==NULL)
    {
        qDebug() << "false";
//        return false;
    }
    QAxObject *workbooks = excel->querySubObject("Workbooks");
    if(workbooks==NULL)
    {
        qDebug() << "false";
//        return false;
    }
    // на директорию, откуда грузить книгу
    QAxObject *workbook = workbooks->querySubObject(
                "Open(const QString&)",
                str);

//    QAxObject *workbook = workbooks->querySubObject("Add()");
//    if(workbook==NULL)
//    {
//        qDebug() << "false";
////        return false;
//    }
    QAxObject *sheets = workbook->querySubObject("Sheets");
    if(sheets==NULL)
    {
        qDebug() << "false";
//        return false;
    }

    int count = sheets->dynamicCall("Count()").toInt();
    QString firstSheetName;
    for (int i=1; i<=count; i++)
    {
        QAxObject *sheet1 = sheets->querySubObject("Item(int)", i);
        if(sheet1==NULL)
        {
            qDebug() << "false";
//            return false;
        }
        firstSheetName = sheet1->dynamicCall("Name()").toString();
        sheet1->clear();
        delete sheet1;
        sheet1 = NULL;
        break;
    }

    QAxObject *sheet = NULL;
    if(!firstSheetName.isEmpty())
        sheet = sheets->querySubObject(
                    "Item(const QVariant&)", QVariant(firstSheetName));
    if(sheet==NULL)
    {
        qDebug() << "false";
//        return false;
    }

    connect(excel,SIGNAL(exception(int, QString, QString, QString)),
                     this,SLOT(debugError(int,QString,QString,QString)));
    connect(workbooks, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));
    connect(workbook, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));
    connect(sheets, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));
    connect(sheet, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));

    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    QAxObject *usedRows = usedRange->querySubObject("Rows");
    QAxObject *usedCols = usedRange->querySubObject("Columns");
    int rows = usedRows->property("Count").toInt();
    int cols = usedCols->property("Count").toInt();
    qDebug() << "Количество строк и столбцов в файле: " << rows << cols;

    //заполняем таблицу
    QTableWidget *tbl = ui->tableWidgetIn;
    QTableWidgetItem *ptwi = 0;
    tbl->setColumnCount(cols);
    //получение шапки
    QStringList lst;
    for(int i=1; i<=cols; i++)
    {
        // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
        QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", 1, i);
        // получение содержимого
        QVariant result = cell->property("Value");
        lst.append(result.toString());
        // освобождение памяти
        delete cell;
    }
    tbl->setHorizontalHeaderLabels(lst);

    //заполнение таблицы TableWidget
    for(int row=2; row<=rows; row++)
    {
        tbl->insertRow(tbl->rowCount());
        for(int col=1; col<=cols; col++)
        {
            // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
            QAxObject* cell = sheet->querySubObject("Cells(QVariant,QVariant)", row, col);
            // получение содержимого
            QVariant result = cell->property("Value");
//            qDebug() << row << col << result.toString();
            ptwi = new QTableWidgetItem(result.toString());
            tbl->setItem(row-2, col-1, ptwi);
            // освобождение памяти
            delete cell;
        }
    }

    // освобождение памяти
    usedRange->clear();
    delete usedRange;
    usedRange = NULL;

//    delete usedRows;
//    usedRows = NULL;

//    delete usedCols;
//    usedCols = NULL;

    sheet->clear();
    delete sheet;
    sheet = NULL;

    sheets->clear();
    delete sheets;
    sheets = NULL;

    workbook->clear();
    delete workbook;
    workbook = NULL;

    workbooks->dynamicCall("Close()");
    workbooks->clear();
    delete workbooks;
    workbooks = NULL;

    excel->dynamicCall("Quit()");
    delete excel;
    excel = NULL;

    ui->pushButtonConvert->setEnabled(true);
}

void MainWindow::on_pushButtonClear_clicked()
{
    ui->plainTextEditFrom->clear();
    ui->pushButtonConvert->setEnabled(false);
    ui->pushButtonSave->setEnabled(false);
    ui->pushButtonClear->setEnabled(false);
    ui->lineEditOpen->setEnabled(false);
    ui->lineEditOpen->setText(trUtf8("Нажмите открыть файл"));
    ui->lineEditOpen->clear();
    clearResultData();
}

void MainWindow::clearResultData()
{
    _listData.clear();
    _stringResultForCsv.clear();
    ui->tableWidget->clear();
    ui->tableWidget->setRowCount(0);
    ui->tableWidget->setColumnCount(0);
    ui->tableWidgetBase->clear();
    ui->tableWidgetBase->setRowCount(0);
    ui->tableWidgetBase->setColumnCount(0);
    ui->tableWidgetIn->clear();
    ui->tableWidgetIn->setRowCount(0);
    ui->tableWidgetIn->setColumnCount(0);
}

void MainWindow::debugError(int i, QString s1, QString s2, QString s3)
{
#if 0
    qDebug() << i << endl
             << s1 << endl
             << s2 << endl
             << s3;
#else
    Q_UNUSED(i);
    Q_UNUSED(s1);
    Q_UNUSED(s2);
    Q_UNUSED(s3);
#endif
}

bool MainWindow::saveToXls()
{
//    //WORK!!! Работа со значением в ячейке
//    // получение указателя на ячейку [row][col] ((!)нумерация с единицы)
//    QAxObject *cell = statSheet->querySubObject(
//                "Cells(QVariant,QVariant)", 1, 1);
//    // вставка значения переменной data (любой тип, приводимый к QVariant) в полученную ячейку
//    cell->setProperty("Value", QVariant(100));
//    delete cell;

      // Открытие документа
//    // на директорию, откуда грузить книгу
//    QAxObject *workbook = workbooks->querySubObject(
//                "Open(const QString&)", "C:\\Shared\\Code\\build-TxtToExcel-Desktop_Qt_5_5_0_MinGW_32bit-Debug\\book1.xls" );

      // Получение возможных команды
//    QFile outfile("excel_sheets.html");
//    outfile.open(QFile::WriteOnly | QFile::Truncate);
//    QTextStream out( &outfile );
//    QString docu = sheets->generateDocumentation();
//    out << docu;
//    outfile.close();

//    // Получение количесва листов и списка их имен
//    int count = sheets->dynamicCall("Count()").toInt();
//    QStringList sheetsList;
//    for (int i=1; i<=count; i++)
//    {
//        QAxObject *sheet = sheets->querySubObject("Item(int)", i);
//        QString name = sheet->dynamicCall("Name()").toString();
//        sheetsList.append(name);
//    }
//    qDebug() << sheetsList;

//    //добавляем новый лист и обзываем его sSheetName
//    QAxObject *newSheet = sheets->querySubObject("Add");
//    newSheet->setProperty("Name", "List1");

//    // выбор листа по имени
//    QAxObject *statSheet = sheets->querySubObject(
//                "Item(const QVariant&)", QVariant("List1") );

//    // управление параметрами excel
//    excel->dynamicCall( "SetVisible(bool)", true ); //делаем его видимым
//    excel->dynamicCall( "SetVisible(bool)", false ); //делаем его невидимым
//    excel->dynamicCall("SetDisplayAlerts(bool)", false); //отключение вопросов

    QAxObject *excel = new QAxObject("Excel.Application",this); // получаем указатель на Excel
    if(excel==NULL)
        return false;
    QAxObject *workbooks = excel->querySubObject("Workbooks");
    if(workbooks==NULL)
        return false;
    QAxObject *workbook = workbooks->querySubObject("Add()");
    if(workbook==NULL)
        return false;
    QAxObject *sheets = workbook->querySubObject("Sheets");
    if(sheets==NULL)
        return false;

    int count = sheets->dynamicCall("Count()").toInt();
    QString firstSheetName;
    for (int i=1; i<=count; i++)
    {
        QAxObject *sheet1 = sheets->querySubObject("Item(int)", i);
        if(sheet1==NULL)
            return false;
        firstSheetName = sheet1->dynamicCall("Name()").toString();
        sheet1->clear();
        delete sheet1;
        sheet1 = NULL;
        break;
    }

    QAxObject *sheet = NULL;
    if(!firstSheetName.isEmpty())
        sheet = sheets->querySubObject(
                    "Item(const QVariant&)", QVariant(firstSheetName));
    if(sheet==NULL)
        return false;

    connect(excel,SIGNAL(exception(int, QString, QString, QString)),
                     this,SLOT(debugError(int,QString,QString,QString)));
    connect(workbooks, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));
    connect(workbook, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));
    connect(sheets, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));
    connect(sheet, SIGNAL(exception(int,QString,QString,QString)),
            this, SLOT(debugError(int,QString,QString,QString)));

    int row = 1;
    int col = 1;
    // получение указателя на левую верхнюю ячейку [row][col] ((!)нумерация с единицы)
    QAxObject* cell1 = sheet->querySubObject("Cells(QVariant&,QVariant&)", row, col);
    if(cell1==NULL)
        return false;
    // получение указателя на правую нижнюю ячейку [row + numRows - 1][col + numCols - 1] ((!) numRows>=1,numCols>=1)
    QAxObject* cell2 = sheet->querySubObject("Cells(QVariant&,QVariant&)", row + _listData.size() - 1, col + 3 - 1);
    if(cell2==NULL)
        return false;
    // получение указателя на целевую область
    QAxObject* range = sheet->querySubObject("Range(const QVariant&,const QVariant&)", cell1->asVariant(), cell2->asVariant() );
    if(range==NULL)
        return false;
    // собственно вывод
    range->setProperty("Value", QVariant(_listData) );

    // сохранение
    _filenameSave.remove(".xls");
    workbook->dynamicCall("SaveAs(const QString &)",
                          _filenameSave);

    // освобождение памяти
    range->clear();
    delete range;
    range = NULL;

    cell1->clear();
    delete cell1;
    cell1 = NULL;

    cell2->clear();
    delete cell2;
    cell1 = NULL;

    sheet->clear();
    delete sheet;
    sheet = NULL;

    sheets->clear();
    delete sheets;
    sheets = NULL;

    workbook->clear();
    delete workbook;
    workbook = NULL;

    workbooks->dynamicCall("Close()");
    workbooks->clear();
    delete workbooks;
    workbooks = NULL;

    excel->dynamicCall("Quit()");
    delete excel;
    excel = NULL;

    return true;
}

void MainWindow::on_pushButtonConvert_clicked()
{
    QMessageBox::information(0, "Information", trUtf8("Действие не реализовано"));
    return;

    clearResultData();

    QString str = ui->plainTextEditFrom->toPlainText();
    appendText(str);
    str.remove("\n");
    appendText(str);

    QStringList fios = str.split(";", QString::SkipEmptyParts);
    foreach (QString fio, fios) {
        QStringList l2 = fio.split(QRegExp("((\\]-\\[)|(\\]\\s+-\\s+\\[)|(\\]-\\s+\\[)|(\\]\\s+-\\[))"));
//        QStringList l2 = fio.split("]-[");
        QList<QVariant> celList;
        QStringList::iterator it=l2.begin();
        for(; it!=l2.end(); it++ )
        {
            it->remove(QRegExp("(\\]|\\[)"));
            *it = it->simplified();
            celList.append(*it);
        }
        _listData.append(QVariant(celList));
    }

    //заполняем таблицу
    QTableWidget *tbl = ui->tableWidget;
    QTableWidgetItem *ptwi = 0;
    tbl->setColumnCount(3);
    QStringList lst;
    lst << "Имя" << "Фамилия" << "Отчество";
    tbl->setHorizontalHeaderLabels(lst);
    for(int row=0; row<_listData.size(); row++)
    {
        QList<QVariant> var = _listData.at(row).toList();
        if(var.size()==3)
        {
            tbl->insertRow(tbl->rowCount());
            for(int col=0; col<var.size(); col++)
            {
                QString s = var.at(col).toString();
                ptwi = new QTableWidgetItem(s);
                tbl->setItem(row, col, ptwi);
                _stringResultForCsv+=s+';';
            }
            _stringResultForCsv+="\n";
        }
    }
    appendText(_stringResultForCsv);
    ui->pushButtonClear->setEnabled(true);
    ui->pushButtonSave->setEnabled(true);
}


void MainWindow::aboutTriggered(bool checked)
{
    Q_UNUSED(checked)
    QString text;
    text.append(tr("<p><h2>Конвертер txt в Excel</h2></p>"));
    text.append(tr("<p><h3>По заказу компании <a href='rt.ru'>Ростелеком</a>.</h3></p>"));
    text.append(tr("<p>Программу разработал <a href='http://gromr1.blogspot.ru/p/about-me.html'>Гайнанов Руслан</a>, студент ИТМО.</p>"));
    text.append(tr("<p><br />Связаться с автором: <a href='mailto:ruslan.r.gainanov@gmail.com'>ruslan.r.gainanov@gmail.com</a>.</p>"));
    QMessageBox::about(this, tr("О программе"), text);
}


void MainWindow::workWitkRow(QStringList &row)
{
    if(row.size()!=5)
    {
        row.clear();
        return;
    }
    if(row.at(1).contains("г. Санкт-Петербург", Qt::CaseInsensitive))
    {
        row.clear();
        return;
    }
    for(int i=0; i<row->size(); i++)
    {
        //работаем с STR
        if(i==1)
        {
            //удаляем фразу "г. Санкт-Петербург"
            //..
            //удаляем ","
            //..
            //приводим к нижнему регистру
            //..
            //работаем с именами элементов (приведение их к одному формату)
            //(ул., пр., наб., ш., б., пер. и пр.)
            //..
            //работа со скобками
            //..
        }

        //работаем с SID и BID
        if(i==0 || i==2)
        {
            //..
        }

        //работаем с B
        if(i==3)
        {
            //выделение корпуса (литеры) из содержимого ячейки
            //..
            //приведение к формату: "%n" - %n - число
            //..
            //удаление записи если это не адрес (напр. "а/я" или "нетр..")
            //..
        }

        //работаем с K
        if(i==4)
        {
            //удаление названия элемента ("корп.", "лит." и пр.)
            //..
            //приведение к формату: "%n|%c" - %n - число %a - буква(-ы)
            //..
        }

        //удаляем боковые символы
        row[i].remove("\"");
        row[i] = row.at(i).trimmed();

    }
}

void MainWindow::on_pushButtonOpenBase_clicked()
{
    ui->lineEditOpenBase->setEnabled(true);
    QString str =
            QFileDialog::getOpenFileName(this, trUtf8("Укажите файл базы данных"),
                                         "",
                                         tr("Excel (*.csv)"));
    if(str.isEmpty())
        return;
    ui->lineEditOpenBase->setText(str);

    QFile file1(str);
    if (!file1.open(QIODevice::ReadOnly))
    {
        qDebug() << "Ошибка открытия для чтения";
        return;
    }
    QTextCodec *defaultTextCodec = QTextCodec::codecForName("Windows-1251");
    QTextDecoder *decoder = new QTextDecoder(defaultTextCodec);
    QString *str3 = new QString (decoder->toUnicode(file1.readAll()));
    delete decoder;
    file1.close();

    QVector<QStringList> *vect=&_vector;
    QStringList *rowList = new QStringList (str3->split("\n"));
    QString header = rowList->at(0);
    delete str3;
    for(int i=1; i<rowList->size(); i++)
    {
//        qDebug() << rowList.at(i);
        QStringList row = rowList->at(i).split(";");
        workWitkRow(row); // **** DO IT Rishat ***
        if(!row.isEmpty())
            vect->append(row);
    }
    delete rowList;
    int cols=5;
    qDebug() << "Всего адресов из СПб в базе: " << vect->size();
    qDebug() << "Столбцов: " << cols;

    //заполняем таблицу
    QTableWidget *tbl = ui->tableWidgetBase;
    QTableWidgetItem *ptwi = 0;
    tbl->setColumnCount(cols);
    //получение шапки
    QStringList lst;
    for(int i=0; i<cols; i++)
    {
        QStringList head = header.split(";");
        head[i].remove("\"");
        head[i] = head.at(i).trimmed();
        lst.append(head.at(i));
    }
    tbl->setHorizontalHeaderLabels(lst);

    //заполнение таблицы TableWidget
    for(int row=0; row<500; row++)
    {
        tbl->insertRow(tbl->rowCount());
        for(int col=0; col<cols; col++)
        {
            ptwi = new QTableWidgetItem(vect->at(row).at(col));
            tbl->setItem(row, col, ptwi);
        }
    }
}
