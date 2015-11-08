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
   ui->radioButtonCsv->hide();
   ui->radioButtonXls->hide();
//   ui->pushButtonClear->hide();
   ui->pushButtonConvert->hide();
   ui->tableWidgetIn->hide();
   ui->lineEditOpen->hide();
   ui->pushButtonOpen->hide();
   ui->tableWidget->hide();
   ui->pushButtonSave->setEnabled(true);
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

bool MainWindow::openFromCsv(QString filename, QTableWidget *tbl, int countRow=-1)
{
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
    QStringList head = rowList->at(0).split(";");//получение шапки
    if(head.size()<6)
        head.append("ENAME");
    if(head.size()<7)
        head.append("ADD");
    int cols=head.size();
    for(int i=0; i<cols; i++)
    {
        head[i].remove("\"");
        head[i] = head.at(i).trimmed();
    }
    QVector<QStringList> vect;
    int rows = 0;
    for(int i=1; i<rowList->size(); i++)
    {
//        qDebug() << rowList.at(i);
        QStringList row = rowList->at(i).split(";");
        workWitkRow(row); // **** DO IT Rishat ***
        if(!row.isEmpty())
        {
            rows++;
            vect.append(row);
            //оставливаем обработку если получено нужное количество строк
            if(countRow > 0 && rows >= countRow)
            {
                rows = countRow;
                break;
            }
        }
    }
    delete rowList;
//    qDebug() << "Всего строк подходящих по условию:" << vect.size();
    qDebug() << "Столбцов:" << cols;

    //заполняем таблицу
    QTableWidgetItem *ptwi = 0;
    tbl->setColumnCount(cols);
    tbl->setHorizontalHeaderLabels(head);    //получение шапки

    //заполнение таблицы TableWidget
    if(countRow > 0)
        if(rows > countRow)
            rows = countRow;
    qDebug() << "Всего строк открыто (будет отображено):" << rows;
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
    qDebug() << stringResultForCsv;
    for(int row=0; row<rows; row++)
    {
        for(int col=0; col<cols; col++)
        {
            QTableWidgetItem *ptwi = tbl->item(row, col);
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


void MainWindow::on_pushButtonSave_clicked()
{
    //TO DO Only for debug
    ui->radioButtonCsv->setChecked(true);

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
        if(saveToCsv(_filenameSave, ui->tableWidgetBase))
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
//    ui->plainTextEditFrom->clear();
//    ui->pushButtonConvert->setEnabled(false);
//    ui->pushButtonSave->setEnabled(false);
//    ui->pushButtonClear->setEnabled(false);
//    ui->lineEditOpen->setEnabled(false);
//    ui->lineEditOpen->setText(trUtf8("Нажмите открыть файл"));
    ui->lineEditOpenBase->setText(trUtf8("Нажмите открыть файл"));
//    ui->lineEditOpen->clear();
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
    while(row.size()<7)
        row.append("");
    if(!row.at(1).contains("г. Санкт-Петербург", Qt::CaseInsensitive))
    {
        row.clear();
        return;
    }
    QString ename;
    QString additional;
    for(int i=0; i<row.size(); i++)
    {

        //работаем с STR
        if(i==1)
        {
            //удаляем фразу "г. Санкт-Петербург"
            if (row.at(i).contains("г. Санкт-Петербург,"))
            {
                row[i].remove("г. Санкт-Петербург,");
            }
            //переносм п. и г. в столбец ADD
            if (row.at(i).contains("п."))
            {
             row[i].remove("п. ");
             QStringList field = row.at(i).split(",");
             row[i].remove(field[0]+",");
             row[6].append(field[0]);

            }
            if (row.at(i).contains("г."))
            {
             row[i].remove("г. ");
             QStringList field = row.at(i).split(",");
             row[i].remove(field[0]+",");
             row[6].append(field[0]);

            }
            //приводим к нижнему регистру
            row[i] = row[i].toLower();
            //работаем с именами элементов (приведение их к одному формату)
            //(ул., пр., наб., ш., б., пер. и пр.)
            if (row.at(i).contains("ул.,"))
            {
              row[i].remove("ул.,");
              row[5].append("ул");
            }
            if (row.at(i).contains("пр-кт.,"))
            {
              row[i].remove("пр-кт.,");
              row[5].append("пр-кт");
            }
            if (row.at(i).contains("пер.,"))
            {
              row[i].remove("пер.,");
              row[5].append("пер");
            }
            if (row.at(i).contains("проезд.,"))
            {
              row[i].remove("проезд.,");
              row[5].append("проезд");
            }
            if (row.at(i).contains("линия.,"))
            {
              row[i].remove("линия.,");
              row[5].append("линия");
            }
            if (row.at(i).contains("наб.,"))
            {
              row[i].remove("наб.,");
              row[5].append("наб");
            }
            if (row.at(i).contains("парк.,"))
            {
              row[i].remove("парк.,");
              row[5].append("парк");
            }
            if (row.at(i).contains("б-р.,"))
            {
              row[i].remove("б-р.,");
              row[5].append("б-р");
            }
            if (row.at(i).contains("ш.,"))
            {
              row[i].remove("ш.,");
              row[5].append("ш");
            }
            if (row.at(i).contains("сад.,"))
            {
              row[i].remove("сад.,");
              row[5].append("сад");
            }
            if (row.at(i).contains("остров.,"))
            {
              row[i].remove("остров.,");
              row[5].append("остров");
            }
            if (row.at(i).contains("пл.,"))
            {
              row[i].remove("пл.,");
              row[5].append("пл");
            }
            if (row.at(i).contains("аллея.,"))
            {
              row[i].remove("аллея.,");
              row[5].append("аллея");
            }
            if (row.at(i).contains("кв-л.,"))
            {
              row[i].remove("кв-л.,");
              row[5].append("кв-л");
            }
            if (row.at(i).contains("снт.,"))
            {
              row[i].remove("снт.,");
              row[5].append("снт");
            }
            if (row.at(i).contains("тер.,"))
            {
              row[i].remove("тер.,");
              row[5].append("тер");
            }


            //работа со скобками
           /* if (row.at(i).contains("(") && row.at(i).contains(")"))
            {
                row[i].remove(QRegExp("( (.+) )"));
            }*/
//            QString str = "п. Парголово, Дальневосточная (Михайловка) ул., ";
//            // QString str1 = "aaa bbb ccc";
//             if (str.contains("(") && str.contains(")")){

//                 str.remove(QRegExp("\*(.+)"));
//             }

//            qDebug() << str;

        }

        //работаем с SID и BID
        if(i==0 || i==2)
        {
            //..
        }

        //работаем с B
        if(i==3)
        {

            //приведение к формату: "%n" - %n - число
            if (row.at(i).contains("д."))
            {
                row[i].remove("д.");
            }
            //удаление записи если это не адрес (напр. "а/я" или "нетр..")
            if (row.at(i).contains("нетр") || row.at(i).contains("а/я")|| row.at(i).contains("ая"))
            {
                row.clear();
                return;
            }
            //выделение корпуса (литеры) из содержимого ячейки
            if (row.at(i).contains("/"))
            {
                QRegExp reg("/(.+)");
                int pos(0);
                while ((pos = reg.indexIn(row.at(i), pos)) != -1) {
                    row[4].append( reg.cap(1) );
                    pos += reg.matchedLength();
                }
                QString copy = row.at(i);
                row[i].clear();
                copy.remove(reg);
                row[3].append(copy);
            }

            // работа с "\"
            if (row.at(i).contains('\\'))
            {
               // row[i].clear();
                //row[i].append("1");
              /*  QRegExp reg("\\(.+)");
                int pos(0);
                while ((pos = reg.indexIn(row.at(i), pos)) != -1) {
                    row[4].append( reg.cap(1) );
                    pos += reg.matchedLength();
                }
                QString copy = row.at(i);
                row[i].clear();
                copy.remove(reg);
                row[3].append(copy);*/
            }

            if (row.at(i).contains(QRegExp ("[а-я]|[А-Я]")))
            {
               QRegExp reg("[а-я]|[А-Я]");
                int pos(0);
                QString copy1 = row.at(4);
                row[4].clear();
                while ((pos = reg.indexIn(row.at(i), pos)) != -1) {
                    row[4].append(copy1+ reg.cap());
                    pos += reg.matchedLength();
                }
                QString copy = row.at(i);

                row[i].clear();
                copy.remove(reg);
                row[3].append(copy);
            }
        }

        //работаем с K
        if(i==4)
        {
            //удаление названия элемента ("корп.", "лит." и пр.)
            if (row.at(i).contains("ЛИТ."))
            {
              row[i].remove("ЛИТ.");
            }
            if (row.at(i).contains("лит"))
            {
              row[i].remove("лит");
            }
            if (row.at(i).contains("лит."))
            {
              row[i].remove("лит.");
            }
            if (row.at(i).contains("ЛИТЕ"))
            {
              row[i].remove("ЛИТЕ");
            }
            if (row.at(i).contains("ЛИТ-"))
            {
              row[i].remove("ЛИТ-");
            }
            if (row.at(i).contains(",ПОМ."))
            {
              row[i].remove(",ПОМ.");
            }
            if (row.at(i).contains("ЛИТ"))
            {
              row[i].remove("ЛИТ");
            }
            if (row.at(i).contains("ер"))
            {
              row[i].remove("ер");
            }
            //приведение к формату: "%n|%c" - %n - число %a - буква(-ы)
            //..
        }

        //работаем с ENAME
        if(i==5)
        {
            //присваивание выделеннного названия структурного элемента из столбца STREET
            //например, ул., пр-т, ш. и пр.
            if(!ename.isEmpty())
                row[i] = ename;
        }

        //работаем с ADD
        if(i==6)
        {
            //присваивание выделенной доп.информации из столбца STREET
            //например, то что содержится в скобках
            if(!additional.isEmpty())
                row[i] = additional;
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
    clearResultData();
    openFromCsv(str, ui->tableWidgetBase, ui->spinBoxRowCount->value());
}
