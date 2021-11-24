#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    animation = ui->animation;
    btnPath1 = ui->btnFile1;
    btnPath2 = ui->btnFile2;
    btnOutDir = ui->btnOutDir;
    btnStart = ui->btnStart;
    btnExit = ui->btnExit;
    btnGo = ui->btnReset;

    path1 = ui->pathFile1;
    path2 = ui->pathFile2;
    path3 = ui->pathFile3;

    range_start1 = ui->range_start1;
    range_end1 = ui->range_end1;
    range_start2 = ui->range_start2;
    range_end2 = ui->range_end2;
    progress = ui-> progressBar;

    exitFileName = ui->filenameEdit;

    listterminal = ui->listterminal;
    listterminal->addItem("Старт программы;");

    comboSheet1 = ui->comboSheet1;
    comboSheet2 = ui->comboSheet2;
    comboExitCell1 =ui->comboExitCell1;
    comboExitCell2 = ui->comboExitCell2;
    cellAddress1 = ui->cellAddress1;
    cellAddress2 =ui->cellAddress2;

    checkSaveFile = ui->checkSaveFile;

    menubar = ui->menubar;
    aboutAction = ui->action_2;
    connect(aboutAction, SIGNAL(triggered()), this, SLOT(s_about()));
    connect(comboSheet1, SIGNAL(currentTextChanged(QString)), this, SLOT(s_selectCombo(QString)));
    connect(comboSheet2, SIGNAL(currentTextChanged(QString)), this, SLOT(s_selectCombo(QString)));

    bar = ui->statusbar;
    bar->showMessage(tr("Заполните все поля..."));

    loadSettings();
    loadsCell();

    movie = new QMovie(":/animation/loads.gif");
    animation->setMovie(movie);

    btnGo ->setEnabled(false);

}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::print_Terminal(QString s){
    listterminal ->addItem(s);
}

void MainWindow::loadsCell(){
    foreach(QString s, alphabet){
        cellAddress1->addItem(s);
        cellAddress2->addItem(s);
        comboExitCell1->addItem(s);
        comboExitCell2->addItem(s);
    }
}
//++++++++++SLOTS+++++++++++++//
void MainWindow::s_selectCombo(QString s){
    if(sender()->objectName() == "comboSheet1"){
        sheet1 = s;
        qDebug()<< "signal 1";
    }else if(sender()->objectName() == "comboSheet2"){
        sheet2 = s;
                   qDebug()<< "signal 2";
    }
}

void MainWindow::s_getPath1(){
    QString filename;
    filename = QFileDialog::getOpenFileName(this,"Выбрать файл","/home/rza",tr("Xlsx файлы:(*.xlsx)"));
    path1->setText(filename);
}

void MainWindow::s_getPath2(){
    QString filename;
    filename = QFileDialog::getOpenFileName(this,"Выбрать файл","/home/rza",tr("Xlsx файлы:(*.xlsx)"));
    path2->setText(filename);
}

void MainWindow::s_getPath3(){
    QString dir;
    dir = QFileDialog::getExistingDirectory(this,"Выберите директорию","/home/rza");
    path3->setText(dir);
}
//++//
void MainWindow::s_begin(){

    int exitRange1 = comboExitCell1->currentIndex();
    int exitRange2 = comboExitCell2->currentIndex();
    QString filename = exitFileName->text();
    filename.append(".xlsx");
    if((ui->checkSaveFile->isChecked() && (filename=="")) || (exitRange1==0) || (exitRange2==0)){
        qDebug() << "Ошибка!";
        QMessageBox::critical(0, "Ошибка входных данных", "Вы не заполнили какие то из полей, проверьте еще раз!");

    }else{
        qDebug() << "Идет дальше";
    btnStart->setEnabled(false);
    movie->start();
    QVariant tmp = 0;
    quint16 count = 0;
    print_Terminal("Открываем файл" + path1->text());
    QXlsx::Document *document1 = new QXlsx::Document(path1->text());
    document1->selectSheet(sheet1);
    quint16 range1 = static_cast<quint16>(range_start1->text().toInt());
    quint16 end1 = static_cast<quint16>(range_end1->text().toInt());
    progress->setMinimum(range1);
    progress->setMaximum(end1);
    progress->setValue(range1);
    bar->showMessage(QString("Считываем данные"));
    for(quint16 i = range1; i <= end1; i++){
        tmp = document1->read(i, cellAddress1->currentIndex());
        qDebug() << i;
        qDebug() << tmp;
        vec.push_back(tmp);
        tmp = 0;
        count++;
        progress->setValue(count);
        QMovie movie("");
    }
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
    tmp = 0;
    count = 0;
    bar->showMessage(QString("Открываем файл" + ui->pathFile2->text()));
    QXlsx::Document *document2 = new QXlsx::Document(ui->pathFile2->text());
    document2->selectSheet(sheet2);
    range1 = static_cast<quint16>(range_start2->text().toInt());
    end1 = static_cast<quint16>(range_end2->text().toInt());
    progress->setMinimum(range1);
    progress->setMaximum(end1);
    progress->setValue(range1);
    bar->showMessage(QString("Считываем данные"));
    for(quint16 i = range1; i <= end1; i++){
        tmp = document2->read(i, cellAddress2->currentIndex());
        qDebug() <<"Документ2: "<< i << ": " << tmp;
        vec2.push_back(tmp);
        tmp = 0;
        count++;
        progress->setValue(count);
    }
    bar->showMessage(QString("Поиск"));
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
    QXlsx::Document exitDoc;

    for(int i= 0; i < vec.size(); i++){
        for(int j = 0; j < vec2.size(); j++){
            if(vec.at(i) == vec2.at(j)){
                find.push_back(j);
                print_Terminal("Найдено совпадение: ");
                print_Terminal(QVariant(vec2.at(j)).toString());
                qDebug() << "Найдено совпадение: " << vec2.at(j);
                if(!checkSaveFile->isChecked()){

                }else{
                saveInFile(document1, exitDoc, i+1, exitRange1, exitRange2);
                }
            }
        }
    }
    if(!checkSaveFile->isChecked()){
        print_Terminal("Файл не был записан.");
    }else{
    exitDoc.saveAs(filename);
    print_Terminal("Файл был записан.");
    }
    //qDebug() << "Поиск закончен, найдено: " << find.size() << " совпадений.";
    print_Terminal(QString("Поиск закончен, найдено: ").append(QString::number(find.size())).append(" совпадений."));
    bar->showMessage("Готово!");

    //===================ЗАПИСЬ===================//

    vec.clear();
    vec2.clear();
    find.clear();
    btnStart->setEnabled(true);
    movie->stop();
    }
}

void MainWindow::s_exit(){
    saveSettings();
}

void MainWindow::s_about(){
    qDebug() << "About!";
    QMessageBox::information(0, "О программе...", "Программа для поиска сопадений в двух Excel файлах по определенному диапазону.\n Разработана на Qt5 под лицензией GPL.\n");
}

void MainWindow::s_start(){
    if((path1->text()=="") || (path2->text()=="") || (path3->text()=="")){
        QMessageBox::critical(0, "Ошибка входных данных", "Проверьте все поля на заполнение...");
    }else{
   print_Terminal(tr("Анализ входных файлов.."));
   QXlsx::Document *document1 = new QXlsx::Document(path1->text());
   QStringList sheets = document1->sheetNames();
   comboSheet1->addItems(sheets);

   QXlsx::Document *document2 = new QXlsx::Document(path2->text());
   sheets.clear();
   sheets = document2->sheetNames();
   comboSheet2->addItems(sheets);
   print_Terminal(tr("Анализ завершен.."));
   btnGo->setEnabled(true);
   }
}
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
void MainWindow::saveSettings(){
    settings.setValue("/path1/", ui->pathFile1->text());
    settings.setValue("/path2/", ui->pathFile2->text());
    settings.setValue("/path3/", ui->pathFile3->text());
    settings.setValue("filename/", ui->filenameEdit->text());
}
void MainWindow::loadSettings(){
    ui->pathFile1->setText((settings.value("/path1", "/home/rza")).toString());
    ui->pathFile2->setText((settings.value("/path2", "/home/rza")).toString());
    ui->pathFile3->setText((settings.value("/path3", "/home/rza")).toString());
    ui->filenameEdit->setText((settings.value("/filename", "outfile.xlsx")).toString());
}
void MainWindow::saveInFile(QXlsx::Document *doc1, QXlsx::Document &doc2, int index, int exitRange1, int exitRange2){
    for(int i=exitRange1; i<=exitRange2; i++){
        doc2.write(index, i, doc1->read(index, i));

    }
}



