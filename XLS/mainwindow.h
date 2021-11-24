#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QPushButton>
#include <QLineEdit>
#include <QtXlsx>
#include <QMessageBox>
#include <QFileDialog>
#include <QDebug>
#include <QStatusBar>
#include <QProgressBar>
#include <QVector>
#include <QSettings>
#include <QMenuBar>
#include <QAction>
#include <QListWidget>
#include <QComboBox>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();


    QPushButton *btnPath1;
    QPushButton *btnPath2;
    QPushButton *btnOutDir;
    QPushButton *btnStart;
    QPushButton *btnExit;
    QPushButton *btnGo;
    QLineEdit *path1;
    QLineEdit *path2;
    QLineEdit *path3;
    QMessageBox *msg;
    QLineEdit *range_start1;
    QLineEdit *range_end1;
    QLineEdit *range_start2;
    QLineEdit *range_end2;
    QLineEdit *exitFileName;
    QStatusBar *bar;
    QProgressBar *progress;
    QVector<QVariant> vec;
    QVector<QVariant> vec2;
    QVector<int> find;
    QSettings settings;
    QMenuBar *menubar;
    QAction *aboutAction;
    QLabel *animation;
    QMovie *movie;
    QListWidget *listterminal;
    QComboBox *comboSheet1;
    QComboBox *comboSheet2;
    QComboBox *cellAddress1;
    QComboBox *cellAddress2;
    QComboBox *comboExitCell1;
    QComboBox *comboExitCell2;
    QCheckBox *checkSaveFile;


private:
    Ui::MainWindow *ui;
    void saveSettings();
    void loadSettings();
    void print_Terminal(QString);
    void saveInFile();
    void loadsCell();
    void saveInFile(QXlsx::Document*doc1, QXlsx::Document &doc2, int index,int,int);
    const QString alphabet = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    QString sheet1, sheet2;
private slots:
    void s_getPath1();
    void s_getPath2();
    void s_getPath3();
    void s_start();
    void s_exit();
    void s_about();
    void s_begin();
    void s_selectCombo(QString);

};
#endif // MAINWINDOW_H
