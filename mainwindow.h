#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    void Exceledit(QString s5, QString s6, QString s7, QString s8, int a, int b, int c, int d, int e, int f, int g, int h , int i, int j, int k, int l, int n, int o
                   ,int w, int x, int y, int z, int aa,int bb,int cc, int dd, int ee, int ff);

private:
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
