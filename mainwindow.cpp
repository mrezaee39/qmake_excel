#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFile>
#include <QDebug>
#include <QTextStream>
#include <xlsxdocument.h>
#include <xlsxworksheet.h>
#include <xlsxformat.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>

using namespace QXlsx;


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    csv_input.setFileName("://Input.csv");
    csv_output.setFileName(":/OutPut.csv");
    excel_file_input = new Document("/home/mohammad/input.xlsx");
    excel_file_output = new Document("/home/mohammad/output.xlsx");
    for (auto x: excel_file_input->sheetNames())
        excel_file_input->deleteSheet(x);
    for (auto x: excel_file_output->sheetNames())
        excel_file_output->deleteSheet(x);
    excel_file_input->save();
    excel_file_output->save();


    Exceledit("A","RG","/home/mohammad/input_RG.xlsx","://Input.csv",1,2,4,5,22,23,90,105,110,117,154,163,163,163,163,163,163,163,163,163,163,163,163,163);
    Exceledit("B","LKA","/home/mohammad/input_LKA.xlsx",":/Input.csv",1,3,24,25,102,107,120,153,164,173,0,173,173,173,173,173,173,173,173,173,173,173,173,173);
    Exceledit("C","MCIG","/home/mohammad/input_MCIG.xlsx",":/Input.csv",4,9,12,15,26,71,74,89,89,89,89,89,89,89,89,89,89,89,89,89,89,89,89,89);
    Exceledit("D","SM","/home/mohammad/input_SM.xlsx",":/Input.csv",1,2,4,11,16,25,28,29,38,39,68,69,72,74,104,109,118,125,76,78,80,84,86,86);
    Exceledit("RG","RG","/home/mohammad/output_RG.xlsx",":/OutPut.csv",29,31,40,43,49,50,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000,1000);
    Exceledit("LKA","LKA","/home/mohammad/output_LKA.xlsx",":/OutPut.csv",3,10,14,15,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43);
    Exceledit("MCIG","MCIG","/home/mohammad/output_MCIG.xlsx",":/OutPut.csv",1,4,14,16,45,47,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43,43);
    Exceledit("SM","SM","/home/mohammad/output_SM.xlsx",":/OutPut.csv",3,4,51,64,11,11,67,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11,11);
    Exceledit("FOC","FOC","/home/mohammad/output_FOC.xlsx",":/OutPut.csv",63,64,63,64,63,64,63,64,63,64,63,64,63,64,63,64,63,64,63,64,63,64,63,64);
    Exceledit("Platform","Platform","/home/mohammad/output_Platform.xlsx",":/OutPut.csv",12,13,17,28,32,39,44,45,65,66,68,103,9,9,9,9,9,9,9,9,9,9,9,48);

    excel_file_input->save();

    excel_file_output->save();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::Exceledit (QString s5, QString s6, QString s7, QString s8, int a, int b, int c, int d, int e, int f, int g, int h , int i, int j, int k, int l, int n, int o
                            ,int w, int x, int y, int z, int aa,int bb,int cc, int dd, int ee, int ff)
{   
    /*s5 : excel_file.addSheet(s5); for us
     *s6 : page name output       ; for them
     *s7 : excel_file output address
     *s8 : csv   file input Path path (which is added in source)
     */
    Q_UNUSED(s7)
    Q_UNUSED(s5)
    QTextStream stream;
    bool is_input;
    Format italic;
    italic.setFontItalic(true);
    Format plain;

    if (s8.contains("input",Qt::CaseInsensitive))
    {
        if (!csv_input.exists())
            exit(1);
        if (!csv_input.isOpen())
            csv_input.open(QFile::ReadOnly);
        stream.setDevice(&csv_input);
        excel_file_input->workbook()->setHtmlToRichStringEnabled(true);
        excel_file_input->addSheet(s6);
        //        excel_file_output->selectSheet(s6);
        is_input = true;
    }
    else
    {
        if (!csv_output.exists())
            exit(1);
        if (!csv_output.isOpen())
            csv_output.open(QFile::ReadOnly);
        stream.setDevice(&csv_output);
        //        excel_file_output->workbook()->setHtmlToRichStringEnabled(true);
        excel_file_output->addSheet(s6);
        //        excel_file_output->selectSheet(s6);
        //        excel_file_output->currentSheet()->setSheetState(AbstractSheet::SS_Visible);
        is_input = false;
    }



    int line_number = 1;
    int signal_number = 0 ;


    while(!stream.atEnd()){

        QString line;
        QStringList linelist;
        line = stream.readLine();
        linelist=line.split(",");
        signal_number++;
        if( ((signal_number>=a)&&(signal_number<=b))||((signal_number>=c)&&(signal_number<=d))||((signal_number>=g)&&(signal_number<=h))||((signal_number>=e)&&(signal_number<=f))||((signal_number>=i)&&(signal_number<=j))
                ||((signal_number>=k)&&(signal_number<=l))||((signal_number>=n)&&(signal_number<=o))||((signal_number>=w)&&(signal_number<=x))||((signal_number>=y)&&(signal_number<=z))||signal_number==aa||signal_number==bb||
                signal_number==cc||signal_number==dd||signal_number==ee||signal_number==ff)

        {


            qDebug()<<"method"<<line_number;
            RichString signal;
            QString str;
            str.setNum(signal_number);
            signal.addFragment( str,plain);
            //excel_file.write(line_number,1,str);

            //line_number++;
            RichString cell_format0;
            cell_format0.addFragment(s6, italic);
            cell_format0.addFragment(" component shall receive the output signal ",plain); // or inputs
            cell_format0.addFragment(linelist[0].remove("\""), italic);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format0);
            else
                excel_file_output->write(line_number,1,cell_format0);
            line_number++;

            RichString cell_format1;
            cell_format1.addFragment("the signal ",plain);
            cell_format1.addFragment(linelist[0].remove("\""), italic);
            cell_format1.addFragment(" shall have the Data Type ",plain);
            cell_format1.addFragment(linelist[4].remove("\""),plain);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format1);
            else
                excel_file_output->write(line_number,1,cell_format1);
            line_number++;
            RichString cell_format2;
            cell_format2.addFragment("the signal ",plain);
            cell_format2.addFragment(linelist[0].remove("\""), italic);
            cell_format2.addFragment(" shall have the unit ",plain);
            cell_format2.addFragment(linelist[1].remove("\""),plain);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format2);
            else
                excel_file_output->write(line_number,1,cell_format2);
            line_number++;

            RichString cell_format3;
            cell_format3.addFragment("the signal ",plain);
            cell_format3.addFragment(linelist[0].remove("\""), italic);
            cell_format3.addFragment(" shall have the resolution 0.001",plain);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format3);
            else
                excel_file_output->write(line_number,1,cell_format3);
            line_number++;

            RichString cell_format4;
            cell_format4.addFragment("the signal ",plain);
            cell_format4.addFragment(linelist[0].remove("\""), italic);
            cell_format4.addFragment(" shall have the max value ",plain);
            cell_format4.addFragment(linelist[3].remove("\""),plain);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format4);
            else
                excel_file_output->write(line_number,1,cell_format4);
            line_number++;
            RichString cell_format5;
            cell_format5.addFragment("the signal ",plain);
            cell_format5.addFragment(linelist[0].remove("\""), italic);
            cell_format5.addFragment(" shall have the min value ",plain);
            cell_format5.addFragment(linelist[2].remove("\""),plain);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format5);
            else
                excel_file_output->write(line_number,1,cell_format5);
            line_number++;
            RichString cell_format6;
            cell_format6.addFragment("the signal ",plain);
            cell_format6.addFragment(linelist[0].remove("\""), italic);
            cell_format6.addFragment(" shall have the default value XXX ",plain);
            if (is_input)
                excel_file_input->write(line_number,1,cell_format6);
            else
                excel_file_output->write(line_number,1,cell_format6);
            line_number++;
            line_number++;
            if (is_input)
            {
                excel_file_input->moveSheet(s6,1);
                excel_file_input->write(line_number,1,cell_format6);
                excel_file_input->save();
            }
            else
            {
                excel_file_output->moveSheet(s6,1);
                excel_file_output->write(line_number,1,cell_format6);
                excel_file_output->save();
            }

        }
    }

}


