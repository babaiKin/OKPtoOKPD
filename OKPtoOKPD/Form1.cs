using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.IO;

namespace OKPtoOKPD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label1.Visible = false;
            progressBar1.Visible = false;
        }

        Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
        string fileName;
        //string upd;
        string eXt;
        string newCode;
        string name;
        int codePosition=1;
        bool flag = false;
        string error;
        private void ExportToExcel()
        {
            //workSheet.SaveAs(fileName);

            //exApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Visible = true;
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            try
            {
                //OpenFileDialog openFileDialog1 = new OpenFileDialog();
                //openFileDialog1.Filter = "Доступные форматы (*.xls ; *.xlsx)|*.xls; *.xlsx";
                //openFileDialog1.FilterIndex = 2;
                //openFileDialog1.RestoreDirectory = true;
                //openFileDialog1.Title = "Select File";

                //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                //{
                //    eXt = Path.GetExtension(openFileDialog1.SafeFileName);
                //    fileName = openFileDialog1.FileName;
                //}

                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Excel (*.XLS)|*.XLS";
                opf.ShowDialog();
                DataTable tb = new DataTable();
                string filename = opf.FileName;
                MessageBox.Show(filename);
                //if (filename == "")
                //   break;
                string ConStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=Excel 8.0;", filename);
                System.Data.DataSet ds = new System.Data.DataSet("EXCEL");
                OleDbConnection cn = new OleDbConnection(ConStr);
                cn.Open();
                DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                string select = String.Format("SELECT * FROM [{0}]", sheet1);
                OleDbDataAdapter ad = new OleDbDataAdapter(select, cn);
                ad.Fill(ds);
                tb = ds.Tables[0];
                cn.Close();
                //dataGridView1.DataSource = tb;
                //MessageBox.Show(tb.Rows.Count+ "");
                int lc;//последняя ячейка в первом столбце
                for (lc = 0; lc < tb.Rows.Count; lc++)
                {
                    if (tb.Rows[lc][0].ToString() == "")
                        break;
                }

                string[] TNVD = new string[lc];
                for (int i=0; i < lc; i++)
                {
                    newCode = "";
                    name = "";
                    TNVD[i] = tb.Rows[i][0].ToString();
                    
                    //сокращение до 6 символов
                    //только для тех, что не найдены в 10зн
                    //int startIndex = 0;
                    //int length = 6;
                    //String tnvd6 = TNVD[i].Substring(startIndex, length);
                    //TNVD[i] = tnvd6;

                    for (int jj = 2; jj <57447; jj++) //для всех строк nой колонки
                    {
                        string OKPD = tb.Rows[jj][5].ToString(); //текст из jjой ячейки 6 колонки в переменную OKPD
                                                                 //MessageBox.Show(tb.Rows[jj + 1][8].ToString());
                                                                 //string okpd6 = tb.Rows[jj][6].ToString();

                        //OKP[jj] = tb.Rows[jj][5].ToString();
                        //string okpd4 = tb.Rows[jj][4].ToString();
                        String tnvd6 = TNVD[i].Substring(0, 6);
                        String tnvd4 = TNVD[i].Substring(0, 4);
                        //String okpd6 = OKPD.Substring(0, 6);
                        String okpd4 = OKPD.Substring(0, 4);

                        //MessageBox.Show(TNVD[i] + " || " + OKPD);
                        error = TNVD[i] + " || " + OKPD;
                        error = error + "\n" + tnvd6 + " || " + OKPD;
                        error = error + "\n" + tnvd4 + " || " + okpd4;

                        if (TNVD[i] == OKPD) //если OKP=OKPD, то в jую ячейку второй колонки пишем значение из jjой ячейки n+2 колонки
                        {
                            newCode = newCode + "\n" + tb.Rows[jj][8].ToString();
                            name = name + "\n" + tb.Rows[jj][9].ToString();
                        }

                        else
                        {
                            
                            if (tnvd6 == OKPD) //если OKP=OKPD, то в jую ячейку второй колонки пишем значение из jjой ячейки n+2 колонки
                            {
                                newCode = newCode + "\n" + tb.Rows[jj][8].ToString();
                                name = name + "\n" + tb.Rows[jj][9].ToString();
                            }

                            else
                            {
                                if (tnvd4 == okpd4) //если OKP=OKPD, то в jую ячейку второй колонки пишем значение из jjой ячейки n+2 колонки
                                {
                                    newCode = newCode + "\n" + tb.Rows[jj][8].ToString();
                                    name = name + "\n" + tb.Rows[jj][9].ToString();
                                }
                            }
                            
                            /*
                            if (tnvd6 == okpd6) //если OKP=OKPD, то в jую ячейку второй колонки пишем значение из jjой ячейки n+2 колонки
                            {
                                newCode = newCode + "\n" + tb.Rows[jj][8].ToString();
                                name = name + "\n" + tb.Rows[jj][9].ToString();
                                //newCode = tb.Rows[jj][8].ToString();
                                //MessageBox.Show(newCode);
                                //codePosition = jj;
                                //flag = true;
                            }*/
                        }

                        ////if (flag)
                        ////    for (codePosition = jj; codePosition < tb.Rows.Count; codePosition++)
                        ////    //типо диапазон ключей ОКПД
                        ////    //начало определяет отлично, но идет по строчкам до талого
                        ////    //не есть гуд, ибо все строки сюда впихивает............................
                        ////    {
                        ////        if (tb.Rows[codePosition][5].ToString() == "" && tb.Rows[codePosition][8].ToString() != ""/* && flag*/)
                        ////        {
                        ////            //MessageBox.Show(codePosition + " | " + tb.Rows[codePosition + 1][8].ToString());
                        ////            newCode = newCode + "\n" + tb.Rows[codePosition][8].ToString();
                        ////            if (tb.Rows[codePosition + 1][7].ToString() != "")
                        ////                break;
                        ////            //ObjWorkSheet.Cells[j, 2].Value = "\n" + ObjWorkSheet.Cells[jj, 8].Text.ToString();
                        ////        }
                        ////    }
                        codePosition = 0;
                        //flag = false;
                        //MessageBox.Show(newCode);
                        tb.Rows[i][1] = newCode;
                        tb.Rows[i][2] = name;
                        
                    }
                }
                //newCode = "";
                //name = "";
                dataGridView1.DataSource = tb;


                tb.WriteXml(filename);
                
                //экспорт
                //Excel.Application exApp = new Excel.Application();
                //Excel.Workbook workbook = exApp.Workbooks.OpenXML(filename, Type.Missing, LoadOption.PreserveChanges);
                //exApp.Quit();

                //ExportToExcel();


























                ////OpenFileDialog openFileDialog1 = new OpenFileDialog();
                ////openFileDialog1.Filter = "Доступные форматы (*.xls ; *.xlsx)|*.xls; *.xlsx";
                ////openFileDialog1.FilterIndex = 2;
                ////openFileDialog1.RestoreDirectory = true;
                ////openFileDialog1.Title = "Select File";

                ////if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                ////{
                ////    eXt = Path.GetExtension(openFileDialog1.SafeFileName);
                ////    fileName = openFileDialog1.FileName;
                ////}
                //ObjWorkExcel.Visible = true;



                ////Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
                ////                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                ////                       Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                ////                       Type.Missing, Type.Missing, Type.Missing); //открыть файл
                ////Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                ////var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ////string[] str = new string[lastCell]; // массив значений с листа равен по размеру листу


                ///////поиск ОКП в первой колонке
                ///////поиск по ОКП в nой колонке
                ///////сопоставление первой колонки n+2 колонки 

                ////int lc; //ластселл в первой колонке //lastCell - ластселл всего (6-9 колонки)

                ////for (lc=1; lc < lastCell; lc++)
                ////{
                ////    string OKP = ObjWorkSheet.Cells[lc, 1].Text.ToString(); //текст из jой ячейки первой колонки
                ////    if (OKP == "")
                ////        break;

                ////}

                //////MessageBox.Show(lc+"");
                ////for (int j = 2; j < lc; j++) // по всем строкам в первой колонке
                ////{
                ////    string OKP = ObjWorkSheet.Cells[j, 1].Text.ToString(); //текст из jой ячейки первой колонки в переменную OKP
                ////    for (int jj = 2; jj < lastCell; jj++) //для всех строк nой колонки
                ////    {
                ////        string OKPD = ObjWorkSheet.Cells[jj, 6].Text.ToString(); //текст из jjой ячейки 6 колонки в переменную OKPD
                ////                                                                 //MessageBox.Show(ObjWorkSheet.Cells[jj, 6].Text.ToString());
                ////        if (OKP == OKPD) //если OKP=OKPD, то в jую ячейку второй колонки пишем значение из jjой ячейки n+2 колонки
                ////        {
                ////            newCode = ObjWorkSheet.Cells[jj, 8].Text.ToString();
                ////            //codePosition = jj;
                ////            flag = true;
                ////        }

                ////        if (flag)
                ////            for (codePosition = jj; codePosition < lastCell; codePosition++)
                ////            //типо диапазон ключей ОКПД
                ////            //начало определяет отлично, но идет по строчкам до талого
                ////            //не есть гуд, ибо все строки сюда впихивает............................
                ////            {
                ////                if (ObjWorkSheet.Cells[codePosition, 6].Text.ToString() == "" && ObjWorkSheet.Cells[codePosition, 8].Text.ToString() != ""/* && flag*/)
                ////                {
                ////                    //MessageBox.Show(codePosition + " | " + ObjWorkSheet.Cells[codePosition, 8].Text.ToString());
                ////                    newCode = newCode + "\n" + ObjWorkSheet.Cells[codePosition, 8].Text.ToString();
                ////                    if (ObjWorkSheet.Cells[codePosition+1, 6].Text.ToString() != "")
                ////                        break;
                ////                    //ObjWorkSheet.Cells[j, 2].Value = "\n" + ObjWorkSheet.Cells[jj, 8].Text.ToString();
                ////                }

                ////            }
                ////        codePosition =0;
                ////        flag = false;
                ////        ObjWorkSheet.Cells[j, 2].Value = newCode;


                ////    }
                ////    newCode = "";

                ////    progressBar1.Maximum = lc;
                ////    progressBar1.Value = j;
                ////    label1.Text = (j + " строка из " + lc);
                ////}

                ////label1.Visible = false;
                ////progressBar1.Visible = false;
                ////ObjWorkExcel.DisplayAlerts = false;
                ////ObjWorkSheet.SaveAs(fileName);
                ////ObjWorkExcel.Interactive = true;
                ////ObjWorkExcel.ScreenUpdating = true;
                ////ObjWorkExcel.UserControl = true;

                MessageBox.Show("HAPPY END");
                ////ExcelWorkBook.Close(true, null, null);
                ////ExcelApp.Quit();
            }

            catch (Exception err)
            {
                MessageBox.Show(error + "\nFATAL ERROR: " + err);
                //workFlag = false;
            }
        }
    }
}
