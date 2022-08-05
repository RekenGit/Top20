using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using TopAwarii;

namespace LandTech_Top20
{
    public partial class Form1 : Form
    {
        string Path ="";
        string Path2 ="";
        int rows = 1;
        int column = 1;

        public Form1()
        {
            InitializeComponent();
            label1.Parent = pictureBox1;
            label1.BackColor = Color.Transparent;
            label2.Parent = pictureBox1;
            label2.BackColor = Color.Transparent;
            label3.Parent = pictureBox1;
            label3.BackColor = Color.Transparent;
        }

        private void button1_Click(object sender, EventArgs e) //Wprowadz dane
        {
            if (Path == "" && Path2 == "")
            {
                MessageBox.Show("Podaj Plik tygodniowy i Plik wynikowy");
                return;
            }
            else if (Path == "" && Path2 != "")
            {
                MessageBox.Show("Podaj Plik tygodniowy");
                return;
            }
            else if (Path != "" && Path2 == "")
            {
                MessageBox.Show("Podaj Plik wynikowy");
                return;
            }

            progressBar2.Maximum = 8000;

            policzExcel(Path); //liczy ilosc wierszy / zazwyczaj 20
            Excel excel = new Excel(Path, 1);
            string[,] ID = excel.ReadRange(1, 1, rows, 1);
            string[,] Ile = excel.ReadRange(1, 2, rows, 2);
            excel.Close();

            int rowsFirst = rows;

            policzExcel(Path2); //liczy z koncowego
            Excel ex = new Excel(Path2, 1);
            string[,] sklepID = ex.ReadRange(2, 2, rows, 2);
            ex.Close();

            progressBar2.Maximum = 10;
            progressBar2.Value = 10;
            progressBar1.Maximum = rowsFirst;
            progressBar1.Value = 40;//procesbar

            int DodajRows = rows;

            button1.Enabled = false;

            Excel exl = new Excel(Path2, 1); //otwiera by edytowac w trakcie
            int tydzien = column - 5;
            exl.WriteCell(1, column + 1, "T "+tydzien);
            progressBar1.Increment(1);//procesbar
            int sprawdzone = 0;
            if (rows - 1 == 0)
            {
                for (int i = 0; i < rowsFirst; i++)
                {
                    progressBar1.Increment(1);//procesbar
                    DodajRows++;
                    int r = DodajRows - 1;
                    exl.WriteCell(DodajRows, 1, r.ToString());
                    exl.WriteCell(DodajRows, 2, ID[i, 0]);
                    exl.WriteCell(DodajRows, column + 1, Ile[i, 0]);    //nie istnieje nicc
                }
            }
            for (int i = 0; i < rowsFirst; i++)
            {
                progressBar1.Increment(1);//procesbar
                for (int x = 0; x < rows-1; x++)
                {
                    if (sklepID[x, 0] == ID[i, 0])
                    {
                        exl.WriteCell(x + 2, column + 1, Ile[i, 0]);    // istnieje id
                        sprawdzone = 0;
                        break;
                    }
                    else
                    {
                        sprawdzone++;
                        if (sprawdzone == rows - 1)
                        {
                            DodajRows++;
                            int r = DodajRows - 1;
                            exl.WriteCell(DodajRows, 1, r.ToString());
                            exl.WriteCell(DodajRows, 2, ID[i, 0]);
                            exl.WriteCell(DodajRows, column + 1, Ile[i, 0]);    //nie istnieje id

                            sprawdzone = 0;
                        }
                    }
                }
            }
            exl.Save();
            exl.Close();
            progressBar1.Maximum = 100;
            progressBar1.Value = 98;
            Thread.Sleep(2000);
            progressBar1.Value = 100;
            button1.Text = "GOTOWE!";
            button1.BackColor = Color.Gray;
        }

        private void panel1_DragEnter_1(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void panel1_DragDrop_1(object sender, DragEventArgs e)
        {
            string[] Files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string file in Files)
            {
                Path = file.ToString();
                FileInfo fi = new FileInfo(Path);
                textBox1.Text = fi.Name + " • " + Path;
                panel1.BackColor = Color.LightGreen;
            }
        }

        private void panel2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void panel2_DragDrop(object sender, DragEventArgs e)
        {
            string[] Files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string file in Files)
            {
                Path2 = file.ToString();
                FileInfo fi = new FileInfo(Path2);
                textBox2.Text = fi.Name + " • " + Path2;
                panel2.BackColor = Color.LightGreen;
            }
        }

        void policzExcel(string x) {
            Excel excel = new Excel(x, 1);
            rows = 1;
            column = 1;
            for (;;)
            {
                progressBar2.Increment(1);
                if (excel.ReadCell(rows, 1) != "")
                {
                    rows++;
                }
                else
                {
                    break;
                }
            }
            for (;;)
            {
                progressBar2.Increment(1);
                if (excel.ReadCell(1, column) != "")
                {
                    column++;
                }
                else
                {
                    break;
                }
            }
            rows--;
            column--;
            excel.Close();
        }
    }
}
