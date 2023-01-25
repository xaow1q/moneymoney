using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using static System.Windows.Forms.DataFormats;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Collections;
//
namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        int maxd, maxr, minr, mind;
        string dcateg;

        


        public Form1()
        {
            InitializeComponent();
        }

        public void refreshdatagrid(DataGridView dg)
        {
            ArrayList col1Items = new ArrayList();
            ArrayList col2Items = new ArrayList();
            ArrayList col3Items = new ArrayList();
            ArrayList col4Items = new ArrayList();
            ArrayList col5Items = new ArrayList();

            int i = 0; int a = 0;
            foreach (DataGridViewRow dr in dg.Rows)
            {
                col1Items.Add(dr.Cells[0].Value);
                col2Items.Add(dr.Cells[2].Value);
                col3Items.Add(dr.Cells[1].Value);

            }
            for (int iIndex = 0; iIndex < col1Items.Count; iIndex++)
            {
                object o = col1Items[iIndex];
                i += Convert.ToInt32(o);
                object b = col2Items[iIndex];
                a += Convert.ToInt32(b);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
            {
                button5.Enabled = false;
            }

            refreshdatagrid(dgv);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgv.DataSource != null)
                {
                    DataTable dataTable = (DataTable)dgv.DataSource;
                    DataRow drToAdd = dataTable.NewRow();

                    drToAdd["Доходы"] = "Value2";
                    drToAdd["Расходы"] = "Value2";
                    drToAdd["Категория доходов"] = "Value2";
                    drToAdd["Расходы"] = "Value2";
                    drToAdd["Категория расходов"] = "Value2";
                    drToAdd["Комментарий"] = "Value2";
                    drToAdd["Дата"] = "Value2";
                    dataTable.Rows.Add(drToAdd);
                    dataTable.AcceptChanges();
                }
                else
                {
                    MessageBox.Show("Добавьте таблицу");
                }

            }
            catch
            {
                if (dgv.DataSource == null)
                    MessageBox.Show("Добавьте таблицу");
                if (dgv.DataSource != null)
                    MessageBox.Show("Ошибка");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string fileName = "Бюджет.xlsx";
                string path = Path.Combine(Environment.CurrentDirectory, fileName);
                string file = "";   
                DataTable dt = new DataTable();   
                DataRow row;
                DialogResult result = openFileDialog1.ShowDialog();  
                if (result == DialogResult.OK)   
                {
                    file = openFileDialog1.FileName; 
                    try
                    {
                        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);
                        Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                        Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

                        int rowCount = excelRange.Rows.Count;  
                        int colCount = excelRange.Columns.Count;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                            }
                            break;
                        }
                                   
                        int rowCounter;  

                        for (int i = 2; i <= rowCount; i++) 
                        {
                            row = dt.NewRow();  
                            rowCounter = 0;
                            for (int j = 1; j <= colCount; j++) 
                            {                               
                                if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                                {
                                    row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                                }
                                else
                                {
                                    row[i] = "";
                                }

                                rowCounter++;
                            }
                            dt.Rows.Add(row); 
                        }
                        dgv.DataSource = dt; 

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(excelRange);
                        Marshal.ReleaseComObject(excelWorksheet);
                        excelWorkbook.Close();
                        Marshal.ReleaseComObject(excelWorkbook);
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void button4_Click(object sender, EventArgs e) //Удаление строки
        {
            if (dgv.DataSource != null)
            {
                if (dgv.Rows[0].Cells[1].Value == null)
                {
                    MessageBox.Show("Больше нечего удалять");
                }
                else
                {
                    foreach (DataGridViewRow row in dgv.SelectedRows)
                    {
                        dgv.Rows.Remove(row);
                    }
                }

            }
            else
            {
                MessageBox.Show("Нечего удалять");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
            {
                MessageBox.Show("Нечего редактировать");
            }
            else
            {
                dgv.CurrentCell.Value = textBox1.Text;
                dgv.EndEdit();
                dgv.CurrentCell.Value = textBox1.Text;
                dgv.EndEdit();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Worksheet ws = ExcelApp.Workbooks.Add().ActiveSheet;
            ws.Columns.ColumnWidth = 17;
            ws.get_Range("A1", "z100").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.get_Range("A1", "z100").Cells.Font.Name = "Comic Sans MS";
            Excel.Range chartRange;
            Excel.Range HeaderRange;
            chartRange = ExcelApp.get_Range("a3", "u12");

            foreach (Excel.Range cell in chartRange.Cells)
            {
                chartRange.Cells.BorderAround(Type.Missing, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);
                chartRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            HeaderRange = ExcelApp.get_Range("a3", "a12");
            foreach (Excel.Range cell in HeaderRange.Cells)
            {
                HeaderRange.Cells.Font.Bold = true;
            }

            ExcelApp.Cells[1, 1] = "Доходы";
            ExcelApp.Cells[2, 1] = "Категории доходов";
            ExcelApp.Cells[3, 1] = "Расходы";
            ExcelApp.Cells[4, 1] = "Категории расходов";
            ExcelApp.Cells[5, 1] = "Комментарий";
            ExcelApp.Cells[6, 1] = "Дата";

            for (int i = 0; i < dgv.ColumnCount - 1; i++)
            {
                for (int j = 0; j < dgv.RowCount; j++)
                {
                    ExcelApp.Cells[i, j] = (dgv[i, j].Value).ToString();
                    if (ExcelApp.Cells[1, j] == null)
                    {
                        ExcelApp.Cells[i, j] = null;
                    }
                }
            }
            ExcelApp.Visible = true;
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgv.RowCount; i++)
            {
                dgv.Rows[i].Selected = false;
                for (int j = 0; j < dgv.ColumnCount; j++)
                    if (dgv.Rows[i].Cells[j].Value != null)
                        if (dgv.Rows[i].Cells[j].Value.ToString().Contains(textBox2.Text))
                        {
                            dgv.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int min = 0;
            //Convert.ToInt32(dgv.Rows[0].Cells[1].Value); //это если искать минимум в столбце с индексом 1
            for (int i = 0; i < dgv.ColumnCount - 1; i++) //если AllowUserToAddRows=false, иначе i < dataGridView1.RowCount-1
            {
                if (dgv.Rows[0].Cells[i].Value != null && (int)dgv.Rows[0].Cells[i].Value < (int)dgv.Rows[0].Cells[0].Value)
                    min = (int)dgv.Rows[i].Cells[1].Value; 
            }
              MessageBox.Show(min.ToString());
        }

        private void button11_Click(object sender, EventArgs e)
        {
            refreshdatagrid(dgv);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form prompt = new Form();
            prompt.Width = 500;
            prompt.Height = 500;
            prompt.Text = "";
            FlowLayoutPanel panel = new FlowLayoutPanel();
            System.Windows.Forms.Button ok = new System.Windows.Forms.Button() { Text = "Yes" };
            ok.Click += (sender, e) => { prompt.Close(); };
            System.Windows.Forms.Button no = new System.Windows.Forms.Button() { Text = "No" };
            no.Click += (sender, e) => { prompt.Close(); };
            System.Windows.Forms.TextBox txt = new System.Windows.Forms.TextBox();
            txt.Multiline = true;
            txt.Width = 350;
            txt.Height = 420;
            txt.Text = $"Общий доход = {maxd}\n Общие расходы = {maxr}";
            prompt.Controls.Add(txt);
            prompt.ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
        }
    }
}

            



