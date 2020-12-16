using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelListCreation
{
    public partial class Form1 : Form
    {
        string ImportedFilePath;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog vi = new OpenFileDialog();
            vi.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (vi.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = vi.FileName;
                ImportedFilePath = vi.FileName;
                Console.WriteLine(ImportedFilePath); 
                //RichTextBox.Text = Path.GetFileName(vi.FileName);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            readFromExcel(ImportedFilePath);
        }





        private void readFromExcel(string filePath) 
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str=null;
            Object cellValue;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            int nullCell;
            List<RowOfImportedExcel> rows = new List<RowOfImportedExcel>();
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                Console.WriteLine();
                RowOfImportedExcel row = new RowOfImportedExcel();
                nullCell = 0;
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    cellValue = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    if (cellValue is double)
                    {
                        int NumberOfUsers = Convert.ToInt32(cellValue);
                        //Console.Write(NumberOfUsers+" ");
                        row.numberOfUsers = NumberOfUsers;
                        continue;
                    }
                    
                    str = (string)cellValue;

                    //ignore first row if it 's headers
                    if (!String.IsNullOrEmpty(str)){
                        if (str.Equals("StartDate"))
                        {
                            break;
                        }
                    }

                    generateRowValues(str,row,cCnt);
                    //Console.Write(str+" ");

                    if (String.IsNullOrEmpty(str))
                    {
                        nullCell++;
                    }
                    if (nullCell == 5)
                    {
                        break;
                    }
                }

                //not add to list first row with headers
                if (!String.IsNullOrEmpty(str))
                {
                    if (str.Equals("StartDate"))
                    {
                        continue;
                    }
                }
                rows.Add(row);
            }
            printList(rows);

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void printList(List<RowOfImportedExcel> list) 
        {
            foreach (RowOfImportedExcel row in list)
            {
                Console.WriteLine(row.ToString());
            }
        }
        private void generateRowValues(string value,RowOfImportedExcel row,int i)
        { 
            switch(i)
            {
                case 1:
                    DateTime today = DateTime.Today;
                    row.startDate = today.ToString(("dd/MM/yyyy"));
                    break;
                case 2:
                    row.costCenterId = value;
                    break;
                case 3:
                    row.storeName = value;
                    break;
                case 4:
                    row.sapId = value;
                    break;
            }
        }

       
    }
}
