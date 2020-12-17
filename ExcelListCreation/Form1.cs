using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace ExcelListCreation
{
    public partial class Form1 : Form
    {
        private string importedFilePath=null;
        private string exportFilePath=null;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog vi = new OpenFileDialog();
            vi.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (vi.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = vi.FileName;
                importedFilePath = vi.FileName;
                Console.WriteLine(importedFilePath); 
                //RichTextBox.Text = Path.GetFileName(vi.FileName);

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fld = new FolderBrowserDialog();
            if (fld.ShowDialog() == DialogResult.OK)
            {
                exportFilePath = (string)fld.SelectedPath;
                textBox2.Text = exportFilePath;
                //MessageBox.Show(exportFilePath);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
             if (!String.IsNullOrEmpty(importedFilePath) && !String.IsNullOrEmpty(exportFilePath)) 
            {
                List<RowOfImportedExcel> rowsOfImportedExcel = ReadFromExcel(importedFilePath);
                List<RowOfExportedExcel> rowsOfExportedExcel = GenerateRowOfExportedExcel(rowsOfImportedExcel);
                ExportToExcel(rowsOfExportedExcel, exportFilePath);
            }
            else 
            { 
                if (String.IsNullOrEmpty(importedFilePath))
                {
                    MessageBox.Show("Please choose an excel file");
                }
                else if (String.IsNullOrEmpty(exportFilePath)) 
                {
                    MessageBox.Show("Please choose a folder");
                }
           
            }
            
        }

        private void ExportToExcel(List<RowOfExportedExcel> rowsOfExportedExcel, string ExcelSavingPath)
        {
            string fileName="FileName.xlsx";
            string path = ExcelSavingPath;
            string fullExcelSavingPath = path+"\\" +fileName;
            GenerateExcel(ConvertToDataTable(rowsOfExportedExcel),fullExcelSavingPath);
        }

        public static void GenerateExcel(DataTable dataTable, string path)
        {

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);
            // create a excel app along side with workbook and worksheet and give a name to it
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Excel.Worksheet excelWorkSheet= excelWorkBook.Sheets.Add();
            Excel.Range workSheet_range = excelWorkSheet.UsedRange;
            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;
                // add all the columns
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }
                // add all the rows
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }


            //color
            int ColumnsCount = dataTable.Columns.Count;
            int RowsCount = dataTable.Rows.Count;
            object[] Header = new object[ColumnsCount];
            object[] RowsCol = new object[RowsCount];
            for (int i = 0; i < ColumnsCount; i++)
                Header[i] = dataTable.Columns[i].ColumnName;

            Excel.Range HeaderRange = excelWorkSheet.get_Range((Microsoft.Office.Interop.Excel.Range)(excelWorkSheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(excelWorkSheet.Cells[1, ColumnsCount]));
            HeaderRange.Value = Header;
            HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            HeaderRange.Font.Bold = true;

            //border
            RowsCount++;
            workSheet_range = excelWorkSheet.get_Range("A1", ((char)(ColumnsCount+64)).ToString()+ (RowsCount)); //+64 to get ascii character
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();

            // excelWorkBook.Save(); -> this will save to its default location
            excelWorkBook.SaveAs(path); // -> this will do the custom
            excelWorkBook.Close();
            excelApp.Quit();
            MessageBox.Show("Excel created successfully under "+ path + "!");
        }

        // T is a generic class
        static DataTable ConvertToDataTable<RowOfExportedExcel>(List<RowOfExportedExcel> models)
        {
            // creating a data table instance and typed it as our incoming model 
            // as I make it generic, if you want, you can make it the model typed you want. 
            DataTable dataTable = new DataTable(typeof(RowOfExportedExcel).Name);
            //Get all the properties of that model
            PropertyInfo[] Props = typeof(RowOfExportedExcel).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            // Loop through all the properties            
            // Adding Column name to our datatable

            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
                Console.WriteLine(prop.Name);
            }
            // Adding Row and its value to our dataTable
            foreach (RowOfExportedExcel item in models)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows  
                    values[i] = Props[i].GetValue(item, null);
                }
                // Finally add value to datatable  
                dataTable.Rows.Add(values);
            }
            foreach (DataRow dataRow in dataTable.Rows)
            {
                foreach (var item in dataRow.ItemArray)
                {
                    Console.WriteLine(item);
                }
            }
           
            return dataTable;
        }

        private static void PrintList2<RowOfExportedExcel>(List<RowOfExportedExcel> models)
        {
            {
                foreach (RowOfExportedExcel row in models)
                {
                    Console.WriteLine(row.ToString());
                }
            }
        }

        private List<RowOfImportedExcel> ReadFromExcel(string filePath) 
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

                    GenerateObjectWithRowValuesOfImportedExcel(str,row,cCnt);
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
            PrintList(rows);
            return rows;

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void PrintList(List<RowOfImportedExcel> list) 
        {
            foreach (RowOfImportedExcel row in list)
            {
                Console.WriteLine(row.ToString());
            }
        }

        //genetate RowOfImportedExcel with the values of imported excel
        private void GenerateObjectWithRowValuesOfImportedExcel(string value,RowOfImportedExcel row,int i)
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

        private List<RowOfExportedExcel> GenerateRowOfExportedExcel(List<RowOfImportedExcel>ListOfValues) 
        {
            int i;
            List<RowOfExportedExcel> rows = new List<RowOfExportedExcel>();
            foreach (RowOfImportedExcel rowOfImportedExcel in ListOfValues)
            {
                for (i = 1; i <= rowOfImportedExcel.numberOfUsers; i++ )
                {
                    RowOfExportedExcel row = new RowOfExportedExcel();
                    GenerateRowOfExportedExcelWithValues(row, rowOfImportedExcel, i);
                    rows.Add(row);
                }
            }
            return rows;
        }

        private void GenerateRowOfExportedExcelWithValues(RowOfExportedExcel row, RowOfImportedExcel rowOfImportedExcel,int i) 
        {
            row.startDate = rowOfImportedExcel.startDate;
            row.costCenterId = rowOfImportedExcel.costCenterId;
            row.firstName = "PDA" + i.ToString();
            row.lastName = "AB PDA" + i.ToString() + rowOfImportedExcel.storeName;
            row.userId = "PDA" + i.ToString() + rowOfImportedExcel.sapId;
            Console.WriteLine(row.ToString());
        }

        
    }
}
