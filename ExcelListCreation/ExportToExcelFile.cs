using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Reflection;
using System.Windows.Forms;



namespace ExcelListCreation
{
    class ExportToExcelFile
    {
        public static string fileName = "SOP_Create_generic_users.xlsx";
        public static void ExportToExcel(List<RowOfExportedExcel> rowsOfExportedExcel, string ExcelSavingPath)
        {

            string path = ExcelSavingPath;
            string fullExcelSavingPath = path + "\\" + ExportToExcelFile.fileName;
            GenerateExcel(ConvertToDataTable(rowsOfExportedExcel), fullExcelSavingPath);
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
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
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
            workSheet_range = excelWorkSheet.get_Range("A1", ((char)(ColumnsCount + 64)).ToString() + (RowsCount)); //+64 to get ascii character
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();

            // excelWorkBook.Save(); -> this will save to its default location
            excelWorkBook.SaveAs(path); // -> this will do the custom
            excelWorkBook.Close();
            excelApp.Quit();
            MessageBox.Show("Excel created successfully under " + path + "!");
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
                //Console.WriteLine(prop.Name);
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
                    //Console.WriteLine(item);
                }
            }

            return dataTable;
        }
    }
}
