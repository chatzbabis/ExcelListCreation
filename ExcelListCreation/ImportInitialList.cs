using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelListCreation
{
    class  ImportInitialList
    {
        public static List<RowOfImportedExcel> ReadFromExcel(string filePath)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str = null;
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
            List<string> headers = new List<string>();
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                //Console.WriteLine();
                RowOfImportedExcel row = new RowOfImportedExcel();
                nullCell = 0;
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    cellValue = (range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    //for NumbersOfUsers Column where cellValue is double 
                    /*
                    if (cCnt == 5)
                    {
                        int NumberOfUsers = Convert.ToInt32(cellValue);
                        //Console.Write(NumberOfUsers+" ");
                        row.numberOfUsers = NumberOfUsers;
                        Console.Write(NumberOfUsers+" ");
                        continue;
                    }
                    //for ArtemisId Column where cellValue is double 
                    if (cCnt == 6) 
                    {
                        string artemisId = (string)cellValue;
                        row.artemisId = artemisId;
                        Console.Write(artemisId + " ");
                        continue;
                    }*/

                    //str = (string)cellValue;

                    //ignore first row if it 's headers
                    if (cellValue != null)
                    {
                        if (rCnt == 1)
                        {
                            headers.Add((string)cellValue);
                            continue;
                        }
                    }
                    else 
                    {
                        Char cell = (Char)((true ? 65 : 97) + (cCnt - 1));
                        MessageBox.Show("The Cell in row: "+rCnt+" and column: "+cell.ToString()+" is null");
                        Application.Restart();
                        Environment.Exit(0);
                    }
                    int headerIndex = cCnt - 1;
                    GenerateObjectWithRowValuesOfImportedExcel(cellValue, row, headers[headerIndex],rCnt);
                    Console.Write(headers[cCnt - 1] + ": ");
                    Console.Write(cellValue.ToString() + " ");
                    Console.WriteLine();
                    Console.WriteLine();
                    /*
                    if (String.IsNullOrEmpty(str))
                    {
                        nullCell++;
                    }
                    if (nullCell == 7)
                    {
                        break;
                    }*/
                }

                //not add to list first row with headers
                /*if (!String.IsNullOrEmpty(str))
                {
                    if (str.Equals("StartDate"))
                    {
                        continue;
                    }
                }*/
                if (rCnt != 1)
                {
                    rows.Add(row);
                }
            }
            Console.WriteLine("---------------------------------------------");

            Prints.PrintList(rows);
            Console.WriteLine("---------------------------------------------");
            return rows;

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        //genetate RowOfImportedExcel with the values of imported excel
        public static void GenerateObjectWithRowValuesOfImportedExcel(Object cellValue, RowOfImportedExcel row, string header,int numberOfRow)
        {
            string value = null;
            validations(cellValue, header, numberOfRow);
            if (string.Equals(header, "StartDate", StringComparison.CurrentCultureIgnoreCase))
            {
                DateTime today = DateTime.Today;
                row.startDate = today.ToString(("dd/MM/yyyy"));
            }
            else if (string.Equals(header, "CostcenterID", StringComparison.CurrentCultureIgnoreCase))
            {
                value = (string)cellValue;
                row.costCenterId = value;
            }
            else if (string.Equals(header, "SAP_ID", StringComparison.CurrentCultureIgnoreCase))
            {
                value = (string)cellValue;
                row.sapId = value;
            }
            else if (string.Equals(header, "ArtemisID", StringComparison.CurrentCultureIgnoreCase))
            {
                value = cellValue.ToString();
                row.artemisId = value;
            }
            else if (string.Equals(header, "StoreName", StringComparison.CurrentCultureIgnoreCase))
            {
                value = (string)cellValue;
                row.storeName = value;
            }
            else if (string.Equals(header, "NumberOfUsers", StringComparison.CurrentCultureIgnoreCase))
            {
                int NumberOfUsers = Convert.ToInt32(cellValue);
                row.numberOfUsers = NumberOfUsers;
                //Console.Write(NumberOfUsers + " ");
            }


        }
        public static void validations(Object cellValue, string header, int numberOfRow) 
        {
            string value = null;
            
            if (string.Equals(header, "CostcenterID", StringComparison.CurrentCultureIgnoreCase))
            {
                value = cellValue.ToString();
                Regex rgx = new Regex(@"[Gg][Ss][0-9]{6}$");
                if (!rgx.IsMatch(value)) 
                {
                    MessageBox.Show("CostcenterID in row "+ numberOfRow + " must follow pattern 'GS+numbers' e.g 'GS000009'", "Error during list creation");
                    Application.Restart();
                    Environment.Exit(0);
                }
            }
            else if (string.Equals(header, "SAP_ID", StringComparison.CurrentCultureIgnoreCase))
            {

                value = cellValue.ToString();
                Regex rgx =new Regex( @"[Gg][0-9]{3}$");
                if (!rgx.IsMatch(value) /*|| !(value is string)*/ )
                {
                    MessageBox.Show("SAP_ID in row " + numberOfRow + " must follow pattern 'G+numbers' e.g 'G014'", "Error during list creation");
                    Application.Restart();
                    Environment.Exit(0);
                }
            }
            else if (string.Equals(header, "ArtemisID", StringComparison.CurrentCultureIgnoreCase))
            {
                value = cellValue.ToString();
                if (!value.All(char.IsDigit))
                {
                    MessageBox.Show("ArtemisID in row " + numberOfRow + " must contains only numbers", "Error during list creation");
                    Application.Restart();
                    Environment.Exit(0);
                }
            }
            else if (string.Equals(header, "NumberOfUsers", StringComparison.CurrentCultureIgnoreCase))
            {
                value = cellValue.ToString();
                if (!value.All(char.IsDigit))
                {
                    MessageBox.Show("NumberOfUsers in row " + numberOfRow + " must contains only numbers", "Error during list creation");
                    Application.Restart();
                    Environment.Exit(0);
                }
                value = cellValue.ToString();
                //Console.Write(NumberOfUsers + " ");
                int numberOfUsers = Convert.ToInt32(cellValue);
                if (numberOfUsers > 20 && numberOfUsers <= 30)
                {
                   // MessageBox.Show("For store in row " + numberOfRow + " you are requesting " + value + " users. Do you want to continue");
                    DialogResult dialogResult = MessageBox.Show("For store in row " + numberOfRow + " you are requesting " + value + " users. Do you want to continue", "Error during list creation", MessageBoxButtons.YesNo);
                    switch (dialogResult)
                    {
                        case DialogResult.Yes:
                            break;
                        case DialogResult.No:
                            Application.Restart();
                            Environment.Exit(0);
                            break;
                    }
                }
                else if (numberOfUsers > 30) 
                {
                    MessageBox.Show("Error in row " + numberOfRow + ". Maximum number of users is 30.", "Error during list creation");
                    Application.Restart();
                    Environment.Exit(0);
                }
            }

            bool IsDigitsOnly(string str)
            {
                foreach (char c in str)
                {
                    if (c < '0' || c > '9')
                        return false;
                }

                return true;
            }
        }
    }
}
