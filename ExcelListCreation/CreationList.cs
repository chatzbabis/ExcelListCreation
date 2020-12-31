using System;
using System.Collections.Generic;


namespace ExcelListCreation
{
    class CreationList
    {
        

        public static List<RowOfExportedExcel> GenerateRowOfExportedExcel(List<RowOfImportedExcel> ListOfValues)
        {
            int i;
            List<RowOfExportedExcel> rows = new List<RowOfExportedExcel>();
            foreach (RowOfImportedExcel rowOfImportedExcel in ListOfValues)
            {
                for (i = 1; i <= rowOfImportedExcel.numberOfUsers; i++)
                {
                    RowOfExportedExcel row = new RowOfExportedExcel();
                    GenerateRowOfExportedExcelWithValues(row, rowOfImportedExcel, i);
                    rows.Add(row);
                }
                RowOfExportedExcel bofUser = new RowOfExportedExcel();
                GenerateBofUser(bofUser, rowOfImportedExcel);
                rows.Add(bofUser);
            }
            //Console.WriteLine(rows.Capacity);
            return rows;
        }

        public static void GenerateRowOfExportedExcelWithValues(RowOfExportedExcel row, RowOfImportedExcel rowOfImportedExcel, int i)
        {
            DateTime today = DateTime.Today;

            row.StartDate = today.ToString(("dd/MM/yyyy"));
            row.CostCenterId = rowOfImportedExcel.costCenterId;
            row.FirstName = "PDA" + i.ToString();
            row.LastName = "AB PDA" + i.ToString() + rowOfImportedExcel.storeName;
            row.UserId = "PDA" + i.ToString() + rowOfImportedExcel.sapId;
            //Console.WriteLine(row.ToString());
        }

        public static void GenerateBofUser(RowOfExportedExcel bofUser, RowOfImportedExcel rowOfImportedExcel)
        {
            DateTime today = DateTime.Today;

            bofUser.Profile = "Greece: Basic User with Network (No mailbox, no In";
            bofUser.StartDate = today.ToString(("dd/MM/yyyy"));
            bofUser.CostCenterId = rowOfImportedExcel.costCenterId;
            bofUser.FirstName = rowOfImportedExcel.storeName + "_BOF";
            bofUser.LastName = "Store";
            bofUser.UserId = "BOFGR" + rowOfImportedExcel.artemisId;
            //Console.WriteLine(bofUser.ToString());
        }
    }
}
