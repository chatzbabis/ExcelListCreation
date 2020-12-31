using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelListCreation
{
    class Prints
    {
        public static void PrintList2<RowOfExportedExcel>(List<RowOfExportedExcel> models)
        {
            {
                foreach (RowOfExportedExcel row in models)
                {
                    Console.WriteLine(row.ToString());
                }
            }
        }

        public static void PrintList(List<RowOfImportedExcel> list)
        {
            foreach (RowOfImportedExcel row in list)
            {
                Console.WriteLine(row.ToString());
            }
        }
    }
}
