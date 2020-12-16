using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelListCreation
{
    class RowOfImportedExcel
    {
        public string startDate { get; set; }
        public string costCenterId { get; set; }
        public string storeName { get; set; }
        public string sapId { get; set; }
        public int numberOfUsers { get; set; }

        public RowOfImportedExcel()
        {
        }

        public RowOfImportedExcel(string startDate, string costCenterId, string storeName, string sapId, int numberOfUsers)
        {
            this.startDate = startDate;
            this.costCenterId = costCenterId;
            this.storeName = storeName;
            this.sapId = sapId;
            this.numberOfUsers = numberOfUsers;
        }

        public override string ToString()
        {
            return this.startDate+" "+this.costCenterId+" "+this.storeName+" "+this.sapId+" "+this.numberOfUsers;
        }
    }

}
