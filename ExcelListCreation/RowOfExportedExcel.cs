using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelListCreation
{
    class RowOfExportedExcel
    {
        public string employeeType { get; set; } ="Generic_Account";

        public string company { get; set; } = "Alfa-Beta";
        public string perNr { get; set; } = "N/A";
        public string profile { get; set; } = "Greece: PDA Generic account (network, internet, no";
        public string externalEmail { get; set; } = "N/A";
        public string startDate { get; set; }
        public string expirationDate { get; set; } = null;
        public string costCenterId { get; set; }
        public string orgUnit { get; set; } = "Greece – AB Stores";
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string nationality { get; set; } = "Greece";
        public string language { get; set; } = "Greek";
        public string externalCompany { get; set; } = "N/A";
        public string ad { get; set; } = "1";
        public string userId { get; set; }
        public string exceptionNumber { get; set; } = "EXP-01380";


        public override string ToString()
        {
            return this.employeeType + "| " + this.company + "| " + this.perNr + "| " + this.profile + "| " + this.externalEmail + "| " + this.startDate + "| " + this.expirationDate + "| " + this.costCenterId + "| " + this.orgUnit + "| " + this.firstName + "| " + this.lastName + "| " + this.nationality + "| " + this.language + "| " + this.externalCompany + "| " + this.ad + "| " + this.userId + "| " + this.exceptionNumber;
        }
    }
}
