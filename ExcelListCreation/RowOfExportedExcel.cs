using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelListCreation
{
    class RowOfExportedExcel : User
    {
        public string EmployeeType { get; set; } ="Generic_Account";

        public string Company { get; set; } = "Alfa-Beta";
        public string PerNr { get; set; } = "N/A";
        public string Profile { get; set; } = "Greece: PDA Generic account (network, internet, no";
        public string ExternalEmail { get; set; } = "N/A";
        public string StartDate { get; set; }
        public string ExpirationDate { get; set; } = null;
        public string CostCenterId { get; set; }
        public string OrgUnit { get; set; } = "Greece – AB Stores";
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Nationality { get; set; } = "Greece";
        public string Language { get; set; } = "Greek";
        public string ExternalCompany { get; set; } = "N/A";
        public string Ad { get; set; } = "1";
        public string UserId { get; set; }
        public string ExceptionNumber { get; set; } = "EXP-01380";


        public override string ToString()
        {
            return this.EmployeeType + "| " + this.Company + "| " + this.PerNr + "| " + this.Profile + "| " + this.ExternalEmail + "| " + this.StartDate + "| " + this.ExpirationDate + "| " + this.CostCenterId + "| " + this.OrgUnit + "| " + this.FirstName + "| " + this.LastName + "| " + this.Nationality + "| " + this.Language + "| " + this.ExternalCompany + "| " + this.Ad + "| " + this.UserId + "| " + this.ExceptionNumber;
        }
    }
}
