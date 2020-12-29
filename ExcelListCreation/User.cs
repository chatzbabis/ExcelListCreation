using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelListCreation
{
    class User
    {
        public string EmployeeType { get; set; }
        public string Company { get; set; } 
        public string PerNr { get; set; } 
        public string Profile { get; set; }
        public string ExternalEmail { get; set; } 
        public string StartDate { get; set; }
        public string ExpirationDate { get; set; } 
        public string CostCenterId { get; set; }
        public string OrgUnit { get; set; } 
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Nationality { get; set; } 
        public string Language { get; set; } 
        public string ExternalCompany { get; set; } 
        public string Ad { get; set; } 
        public string UserId { get; set; }
        public string ExceptionNumber { get; set; } 


        public override string ToString()
        {
            return this.EmployeeType + "| " + this.Company + "| " + this.PerNr + "| " + this.Profile + "| " + this.ExternalEmail + "| " + this.StartDate + "| " + this.ExpirationDate + "| " + this.CostCenterId + "| " + this.OrgUnit + "| " + this.FirstName + "| " + this.LastName + "| " + this.Nationality + "| " + this.Language + "| " + this.ExternalCompany + "| " + this.Ad + "| " + this.UserId + "| " + this.ExceptionNumber;
        }
    }
}
