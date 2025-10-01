using System;

namespace RetrenchmentSystemManagement.Models
{
    public class RetrenchmentUploadModel
    {
        public int? RetrenchmentId { get; set; }
        public string EmpCode { get; set; }
        public string EmpCodeB { get; set; }
        public string EmpName { get; set; }
        public string EmpNameB { get; set; }
        public int DesigId { get; set; }
        public int SectId { get; set; }
        public DateTime dtJoin { get; set; }
        public DateTime? dtReleased { get; set; }
        public decimal Gross { get; set; }
        public decimal Basic { get; set; }
        public int PresentDays { get; set; }
        public decimal LastMonthSalary { get; set; }
        public decimal LastMonthSalary2 { get; set; }
        public string BankAcNo { get; set; }
        public decimal PayableServiceYear { get; set; }
        public string ActualServiceLength { get; set; }
    }
}
