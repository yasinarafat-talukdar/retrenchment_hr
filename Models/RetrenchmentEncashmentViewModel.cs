using System;

namespace RetrenchmentSystemManagement.Models
{
    public class RetrenchmentEncashmentViewModel
    {
        public int RetrenchmentId { get; set; }
        public string EmpCode { get; set; }
        public string EmpName { get; set; }
        public string Designation { get; set; }
        public string Section { get; set; }
        public DateTime dtJoin { get; set; }
        public DateTime? dtReleased { get; set; }
        public decimal Gross { get; set; }
        public decimal Basic { get; set; }
        public int PresentDays { get; set; }
        public decimal LastMonthSalary { get; set; }
        public decimal LastMonthSalary2 { get; set; }
        public decimal PayableServiceYear { get; set; }

        // Calculated Fields from sp_RetrenchmentProcess
        public decimal ELAmount { get; set; }
        public decimal NoticePeriodAmount { get; set; }
        public decimal RetrenchmentBenefitAmount { get; set; }
        public decimal TotalPayableAmount { get; set; }

        public string BankAcNo { get; set; }
    }
}
