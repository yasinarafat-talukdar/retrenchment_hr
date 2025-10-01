using AspNetCore.Reporting;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using Microsoft.Data.SqlClient; // Microsoft.Data.SqlClient ব্যবহার করুন
using System.IO;
using System.Collections.Generic;

namespace RetrenchmentSystemManagement.Controllers
{
    public class ReportController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public ReportController(IConfiguration configuration, IWebHostEnvironment webHostEnvironment)
        {
            _configuration = configuration;
            _webHostEnvironment = webHostEnvironment;
        }

        // View for report selection
        public IActionResult ViewReport()
        {
            return View();
        }

        // Action to generate report
        [HttpPost]
        public IActionResult GenerateReport(string reportType)
        {
            // 1. Fetch data from stored procedure
            string connectionString = _configuration.GetConnectionString("DefaultConnection");
            DataTable dt = new DataTable();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("sp_GetRetrenchmentEncashmentReport", con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                }
            }

            // 2. RDLC file path
            string rdlcFilePath = Path.Combine(
                _webHostEnvironment.WebRootPath ?? _webHostEnvironment.ContentRootPath,
                "Reports",
                "RetrenchmentReport.rdlc"
            );

            if (!System.IO.File.Exists(rdlcFilePath))
                return NotFound($"RDLC file not found at path: {rdlcFilePath}");

            // 3. LocalReport & add DataSource
            LocalReport report = new LocalReport(rdlcFilePath);
            report.AddDataSource("DataSet1", dt); // DataSet1 = RDLC dataset name

            // 4. Render report
            ReportResult reportResult;
            string mimeType;
            string extension;

            var parameters = new Dictionary<string, string>(); // report parameters if any

            if (string.Equals(reportType, "PDF", System.StringComparison.OrdinalIgnoreCase))
            {
                reportResult = report.Execute(RenderType.Pdf, 1, parameters);
                mimeType = "application/pdf";
                extension = "pdf";
            }
            else // Excel
            {
                reportResult = report.Execute(RenderType.Excel, 1, parameters);
                mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                extension = "xlsx";
            }

            byte[] result = reportResult.MainStream;

            // 5. Return file
            return File(result, mimeType, $"RetrenchmentReport.{extension}");
        }
    }
}
