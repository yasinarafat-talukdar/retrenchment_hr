using AspNetCore.Reporting;
using Microsoft.AspNetCore.Hosting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace RetrenchmentSystemManagement.Services
{
    public class ReportService : IReportService
    {
        private readonly IWebHostEnvironment _env;
        private readonly IDatabaseService _dbService;
        private readonly string _reportRelativePath = Path.Combine("Reports", "RetrenchmentEncashmentReport.rdlc");
        private const string DataSetNameInRdlc = "RetrenchmentDS"; // RDLC dataset name

        public ReportService(IWebHostEnvironment env, IDatabaseService dbService)
        {
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _dbService = dbService ?? throw new ArgumentNullException(nameof(dbService));
        }

        public async Task<(byte[] Content, string MimeType, string Extension)> GenerateReportAsync(DataTable data, string format = "PDF")
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            var rdlcPath = Path.Combine(_env.WebRootPath ?? _env.ContentRootPath, _reportRelativePath);
            if (!File.Exists(rdlcPath))
                throw new FileNotFoundException($"RDLC file not found at path {rdlcPath}");

            LocalReport report = new LocalReport(rdlcPath);
            report.AddDataSource(DataSetNameInRdlc, data);

            Dictionary<string, string> parameters = new Dictionary<string, string>();

            if (string.Equals(format, "PDF", StringComparison.OrdinalIgnoreCase))
            {
                var result = report.Execute(RenderType.Pdf, 1, parameters);
                return (result.MainStream, "application/pdf", "pdf");
            }
            else // Excel
            {
                var result = report.Execute(RenderType.Excel, 1, parameters);
                return (result.MainStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
            }
        }

        public async Task<(byte[] Content, string MimeType, string Extension)> GenerateReportAsync(int? retrenchmentId, string format = "PDF")
        {
            var dt = await _dbService.GetReportDataAsync(retrenchmentId);
            return await GenerateReportAsync(dt, format);
        }
    }
}
