using System.Data;
using System.Threading.Tasks;

namespace RetrenchmentSystemManagement.Services
{
    public interface IReportService
    {
        /// <summary>
        /// Generate report for given datatable (or use db fetch internally).
        /// Returns a tuple: Content bytes, MIME type, file extension (without dot).
        /// </summary>
        Task<(byte[] Content, string MimeType, string Extension)> GenerateReportAsync(DataTable data, string format = "PDF");

        /// <summary>
        /// Convenience: fetches report data by retrenchmentId and generates the report.
        /// </summary>
        Task<(byte[] Content, string MimeType, string Extension)> GenerateReportAsync(int? retrenchmentId, string format = "PDF");
    }
}
