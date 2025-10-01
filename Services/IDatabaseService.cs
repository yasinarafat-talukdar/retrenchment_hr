using System.Data;
using System.Threading.Tasks;

namespace RetrenchmentSystemManagement.Services
{
    public interface IDatabaseService
    {
        /// <summary>
        /// Ensure staging table exists and bulk insert the provided DataTable into it.
        /// </summary>
        Task UploadToStagingAsync(DataTable data);

        /// <summary>
        /// Merge staging table data into main table (tblRetrenchmentsEmp_info).
        /// </summary>
        Task MergeStagingToMainAsync();

        /// <summary>
        /// Convenience: upload DataTable to staging and merge into main in one call.
        /// </summary>
        Task UploadAndMergeAsync(DataTable data);

        /// <summary>
        /// Execute sp_RetrenchmentProcess stored procedure.
        /// </summary>
        Task RunRetrenchmentProcessAsync();

        /// <summary>
        /// Get report data using sp_GetRetrenchmentEncashmentReport (optionally filtered by id).
        /// </summary>
        Task<DataTable> GetReportDataAsync(int? retrenchmentId = null);
    }
}
