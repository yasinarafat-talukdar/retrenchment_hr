using Microsoft.Extensions.Configuration;
using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace RetrenchmentSystemManagement.Services
{
    public class DatabaseService : IDatabaseService
    {
        private readonly string _connectionString;
        private const string StagingTableName = "dbo.StgRetrenchments";

        public DatabaseService(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                                ?? throw new ArgumentNullException("DefaultConnection not found in configuration.");
        }

        public async Task UploadToStagingAsync(DataTable data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (data.Rows.Count == 0) return;

            using (var conn = new Microsoft.Data.SqlClient.SqlConnection(_connectionString))
            {
                await conn.OpenAsync();

                // Ensure staging table exists (create if not)
                var createSql = $@"
IF OBJECT_ID('{StagingTableName}', 'U') IS NULL
BEGIN
    CREATE TABLE {StagingTableName} (
        RetrenchmentId INT NULL,
        EmpCode NVARCHAR(50),
        EmpCodeB NVARCHAR(50),
        EmpName NVARCHAR(200),
        EmpNameB NVARCHAR(200),
        DesigId INT,
        SectId INT,
        dtJoin DATE,
        dtReleased DATE NULL,
        Gross DECIMAL(18,2),
        Basic DECIMAL(18,2),
        PresentDays INT,
        LastMonthSalary DECIMAL(18,2),
        LastMonthSalary2 DECIMAL(18,2),
        BankAcNo NVARCHAR(50),
        PayableServiceYear DECIMAL(5,2),
        ActualServiceLength NVARCHAR(200)
    );
END
ELSE
BEGIN
    TRUNCATE TABLE {StagingTableName};
END
";
                using (var cmd = new Microsoft.Data.SqlClient.SqlCommand(createSql, conn))
                {
                    cmd.CommandTimeout = 120;
                    await cmd.ExecuteNonQueryAsync();
                }

                // Bulk copy into staging table
                using (var bulk = new Microsoft.Data.SqlClient.SqlBulkCopy(conn))
                {
                    bulk.DestinationTableName = StagingTableName;
                    bulk.BulkCopyTimeout = 600;
                    // map columns by name when available
                    foreach (DataColumn col in data.Columns)
                    {
                        // If destination column exists, map it; otherwise ignore (no exception)
                        try
                        {
                            bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                        }
                        catch
                        {
                            // ignore mapping errors for unexpected columns
                        }
                    }
                    await bulk.WriteToServerAsync(data);
                }
            }
        }

        public async Task MergeStagingToMainAsync()
        {
            using (var conn = new Microsoft.Data.SqlClient.SqlConnection(_connectionString))
            {
                await conn.OpenAsync();

                var mergeSql = @"
MERGE INTO dbo.tblRetrenchmentsEmp_info AS Target
USING dbo.StgRetrenchments AS Src
ON (Target.RetrenchmentId = Src.RetrenchmentId AND Src.RetrenchmentId IS NOT NULL)
   OR (Target.EmpCode = Src.EmpCode AND Src.EmpCode IS NOT NULL)
WHEN MATCHED THEN
    UPDATE SET
        Target.EmpCode = Src.EmpCode,
        Target.EmpCodeB = Src.EmpCodeB,
        Target.EmpName = Src.EmpName,
        Target.EmpNameB = Src.EmpNameB,
        Target.DesigId = Src.DesigId,
        Target.SectId = Src.SectId,
        Target.dtJoin = Src.dtJoin,
        Target.dtReleased = Src.dtReleased,
        Target.Gross = Src.Gross,
        Target.Basic = Src.Basic,
        Target.PresentDays = Src.PresentDays,
        Target.LastMonthSalary = Src.LastMonthSalary,
        Target.LastMonthSalary2 = Src.LastMonthSalary2,
        Target.BankAcNo = Src.BankAcNo,
        Target.PayableServiceYear = Src.PayableServiceYear,
        Target.ActualServiceLength = Src.ActualServiceLength,
        Target.ModifiedAt = SYSUTCDATETIME()
WHEN NOT MATCHED BY TARGET THEN
    INSERT (EmpCode, EmpCodeB, EmpName, EmpNameB, DesigId, SectId, dtJoin, dtReleased, Gross, Basic, PresentDays, LastMonthSalary, LastMonthSalary2, BankAcNo, PayableServiceYear, ActualServiceLength, CreatedAt)
    VALUES (Src.EmpCode, Src.EmpCodeB, Src.EmpName, Src.EmpNameB, Src.DesigId, Src.SectId, Src.dtJoin, Src.dtReleased, Src.Gross, Src.Basic, Src.PresentDays, Src.LastMonthSalary, Src.LastMonthSalary2, Src.BankAcNo, Src.PayableServiceYear, Src.ActualServiceLength, SYSUTCDATETIME());
";

                using (var cmd = new Microsoft.Data.SqlClient.SqlCommand(mergeSql, conn))
                {
                    cmd.CommandTimeout = 600;
                    await cmd.ExecuteNonQueryAsync();
                }

                // Optionally truncate staging after merge (keeps table empty for next upload)
                var cleanup = $"TRUNCATE TABLE {StagingTableName};";
                using (var cmd2 = new Microsoft.Data.SqlClient.SqlCommand(cleanup, conn))
                {
                    cmd2.CommandTimeout = 60;
                    await cmd2.ExecuteNonQueryAsync();
                }
            }
        }

        public async Task UploadAndMergeAsync(DataTable data)
        {
            await UploadToStagingAsync(data);
            await MergeStagingToMainAsync();
        }

        public async Task RunRetrenchmentProcessAsync()
        {
            using (var conn = new Microsoft.Data.SqlClient.SqlConnection(_connectionString))
            using (var cmd = new Microsoft.Data.SqlClient.SqlCommand("dbo.sp_RetrenchmentProcess", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 600;
                await conn.OpenAsync();
                await cmd.ExecuteNonQueryAsync();
            }
        }

        public async Task<DataTable> GetReportDataAsync(int? retrenchmentId = null)
        {
            var dt = new DataTable();
            using (var conn = new Microsoft.Data.SqlClient.SqlConnection(_connectionString))
            using (var cmd = new Microsoft.Data.SqlClient.SqlCommand("dbo.sp_GetRetrenchmentEncashmentReport", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@RetrenchmentId", retrenchmentId.HasValue ? (object)retrenchmentId.Value : DBNull.Value);
                using (var da = new Microsoft.Data.SqlClient.SqlDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
            }
            return await Task.FromResult(dt);
        }
    }
}
