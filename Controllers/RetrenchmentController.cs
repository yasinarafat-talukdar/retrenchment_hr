using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using RetrenchmentSystemManagement.Models;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace RetrenchmentSystemManagement.Controllers
{
    public class RetrenchmentController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public RetrenchmentController(IConfiguration configuration, IWebHostEnvironment webHostEnvironment)
        {
            _configuration = configuration;
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Index() => View();


        public IActionResult Upload() => View();

        [HttpPost]
        public async Task<IActionResult> UploadData(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { success = false, message = "No file uploaded." });

            try
            {
                DataTable dt = ReadExcelToDataTable(file);

                if (dt.Rows.Count == 0)
                    return BadRequest(new { success = false, message = "Uploaded file contains no data." });

                string connectionString = _configuration.GetConnectionString("DefaultConnection");
                using (var conn = new SqlConnection(connectionString))
                {
                    await conn.OpenAsync();

                    foreach (DataRow row in dt.Rows)
                    {
                        string sql = @"
                            IF EXISTS(SELECT 1 FROM dbo.tblRetrenchmentsEmp_info WHERE EmpCode=@EmpCode)
                            BEGIN
                                UPDATE dbo.tblRetrenchmentsEmp_info
                                SET EmpCodeB=@EmpCodeB, EmpName=@EmpName, EmpNameB=@EmpNameB,
                                    DesigId=@DesigId, SectId=@SectId, dtJoin=@dtJoin, dtReleased=@dtReleased,
                                    Gross=@Gross, Basic=@Basic, PresentDays=@PresentDays,
                                    LastMonthSalary=@LastMonthSalary, LastMonthSalary2=@LastMonthSalary2,
                                    BankAcNo=@BankAcNo, PayableServiceYear=@PayableServiceYear,
                                    ActualServiceLength=@ActualServiceLength
                                WHERE EmpCode=@EmpCode
                            END
                            ELSE
                            BEGIN
                                INSERT INTO dbo.tblRetrenchmentsEmp_info
                                (EmpCode, EmpCodeB, EmpName, EmpNameB, DesigId, SectId, dtJoin, dtReleased,
                                 Gross, Basic, PresentDays, LastMonthSalary, LastMonthSalary2, BankAcNo, PayableServiceYear, ActualServiceLength)
                                VALUES
                                (@EmpCode, @EmpCodeB, @EmpName, @EmpNameB, @DesigId, @SectId, @dtJoin, @dtReleased,
                                 @Gross, @Basic, @PresentDays, @LastMonthSalary, @LastMonthSalary2, @BankAcNo, @PayableServiceYear, @ActualServiceLength)
                            END";

                        using var cmd = new SqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@EmpCode", row["EmpCode"]);
                        cmd.Parameters.AddWithValue("@EmpCodeB", row["EmpCodeB"]);
                        cmd.Parameters.AddWithValue("@EmpName", row["EmpName"]);
                        cmd.Parameters.AddWithValue("@EmpNameB", row["EmpNameB"]);
                        cmd.Parameters.AddWithValue("@DesigId", row["DesigId"]);
                        cmd.Parameters.AddWithValue("@SectId", row["SectId"]);
                        cmd.Parameters.AddWithValue("@dtJoin", row["dtJoin"]);
                        cmd.Parameters.AddWithValue("@dtReleased", row["dtReleased"]);
                        cmd.Parameters.AddWithValue("@Gross", row["Gross"]);
                        cmd.Parameters.AddWithValue("@Basic", row["Basic"]);
                        cmd.Parameters.AddWithValue("@PresentDays", row["PresentDays"]);
                        cmd.Parameters.AddWithValue("@LastMonthSalary", row["LastMonthSalary"]);
                        cmd.Parameters.AddWithValue("@LastMonthSalary2", row["LastMonthSalary2"]);
                        cmd.Parameters.AddWithValue("@BankAcNo", row["BankAcNo"]);
                        cmd.Parameters.AddWithValue("@PayableServiceYear", row["PayableServiceYear"]);
                        cmd.Parameters.AddWithValue("@ActualServiceLength", row["ActualServiceLength"]);

                        await cmd.ExecuteNonQueryAsync();
                    }
                }

                return Ok(new { success = true, message = "Upload completed.", rows = dt.Rows.Count });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Error during upload.", error = ex.Message });
            }
        }





        [HttpPost]
        public async Task<IActionResult> ProcessRetrenchments()
        {
            try
            {
                string connectionString = _configuration.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(connectionString);
                using var cmd = new SqlCommand("dbo.sp_RetrenchmentProcess", conn)
                {
                    CommandType = CommandType.StoredProcedure,
                    CommandTimeout = 600
                };
                await conn.OpenAsync();
                await cmd.ExecuteNonQueryAsync();

                return Ok(new { success = true, message = "Retrenchment process completed." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    success = false,
                    message = "Processing failed.",
                    error = ex.Message,
                    stackTrace = ex.StackTrace
                });
            }
        }


        #region Helper Methods

        private DataTable ReadExcelToDataTable(IFormFile file)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("EmpCode", typeof(string));
            dt.Columns.Add("EmpCodeB", typeof(string));
            dt.Columns.Add("EmpName", typeof(string));
            dt.Columns.Add("EmpNameB", typeof(string));
            dt.Columns.Add("DesigId", typeof(int));
            dt.Columns.Add("SectId", typeof(int));
            dt.Columns.Add("dtJoin", typeof(DateTime));
            dt.Columns.Add("dtReleased", typeof(object)); // DBNull handled
            dt.Columns.Add("Gross", typeof(decimal));
            dt.Columns.Add("Basic", typeof(decimal));
            dt.Columns.Add("PresentDays", typeof(int));
            dt.Columns.Add("LastMonthSalary", typeof(decimal));
            dt.Columns.Add("LastMonthSalary2", typeof(decimal));
            dt.Columns.Add("BankAcNo", typeof(string));
            dt.Columns.Add("PayableServiceYear", typeof(decimal));
            dt.Columns.Add("ActualServiceLength", typeof(string));

            using (var stream = file.OpenReadStream())
            using (var workbook = new XLWorkbook(stream))
            {
                var ws = workbook.Worksheet(1);
                int row = 2; // header row

                while (!ws.Row(row).IsEmpty())
                {
                    var dr = dt.NewRow();
                    dr["EmpCode"] = TryGetString(ws.Cell(row, 1));
                    dr["EmpCodeB"] = TryGetString(ws.Cell(row, 2));
                    dr["EmpName"] = TryGetString(ws.Cell(row, 3));
                    dr["EmpNameB"] = TryGetString(ws.Cell(row, 4));
                    dr["DesigId"] = TryGetInt(ws.Cell(row, 5));
                    dr["SectId"] = TryGetInt(ws.Cell(row, 6));
                    dr["dtJoin"] = TryGetDateTime(ws.Cell(row, 7));
                    dr["dtReleased"] = (object?)TryGetNullableDateTime(ws.Cell(row, 8)) ?? DBNull.Value;
                    dr["Gross"] = TryGetDecimal(ws.Cell(row, 9));
                    dr["Basic"] = TryGetDecimal(ws.Cell(row, 10));
                    dr["PresentDays"] = TryGetInt(ws.Cell(row, 11));
                    dr["LastMonthSalary"] = TryGetDecimal(ws.Cell(row, 12));
                    dr["LastMonthSalary2"] = TryGetDecimal(ws.Cell(row, 13));
                    dr["BankAcNo"] = TryGetString(ws.Cell(row, 14));
                    dr["PayableServiceYear"] = TryGetDecimal(ws.Cell(row, 15));
                    dr["ActualServiceLength"] = TryGetString(ws.Cell(row, 16));

                    dt.Rows.Add(dr);
                    row++;
                }
            }

            return dt;
        }

        private static string TryGetString(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return string.Empty;
            return cell.GetString().Trim();
        }

        private static int TryGetInt(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return 0;
            if (int.TryParse(cell.GetString().Trim(), out int v)) return v;
            try { return cell.GetValue<int>(); } catch { return 0; }
        }

        private static decimal TryGetDecimal(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return 0m;
            if (decimal.TryParse(cell.GetString().Trim(), out decimal v)) return v;
            try { return cell.GetValue<decimal>(); } catch { return 0m; }
        }

        private static DateTime TryGetDateTime(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return DateTime.MinValue;
            try { return cell.GetDateTime(); }
            catch
            {
                if (DateTime.TryParse(cell.GetString().Trim(), out DateTime dt)) return dt;
                return DateTime.MinValue;
            }
        }

        private static DateTime? TryGetNullableDateTime(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return null;
            try { return cell.GetDateTime(); }
            catch { if (DateTime.TryParse(cell.GetString().Trim(), out DateTime dt)) return dt; return null; }
        }

        #endregion
    }
}
