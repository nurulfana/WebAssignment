using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using NetAssessment.Models;

namespace NetAssessment.Pages_AssignmentList
{
    public class DashboardModel : PageModel
    {
        private readonly AppDbContext _context;

        public DashboardModel(AppDbContext context)
        {
            _context = context;
        }

        [BindProperty(SupportsGet = true)]
        public string? SearchString { get; set; }

        public IList<AssignmentList> AssignmentList { get;set; } = default!;

        public async Task OnGetAsync()
        {
            var query = _context.AssignmentList.AsQueryable();

            if (!string.IsNullOrEmpty(SearchString))
            {
                query = query.Where(a =>
                    a.Task.Contains(SearchString) ||
                    a.Customer.Contains(SearchString) ||
                    a.Mandays.ToString().Contains(SearchString) ||
                    a.StartDate.ToString().Contains(SearchString) ||
                    a.AddDays.ToString().Contains(SearchString) ||
                    a.EstimatedCompletionDate.ToString().Contains(SearchString) ||
                    a.PIC.Contains(SearchString) ||
                    a.Status.Contains(SearchString) ||
                    (a.CompletionDate != null && a.CompletionDate.Value.ToString().Contains(SearchString)) ||
                    (a.Remarks != null && a.Remarks.Contains(SearchString)) ||
                    (a.LastCommunicationDate.ToString().Contains(SearchString))
                );
            }

            AssignmentList = await _context.AssignmentList
            .FromSqlRaw("EXEC sp_GetAssignments")
            .ToListAsync();
        }

        public async Task<IActionResult> OnGetExportAsync()
        {
            var data = await _context.AssignmentList.ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Assignments");
                // Header
                worksheet.Cell(1, 1).Value = "Task";
                worksheet.Cell(1, 2).Value = "Customer";
                worksheet.Cell(1, 3).Value = "Mandays";
                worksheet.Cell(1, 4).Value = "StartDate";
                worksheet.Cell(1, 5).Value = "AddDays";
                worksheet.Cell(1, 6).Value = "EstimatedCompletionDate";
                worksheet.Cell(1, 7).Value = "PIC";
                worksheet.Cell(1, 8).Value = "Status";
                worksheet.Cell(1, 9).Value = "CompletionDate";
                worksheet.Cell(1, 10).Value = "Remarks";
                worksheet.Cell(1, 11).Value = "LastCommunicationDate";

                worksheet.Row(1).Style.Font.Bold = true;

                // Data
                for (int i = 0; i < data.Count; i++)
                {
                    var row = i + 2;
                    worksheet.Cell(row, 1).Value = data[i].Task;
                    worksheet.Cell(row, 2).Value = data[i].Customer;
                    worksheet.Cell(row, 3).Value = data[i].Mandays;
                    worksheet.Cell(row, 4).Value = data[i].StartDate.ToString("yyyy-MM-dd");
                    worksheet.Cell(row, 5).Value = data[i].AddDays;
                    worksheet.Cell(row, 6).Value = data[i].EstimatedCompletionDate.ToString("yyyy-MM-dd");
                    worksheet.Cell(row, 7).Value = data[i].PIC;
                    worksheet.Cell(row, 8).Value = data[i].Status;
                    worksheet.Cell(row, 9).Value = data[i].CompletionDate?.ToString("yyyy-MM-dd");
                    worksheet.Cell(row, 10).Value = data[i].Remarks;
                    worksheet.Cell(row, 11).Value = data[i].LastCommunicationDate.ToString("yyyy-MM-dd");

                    // Alignment
                    worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; 
                    worksheet.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; 
                    worksheet.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; 
                    worksheet.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; 
                    worksheet.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell(row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; 
                    worksheet.Cell(row, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; 
                    worksheet.Cell(row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; 
                    worksheet.Cell(row, 11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right; 
                }

                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    var fileName = $"AssignmentList_{DateTime.Now:yyyy-MM-dd}.xlsx";

                    return File(content,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                fileName);
                }
            }
        }

        [HttpPost]
        public async Task<JsonResult> OnPostDeleteAjax(int id)
        {
            try
            {
                using (var conn = _context.Database.GetDbConnection())
                {
                    await conn.OpenAsync();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "sp_DeleteAssignment";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Id", id));

                        var rowsAffected = await cmd.ExecuteScalarAsync();
                    }
                }

                return new JsonResult(new { success = true });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { success = false, message = ex.Message });
            }
        }

        public async Task<JsonResult> OnPostEditAjax([FromBody] AssignmentList updatedAssignment)
        {
            if (updatedAssignment == null)
            {
                return new JsonResult(new { success = false, message = "Invalid data" });
            }

            try
            {
                using (var conn = _context.Database.GetDbConnection())
                {
                    await conn.OpenAsync();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "sp_UpdateAssignment";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Id", updatedAssignment.Id));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Task", updatedAssignment.Task));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Customer", updatedAssignment.Customer));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Mandays", updatedAssignment.Mandays));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@StartDate", updatedAssignment.StartDate));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@AddDays", (object?)updatedAssignment.AddDays ?? DBNull.Value));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@EstimatedCompletionDate", updatedAssignment.EstimatedCompletionDate));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@PIC", updatedAssignment.PIC));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Status", updatedAssignment.Status));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@CompletionDate", (object?)updatedAssignment.CompletionDate ?? DBNull.Value));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@Remarks", (object?)updatedAssignment.Remarks ?? DBNull.Value));
                        cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@LastCommunicationDate", updatedAssignment.LastCommunicationDate));

                        var rowsAffected = await cmd.ExecuteScalarAsync(); 
                    }
                }

                return new JsonResult(new { success = true });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { success = false, message = ex.Message });
            }
        }
    }
}
