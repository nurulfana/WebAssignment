using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using NetAssessment.Models;

namespace NetAssessment.Pages_AssignmentList
{
    public class CreateModel : PageModel
    {
        private readonly AppDbContext _context;

        public CreateModel(AppDbContext context)
        {
            _context = context;
        }

        public IActionResult OnGet()
        {
            return Page();
        }

        [BindProperty]
        public AssignmentList AssignmentList { get; set; } = default!;

        [TempData]
        public string StatusMessage { get; set; }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                StatusMessage = "Error occurred while processing your request.";
                return Page();
            }

            try
            {
                using (var conn = _context.Database.GetDbConnection())
                {
                    await conn.OpenAsync();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "sp_CreateAssignment";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        cmd.Parameters.Add(new SqlParameter("@Task", AssignmentList.Task));
                        cmd.Parameters.Add(new SqlParameter("@Customer", AssignmentList.Customer));
                        cmd.Parameters.Add(new SqlParameter("@Mandays", AssignmentList.Mandays));
                        cmd.Parameters.Add(new SqlParameter("@StartDate", AssignmentList.StartDate));
                        cmd.Parameters.Add(new SqlParameter("@AddDays", AssignmentList.AddDays));
                        cmd.Parameters.Add(new SqlParameter("@EstimatedCompletionDate", AssignmentList.EstimatedCompletionDate));
                        cmd.Parameters.Add(new SqlParameter("@PIC", AssignmentList.PIC));
                        cmd.Parameters.Add(new SqlParameter("@Status", AssignmentList.Status));
                        cmd.Parameters.Add(new SqlParameter("@CompletionDate", (object)AssignmentList.CompletionDate ?? DBNull.Value));
                        cmd.Parameters.Add(new SqlParameter("@Remarks", (object)AssignmentList.Remarks ?? DBNull.Value));
                        cmd.Parameters.Add(new SqlParameter("@LastCommunicationDate", AssignmentList.LastCommunicationDate));

                        await cmd.ExecuteNonQueryAsync();
                    }
                }

                StatusMessage = "Assignment successfully created.";
                return RedirectToPage("./Create");
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                return Page();
            }
        }
    }
}
