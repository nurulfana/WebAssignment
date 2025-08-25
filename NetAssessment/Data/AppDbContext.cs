using Microsoft.EntityFrameworkCore;

public class AppDbContext : DbContext
    {
        public AppDbContext (DbContextOptions<AppDbContext> options)
            : base(options)
        {
        }

        public DbSet<NetAssessment.Models.AssignmentList> AssignmentList { get; set; } = default!;
    }
