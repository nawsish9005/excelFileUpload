using ExcelUpload.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelUpload.Data
{
    public class ApplicationDbContext:DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }

        public DbSet<ExcelInfo> ExcelInfos { get; set; } = null!;
    }
}
