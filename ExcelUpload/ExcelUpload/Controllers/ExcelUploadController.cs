using ExcelUpload.Data;
using ExcelUpload.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace ExcelUpload.Controllers
{
    public class ExcelUploadController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _env;

        public ExcelUploadController(ApplicationDbContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("upload")]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File is empty");

            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Row 1 = headers
            {
                var account = new ExcelInfo
                {
                    CardNumber = worksheet.Cells[row, 1].Text,
                    CustomerName = worksheet.Cells[row, 2].Text,
                    CustomerLimit = decimal.TryParse(worksheet.Cells[row, 3].Text, out var limit) ? limit : 0,
                    CardBalance = decimal.TryParse(worksheet.Cells[row, 4].Text, out var card) ? card : 0,
                    LoanBalance = decimal.TryParse(worksheet.Cells[row, 5].Text, out var loan) ? loan : 0,
                    TotalBalance = decimal.TryParse(worksheet.Cells[row, 6].Text, out var total) ? total : 0,
                    MinDue = decimal.TryParse(worksheet.Cells[row, 7].Text, out var due) ? due : 0,
                    Buckets = int.TryParse(worksheet.Cells[row, 8].Text, out var buckets) ? buckets : 0,
                    Pgs = int.TryParse(worksheet.Cells[row, 9].Text, out var pgs) ? pgs : 0,
                    CustomerNumber = worksheet.Cells[row, 10].Text,
                    InvoiceNumber = worksheet.Cells[row, 11].Text
                };

                // Optional: Update existing or just add new
                _context.ExcelInfos.Add(account);
            }

            await _context.SaveChangesAsync();
            return Ok("Excel data uploaded.");
        }

    }
}
