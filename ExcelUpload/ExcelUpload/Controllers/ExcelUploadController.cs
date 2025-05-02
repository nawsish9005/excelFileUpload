using ExcelUpload.Data;
using ExcelUpload.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.IO;
using System.Threading.Tasks;

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
                return BadRequest("File is empty.");

            if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                return BadRequest("Only .xlsx files are supported.");

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);

                // ✅ Required for EPPlus 8+
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet == null || worksheet.Dimension == null)
                    return BadRequest("Worksheet is empty or corrupt.");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Skip header row
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

                    _context.ExcelInfos.Add(account);
                }

                await _context.SaveChangesAsync();
                return Ok("Excel data uploaded successfully.");
            }
            catch (Exception ex)
            {
                // Log error if you have logging set up
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }

}
