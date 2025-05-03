using ExcelDataReader;
using ExcelUpload.Data;
using ExcelUpload.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ExcelUpload.Controllers
{
    public class ExcelUploadController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ExcelUploadController(ApplicationDbContext context)
        {
            _context = context;
        }


        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                
                if(file == null || file.Length == 0)
                {
                    return BadRequest("No file uploaded.");
                }
                var uploadsFolder = $"{Directory.GetCurrentDirectory()}/uploads";
                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }
                var filePath = Path.Combine(uploadsFolder, file.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        bool headerskipped = false;
                        do
                        {
                            while (reader.Read())
                            {
                                if (!headerskipped)
                                {
                                    headerskipped = true;
                                    continue; // Skip the header row
                                }

                                ExcelInfo info = new ExcelInfo();
                                info.CardNumber = reader.GetValue(0)?.ToString();
                                info.CustomerName = reader.GetValue(1)?.ToString();
                                info.CustomerLimit = Convert.ToDecimal(reader.GetValue(2));
                                info.CardBalance = Convert.ToDecimal(reader.GetValue(3));
                                info.LoanBalance = Convert.ToDecimal(reader.GetValue(4));
                                info.TotalBalance = Convert.ToDecimal(reader.GetValue(5));
                                info.MinDue = Convert.ToDecimal(reader.GetValue(6));
                                info.Buckets = Convert.ToInt32(reader.GetValue(7));
                                info.Pgs = Convert.ToInt32(reader.GetValue(8));
                                info.CustomerNumber = reader.GetValue(9)?.ToString();
                                info.InvoiceNumber = reader.GetValue(10)?.ToString();

                                _context.Add(info);
                                _context.SaveChanges();
                            }
                        } while (reader.NextResult());

                    }
                }
                return Ok(new { message = "File uploaded successfully." });
            }
            catch (Exception ex)
            {
                // Log error if you have logging set up
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }

}
