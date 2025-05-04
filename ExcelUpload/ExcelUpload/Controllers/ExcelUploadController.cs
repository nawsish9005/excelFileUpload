using ExcelDataReader;
using ExcelUpload.Data;
using ExcelUpload.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUpload.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploadController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly string uploadsFolder = $"{Directory.GetCurrentDirectory()}/uploads";

        public ExcelUploadController(ApplicationDbContext context)
        {
            _context = context;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }
        }

        // POST: Upload file and store data
        [HttpPost]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest("No file uploaded.");

                var filePath = Path.Combine(uploadsFolder, file.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
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
                                continue; // Skip header
                            }

                            ExcelInfo info = new ExcelInfo
                            {
                                CardNumber = reader.GetValue(0)?.ToString(),
                                CustomerName = reader.GetValue(1)?.ToString(),
                                CustomerLimit = Convert.ToDecimal(reader.GetValue(2)),
                                CardBalance = Convert.ToDecimal(reader.GetValue(3)),
                                LoanBalance = Convert.ToDecimal(reader.GetValue(4)),
                                TotalBalance = Convert.ToDecimal(reader.GetValue(5)),
                                MinDue = Convert.ToDecimal(reader.GetValue(6)),
                                Buckets = Convert.ToInt32(reader.GetValue(7)),
                                Pgs = Convert.ToInt32(reader.GetValue(8)),
                                CustomerNumber = reader.GetValue(9)?.ToString(),
                                InvoiceNumber = reader.GetValue(10)?.ToString(),
                                SourceFile = file.FileName
                            };

                            _context.ExcelInfos.Add(info);
                        }
                    } while (reader.NextResult());

                    await _context.SaveChangesAsync();
                }

                return Ok(new { message = "File uploaded and data saved successfully." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        // GET: List of uploaded file names
        [HttpGet("files")]
        public IActionResult GetUploadedFileNames()
        {
            var fileNames = Directory.GetFiles(uploadsFolder)
                                     .Select(Path.GetFileName)
                                     .ToList();

            return Ok(fileNames);
        }

        // GET: Data by filename
        [HttpGet("data")]
        public IActionResult GetDataByFile(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                return BadRequest("Filename is required.");

            var records = _context.ExcelInfos
                                  .Where(e => e.SourceFile == filename)
                                  .ToList();

            return Ok(records);
        }
    }
}
