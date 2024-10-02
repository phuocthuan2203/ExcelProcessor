using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;
using ExcelProcessor.Models;

namespace ExcelProcessor.Controllers
{
    public class ExcelController : Controller
    {
        private const string ExcelSessionKey = "ExcelFilePath";
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    _logger.LogWarning("File not selected for upload");
                    return Content("File not selected");
                }

                var filePath = Path.GetTempFileName();
                HttpContext.Session.SetString(ExcelSessionKey, filePath);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                _logger.LogInformation("File uploaded successfully: {FilePath}", filePath);
                return View("Calculate");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while uploading file");
                return RedirectToAction("Index", new { error = "An error occurred while uploading the file." });
            }
        }

        [HttpGet]
        public IActionResult Calculate()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Calculate(decimal firstNum, decimal secondNum)
        {
            if (firstNum == 0 && secondNum == 0)
            {
                ModelState.AddModelError("", "At least one number should be non-zero.");
                return View();
            }

            try
            {
                var filePath = HttpContext.Session.GetString(ExcelSessionKey);
                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    return RedirectToAction("Index", new { error = "Please upload an Excel file first." });
                }

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension?.Rows ?? 1;

                    worksheet.Cells[rowCount + 1, 1].Value = firstNum;
                    worksheet.Cells[rowCount + 1, 2].Value = secondNum;
                    worksheet.Cells[rowCount + 1, 3].Formula = $"=A{rowCount + 1}+B{rowCount + 1}";

                    package.Save();
                }

                return RedirectToAction("Preview");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while calculating");
                return RedirectToAction("Index", new { error = "An error occurred while processing the numbers." });
            }
        }

        public IActionResult Preview()
        {
            var filePath = HttpContext.Session.GetString(ExcelSessionKey);
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
            {
                return RedirectToAction("Index", new { error = "Please upload an Excel file first." });
            }

            var model = new List<ExcelData>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.Calculate();
                var rowCount = worksheet.Dimension?.Rows ?? 0;

                for (int row = 2; row <= rowCount; row++)
                {
                    var firstNum = Convert.ToDecimal(worksheet.Cells[row, 1].Value ?? 0);
                    var secondNum = Convert.ToDecimal(worksheet.Cells[row, 2].Value ?? 0);
                    var result = Convert.ToDecimal(worksheet.Cells[row, 3].Value ?? 0);

                    model.Add(new ExcelData
                    {
                        RowNumber = row - 1,
                        FirstNum = firstNum,
                        SecondNum = secondNum,
                        Result = result,
                        Operation = "+"
                    });

                    // Ensure the calculated value is saved in the cell
                    worksheet.Cells[row, 3].Value = result;
                }

                package.Save();
            }

            return View(model);
        }

        public IActionResult Download()
        {
            var filePath = HttpContext.Session.GetString(ExcelSessionKey);
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
            {
                return RedirectToAction("Index", new { error = "Please upload an Excel file first." });
            }
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                // Remove empty rows
                for (int row = rowCount; row >= 2; row--)
                {
                    if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text) &&
                    string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text) &&
                    (string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Text) || worksheet.Cells[row, 3].Value.ToString() == "0"))
                    {
                        worksheet.DeleteRow(row);
                    }
                }
                // Save the changes
                var memoryStream = new MemoryStream();
                package.SaveAs(memoryStream);
                memoryStream.Position = 0;
                return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ProcessedExcel.xlsx");
            }
        }
    }
}