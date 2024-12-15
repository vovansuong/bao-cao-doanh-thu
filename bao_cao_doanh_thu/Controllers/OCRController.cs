using bao_cao_doanh_thu.Models;
using bao_cao_doanh_thu.Services;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace bao_cao_doanh_thu.Controllers
{
    public class OCRController : Controller
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly OCRService _ocrService;
        private readonly ILogger<OCRController> _logger;
        public OCRController(IWebHostEnvironment webHostEnvironment, OCRService ocrService, ILogger<OCRController> logger)
        {
            _webHostEnvironment = webHostEnvironment;
            _ocrService = ocrService;
            _logger = logger;   
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View(new List<OCRModel>());
        }

        [HttpPost]
        public IActionResult Index(List<IFormFile> images)
        {
            var results = new List<OCRModel>();
            string uploadFolder = GetUploadFolder();

            try
            {
                foreach (var image in images)
                {
                    if (image.Length > 0)
                    {
                        // Lưu ảnh vào thư mục
                        string filePath = SaveImage(image, uploadFolder);

                        // Thực hiện OCR
                        string extractedText = _ocrService.ExtractTextFromImage(filePath);
                        Console.WriteLine($"Extracted Text: {extractedText}"); // Debug OCR output
                        _logger.LogInformation(extractedText);
                        // Định dạng dữ liệu OCR
                        string formattedText = FormatOCRText(extractedText);
                        Console.WriteLine($"Formatted Text: {formattedText}"); // Debug formatted text

                        // Trích xuất dữ liệu
                        var model = new OCRModel
                        {
                            Ngay = CleanCurrency(ExtractData(formattedText, "Ngày")),
                            NOW = CleanCurrency(ExtractAmount(ExtractData(formattedText, "NOW"))),
                            Be = CleanCurrency(ExtractAmount(ExtractData(formattedText, "Be"))),
                            GRAB = CleanCurrency(ExtractAmount(ExtractData(formattedText, "GRAB"))),
                            MOMO = CleanCurrency(ExtractAmount(ExtractData(formattedText, "MOMO"))),
                            Ca = CleanCurrency(ExtractData(formattedText, "Ca"))
                        };


                        results.Add(model);
                    }
                }

                // Lưu kết quả vào TempData
                TempData["OcrResults"] = JsonSerializer.Serialize(results);
                TempData.Keep("OcrResults");

                return View(results);
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, $"Đã xảy ra lỗi: {ex.Message}");
                return View(new List<OCRModel>());
            }
        }

        [HttpPost]
        public IActionResult Export()
        {
            try
            {
                if (TempData["OcrResults"] is string serializedResults)
                {
                    var ocrResults = JsonSerializer.Deserialize<List<OCRModel>>(serializedResults);
                    var fileContent = ExportToExcel(ocrResults);

                    return File(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "BaoCao.xlsx");
                }
                else
                {
                    TempData["ErrorMessage"] = "Không có dữ liệu để xuất.";
                }
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = $"Lỗi khi xuất file Excel: {ex.Message}";
            }

            return RedirectToAction("Index");
        }

        private string GetUploadFolder()
        {
            string uploadFolder = Path.Combine(_webHostEnvironment.WebRootPath, "uploads");
            if (!Directory.Exists(uploadFolder))
            {
                Directory.CreateDirectory(uploadFolder);
            }
            return uploadFolder;
        }

        private string SaveImage(IFormFile image, string uploadFolder)
        {
            string filePath = Path.Combine(uploadFolder, Guid.NewGuid().ToString() + "_" + image.FileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                image.CopyTo(stream);
            }
            return filePath;
        }

        private string FormatOCRText(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            var formattedLines = new List<string>();
            var lines = input.Split('\n', StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                string formattedLine = line.Trim();

                // Bỏ qua dòng không hợp lệ
                if (Regex.IsMatch(formattedLine, @"^\d+\s+\d+$")) continue;

                if (!string.IsNullOrWhiteSpace(formattedLine))
                {
                    formattedLines.Add(formattedLine);
                }
            }

            return string.Join("\n", formattedLines);
        }

        private string ExtractData(string input, string key)
        {
            // Regex khớp với số liệu chứa dấu phẩy hoặc chấm, kèm đơn vị "đ"
            var regex = new Regex($"{key}[.:\\s]*\\d*\\s*(\\d+(?:[,.]\\d+)*\\.?\\d*đ?)", RegexOptions.IgnoreCase);
            var match = regex.Match(input);
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }




        private string ExtractAmount(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            // Regex để trích xuất số có thể có dấu chấm hoặc phẩy
            var regex = new Regex(@"[\d,.]+(?:\.\d+)?");
            var match = regex.Match(input);

            if (match.Success)
            {
                // Loại bỏ tất cả các ký tự không mong muốn như "đ"
                return match.Value.Replace("đ", "").Trim();
            }

            return string.Empty;
        }


        private byte[] ExportToExcel(List<OCRModel> ocrData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("BaoCao");

                // Header
                worksheet.Cells[1, 1].Value = "Ngày";
                worksheet.Cells[1, 2].Value = "Shopeefood";
                worksheet.Cells[1, 3].Value = "Be";
                worksheet.Cells[1, 4].Value = "GRAB";
                worksheet.Cells[1, 5].Value = "MOMO";
                worksheet.Cells[1, 6].Value = "CA";

                // Dữ liệu
                int row = 2;
                foreach (var data in ocrData)
                {
                    worksheet.Cells[row, 1].Value = data.Ngay;
                    worksheet.Cells[row, 2].Value = data.NOW;
                    worksheet.Cells[row, 3].Value = data.Be;
                    worksheet.Cells[row, 4].Value = data.GRAB;
                    worksheet.Cells[row, 5].Value = data.MOMO;
                    worksheet.Cells[row, 6].Value = data.Ca;
                    row++;
                }

                // Ghi vào memory stream
                var stream = new MemoryStream();
                package.SaveAs(stream);
                return stream.ToArray();
            }
        }

        private string CleanCurrency(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            // Xóa ký tự "đ" và các khoảng trắng thừa
            return input.Replace("đ", "").Trim();
        }

    }
}
