using Tesseract;

namespace bao_cao_doanh_thu.Services
{
    public class OCRService
    {
        public string ExtractTextFromImage(string imagePath)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "vie", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(imagePath))
                    {
                        using (var page = engine.Process(img))
                        {
                            return page.GetText();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
}
