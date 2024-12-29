using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Word.Models;
using IronOcr;
using Microsoft.AspNetCore.Http.Metadata;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace Word.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _env;
        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        string temp_path(string extension)
        {
            string randomfile = Path.GetRandomFileName();
            string filename = Path.ChangeExtension(randomfile, extension);
            string path = Path.Combine(Path.GetTempPath(), filename);
            return path;
        }

        [HttpPost]
        public async Task<IActionResult> getfile(IFormFile _file)
        {
            try {
                string path = temp_path(Path.GetExtension(_file.FileName));
                using (var filestream = new FileStream(path, FileMode.Create))
                {
                    await _file.CopyToAsync(filestream);
                }
                IronOcr.License.LicenseKey = "IRONSUITE.HONG99152.GMAIL.COM.32115-24C0F9A90C-P3X6S-MLD7ONFQJR5N-XNI36M2M2DWS-KEXTIFHBNNH5-OFYVKZYWK3AP-DR6X453GNQUR-GHK4KIX67NBP-OY72BY-TI6K7AIWV7GOEA-DEPLOYMENT.TRIAL-FJI3XF.TRIAL.EXPIRES.28.JAN.2025";
                var ocr = new IronTesseract();
                ocr.Language = OcrLanguage.VietnameseBest;
                string result = "";
                using (var ocrInput = new OcrInput())
                {
                    if (_file.ContentType == "application/pdf") ocrInput.LoadPdf(path);
                    else ocrInput.LoadImage(path);
                    // Optionally Apply Filters if needed:
                    // ocrInput.Deskew();  // use only if image not straight
                    // ocrInput.DeNoise(); // use only if image contains digital noise

                    var ocrResult = ocr.Read(ocrInput);
                    result = ocrResult.Text;
                }
                string randomfile = Path.GetRandomFileName();

                string word_path = Path.Combine(Path.GetTempPath(), "docx");
                Regex reg = new Regex(@"[\w\s\W]*");
                string[] match = result.Split(new[]{"\r\n"}, StringSplitOptions.None);
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(word_path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    // Add a main document part
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    // Create a paragraph and add some text
                    foreach (string text in match)
                    {
                        if (text == "") continue;
                        Paragraph paragraph = new Paragraph();
                        Run run = new Run();
                        run.AppendChild(new Text(text));
                        paragraph.AppendChild(run);

                        // Append the paragraph to the body of the document
                        mainPart.Document.Body.AppendChild(paragraph);
                    }
                    // Save the document
                    mainPart.Document.Save();
                }
                var filebytes = System.IO.File.ReadAllBytes(word_path);  
                return File(filebytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "new_word.docx");
            } catch(Exception ex) {
                return Content(ex.Message);
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
