using GeneralWorkMVC.Models;
using iText.Kernel.Pdf;
using System.IO;
using System.Text;
using System.Web.Mvc;

public class PdfController : Controller
{
    // GET: Pdf
    public ActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public ActionResult Index(PdfModel model)
    {
        if (model.PdfFile != null && model.PdfFile.ContentLength > 0)
        {
            // Process PDF and read content
            var pdfContent = ReadPdfContent(model.PdfFile.InputStream);
            model.PdfContent = pdfContent;
        }

        return View(model);
    }

    // Function to read PDF content
    private string ReadPdfContent(Stream pdfStream)
    {
        StringBuilder text = new StringBuilder();

        using (PdfReader reader = new PdfReader(pdfStream))
        {
            PdfDocument pdfDoc = new PdfDocument(reader);
            var strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.LocationTextExtractionStrategy();
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                var page = pdfDoc.GetPage(i);
                string pageText = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(page, strategy);
                text.Append(pageText);
            }
        }

        return text.ToString();
    }
}
