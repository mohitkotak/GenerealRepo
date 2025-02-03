using System.Web;

namespace GeneralWorkMVC.Models
{
    public class PdfModel
    {
        public HttpPostedFileBase PdfFile { get; set; }
        public string PdfContent { get; set; }
    }
}