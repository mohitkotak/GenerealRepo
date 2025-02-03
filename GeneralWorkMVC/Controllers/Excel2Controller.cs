using ClosedXML.Excel;
using GeneralWorkMVC.Models;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GeneralWorkMVC.Controllers
{
    public class Excel2Controller : Controller
    {
        // GET: Excel2
        public ActionResult Index()
        {
            return View();
        }

        // Action to render the view with the upload form
        [HttpGet]
        public ActionResult Upload()
        {
            return View();
        }

        // Action to handle the uploaded file and return sheet data as JSON
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                var sheet1Data = new List<Sheet1Model>();
                var sheet2Data = new List<Sheet2Model>();

                using (var workbook = new XLWorkbook(file.InputStream))
                {
                    // Reading Sheet1
                    var sheet1 = workbook.Worksheet("Sheet1");
                    foreach (var row in sheet1.RowsUsed().Skip(1)) // Skip header row
                    {
                        sheet1Data.Add(new Sheet1Model
                        {
                            Id = row.Cell(1).GetValue<int>(),
                            Name = row.Cell(2).GetValue<string>(),
                            Amount = row.Cell(3).GetValue<double>()
                        });
                    }

                    // Reading Sheet2
                    var sheet2 = workbook.Worksheet("Sheet2");
                    foreach (var row in sheet2.RowsUsed().Skip(1)) // Skip header row
                    {
                        sheet2Data.Add(new Sheet2Model
                        {
                            Code = row.Cell(1).GetValue<int>(),
                            Description = row.Cell(2).GetValue<string>(),
                            Price = row.Cell(3).GetValue<decimal>()
                        });
                    }
                }

                // Returning the data as JSON to the client
                var data = new { Sheet1Data = sheet1Data, Sheet2Data = sheet2Data };
                return Json(data, JsonRequestBehavior.AllowGet);
            }

            return Json(new { success = false, message = "No file uploaded" });
        }

    }
}