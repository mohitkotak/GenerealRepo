using ClosedXML.Excel;
using GeneralWorkMVC.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GeneralWorkMVC.Controllers
{
    public class TempController : Controller
    {
        // GET: Temp
        public ActionResult Index()
        {
            return View();
        }

        // Action to upload the file and read data
        [HttpGet]
        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                using (var memoryStream = new MemoryStream())
                {
                    file.InputStream.CopyTo(memoryStream);
                    using (var workbook = new XLWorkbook(memoryStream))
                    {
                        // Read data from Sheet1
                        var sheet1 = workbook.Worksheet("Sheet1");
                        var sheet1Data = new List<User>();
                        foreach (var row in sheet1.RowsUsed().Skip(1)) // Skip header row
                        {
                            sheet1Data.Add(new User
                            {
                                Name = row.Cell(1).GetValue<string>(),
                                Age = row.Cell(2).GetValue<int>(),
                                Email = row.Cell(3).GetValue<string>(),
                            });
                        }

                        // Read data from Sheet2
                        var sheet2 = workbook.Worksheet("Sheet2");
                        var sheet2Data = new List<Product>();
                        foreach (var row in sheet2.RowsUsed().Skip(1)) // Skip header row
                        {
                            sheet2Data.Add(new Product
                            {
                                ProductName = row.Cell(1).GetValue<string>(),
                                Price = row.Cell(2).GetValue<decimal>(),
                                Quantity = row.Cell(3).GetValue<int>()
                            });
                        }

                        // Pass both lists to the view (without ViewModel)
                        ViewBag.Sheet1Data = sheet1Data;
                        ViewBag.Sheet2Data = sheet2Data;

                        // Redirect to the display page (same action or different view)
                        return RedirectToAction("DisplayExcelData");
                    }
                }
            }

            // Return to the upload page if no file is selected
            ViewBag.ErrorMessage = "Please select an Excel file.";
            return View();
        }

        // Action to display the Excel data (without ViewModel)
        public ActionResult DisplayExcelData()
        {
            return View();
        }
    }
}