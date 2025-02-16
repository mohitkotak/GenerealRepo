using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GeneralWorkMVC.Models;
using System;
using System.Collections.Generic;
using System.IO;
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
            try
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
                            var image1Base64 = ExtractImagesFromExcel(file.InputStream, "Sheet1");
                            sheet1Data.Add(new Sheet1Model
                            {
                                Id = row.Cell(1).GetValue<int>(),
                                Name = row.Cell(2).GetValue<string>(),
                                Amount = row.Cell(3).GetValue<double>(),
                                Image = image1Base64
                            });
                        }
                    
                        // Reading Sheet2
                        var sheet2 = workbook.Worksheet("Sheet2");
                        foreach (var row in sheet2.RowsUsed().Skip(1)) // Skip header row
                        {
                            var image2Base64 = ExtractImagesFromExcel(file.InputStream, "Sheet2");
                            sheet2Data.Add(new Sheet2Model
                            {
                                Code = row.Cell(1).GetValue<int>(),
                                Description = row.Cell(2).GetValue<string>(),
                                Price = row.Cell(3).GetValue<decimal>(),
                                Picture = image2Base64
                            });
                        }
                    }

                    // Returning the data as JSON to the client
                    var data = new { Sheet1Data = sheet1Data, Sheet2Data = sheet2Data };
                    return Json(data, JsonRequestBehavior.AllowGet);
                }

                return Json(new { success = false, message = "No file uploaded" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }


        private string ExtractImagesFromExcel(Stream inputStream, string sheetName)
        {
            // Load the Excel workbook
            using (var workbook = new XLWorkbook(inputStream))
            {
                var worksheet = workbook.Worksheet(sheetName);

                // Check if there are any images in the sheet
                var images = worksheet.Pictures;

                // Assuming you're extracting the first image from the sheet (if it exists)
                foreach (var image in images)
                {
                    // Convert the image to a byte array
                    using (var memoryStream = new MemoryStream())
                    {
                        image.ImageStream.CopyTo(memoryStream); // Copy image stream to memory stream
                        byte[] imageBytes = memoryStream.ToArray();

                        // Convert byte array to base64 string
                        return Convert.ToBase64String(imageBytes);
                    }
                }
            }

            // Return null if no image is found
            return null;
        }




    }
}