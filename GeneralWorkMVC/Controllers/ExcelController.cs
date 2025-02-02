using ClosedXML.Excel;
using AutoMapper;
using System.Collections.Generic;
using System.IO;
using System.Web.Mvc;
using GeneralWorkMVC.Models;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Spreadsheet;
using GeneralWorkMVC.ViewModel;

namespace GeneralWorkMVC.Controllers
{
    public class ExcelController : Controller
    {
        private readonly IMapper _mapper;

        /// <summary>
        /// Default Excel View
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Upload Excel
        /// </summary>
        /// <returns></returns>
        public ActionResult UploadExcel()
        {
            return View();
        }

        // Action to handle file upload and read data from Excel
        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                // Read data from the uploaded Excel file
                var usersData = new List<User>();
                var products = new List<Product>();

                // Load the uploaded file into memory
                using (var memoryStream = new MemoryStream())
                {
                    file.InputStream.CopyTo(memoryStream);
                    using (var workbook = new XLWorkbook(memoryStream))
                    {
                        // Read data from Sheet1 (Persons)
                        var sheet1 = workbook.Worksheet("Sheet1");
                        var personRows = sheet1.RowsUsed();
                        foreach (var row in personRows.Skip(1)) // Skipping header row
                        {
                            var users = new User
                            {
                                Name = row.Cell(1).GetValue<string>(),
                                Age = row.Cell(2).GetValue<int>(),
                                Email = row.Cell(2).GetValue<string>()
                            };
                            usersData.Add(users);
                        }

                        // Read data from Sheet2 (Products)
                        var sheet2 = workbook.Worksheet("Sheet2");
                        var productRows = sheet2.RowsUsed();
                        foreach (var row in productRows.Skip(1)) // Skipping header row
                        {
                            var product = new Product
                            {
                                ProductName = row.Cell(1).GetValue<string>(),
                                Price = row.Cell(2).GetValue<decimal>(),
                                Quantity = row.Cell(3).GetValue<int>()
                            };
                            products.Add(product);
                        }
                    }
                }

                // Create a model with data from both sheets
                var model = new ExcelViewModel
                {
                    Users = usersData,
                    Products = products
                };

                // Pass the model to the view
                return View("ShowExcelData", model);
            }

            // If no file is uploaded, return to the upload form
            ViewBag.Error = "Please upload a valid Excel file.";
            return View();
        }

        // Action to display the uploaded Excel data
        public ActionResult ShowExcelData(ExcelViewModel model)
        {
            return View(model);
        }


        /// <summary>
        /// Download Excel in multiple sheet
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadExcel()
        {
            // Sample data for Sheet1
            var dataSheet1 = new List<dynamic>
        {
            new { Name = "Raj Doshi", Age = 30, Email = "raj@example.com" },
            new { Name = "Mahesh Patel", Age = 25, Email = "mahesh@example.com" },
            new { Name = "Pradip Vasani", Age = 35, Email = "pradip@example.com" }
        };

            // Sample data for Sheet2
            var dataSheet2 = new List<dynamic>
        {
            new { Product = "Laptop", Price = 1200, Quantity = 5 },
            new { Product = "Smartphone", Price = 700, Quantity = 10 },
            new { Product = "Tablet", Price = 300, Quantity = 15 }
        };

            // Create a new workbook
            using (var workbook = new XLWorkbook())
            {
                // Add Sheet1
                var sheet1 = workbook.Worksheets.Add("Sheet1");

                // Set the header row for Sheet1
                sheet1.Cell(1, 1).Value = "Name";
                sheet1.Cell(1, 2).Value = "Age";
                sheet1.Cell(1, 3).Value = "Email";

                // Fill Sheet1 with data
                for (int i = 0; i < dataSheet1.Count; i++)
                {
                    sheet1.Cell(i + 2, 1).Value = dataSheet1[i].Name;
                    sheet1.Cell(i + 2, 2).Value = dataSheet1[i].Age;
                    sheet1.Cell(i + 2, 3).Value = dataSheet1[i].Email;
                }

                // Add Sheet2
                var sheet2 = workbook.Worksheets.Add("Sheet2");

                // Set the header row for Sheet2
                sheet2.Cell(1, 1).Value = "Product";
                sheet2.Cell(1, 2).Value = "Price";
                sheet2.Cell(1, 3).Value = "Quantity";

                // Fill Sheet2 with data
                for (int i = 0; i < dataSheet2.Count; i++)
                {
                    sheet2.Cell(i + 2, 1).Value = dataSheet2[i].Product;
                    sheet2.Cell(i + 2, 2).Value = dataSheet2[i].Price;
                    sheet2.Cell(i + 2, 3).Value = dataSheet2[i].Quantity;
                }

                // Save the workbook to a MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    // Return the file to the client for download
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelSampleData.xlsx");
                }
            }
        }

        /// <summary>
        /// Download color Excel in multiple sheet
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadColorExcel()
        {
            // Sample data for Sheet1
            var dataSheet1 = new List<dynamic>
        {
            new { Name = "Raj Doshi", Age = 30, Email = "raj@example.com" },
            new { Name = "Mahesh Patel", Age = 25, Email = "mahesh@example.com" },
            new { Name = "Pradip Vasani", Age = 35, Email = "pradip@example.com" }
        };

            // Sample data for Sheet2
            var dataSheet2 = new List<dynamic>
        {
            new { Product = "Laptop", Price = 1200, Quantity = 5 },
            new { Product = "Smartphone", Price = 700, Quantity = 10 },
            new { Product = "Tablet", Price = 300, Quantity = 15 }
        };

            // Create a new workbook
            using (var workbook = new XLWorkbook())
            {
                // Add Sheet1
                var sheet1 = workbook.Worksheets.Add("Sheet1");

                // Set the header row for Sheet1
                var headerSheet1 = sheet1.Range("A1:C1");
                headerSheet1.Cell(1, 1).Value = "Name";
                headerSheet1.Cell(1, 2).Value = "Age";
                headerSheet1.Cell(1, 3).Value = "Email";

                // Format header for Sheet1: Background color, font style, and borders
                headerSheet1.Style.Fill.BackgroundColor = XLColor.CornflowerBlue; // Background color
                headerSheet1.Style.Font.Bold = true; // Bold font
                headerSheet1.Style.Font.FontColor = XLColor.White; // White font color
                headerSheet1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Center alignment
                headerSheet1.Style.Border.TopBorder = XLBorderStyleValues.Thin; // Top border
                headerSheet1.Style.Border.BottomBorder = XLBorderStyleValues.Thin; // Bottom border
                headerSheet1.Style.Border.LeftBorder = XLBorderStyleValues.Thin; // Left border
                headerSheet1.Style.Border.RightBorder = XLBorderStyleValues.Thin; // Right border

                // Fill Sheet1 with data
                for (int i = 0; i < dataSheet1.Count; i++)
                {
                    sheet1.Cell(i + 2, 1).Value = dataSheet1[i].Name;
                    sheet1.Cell(i + 2, 2).Value = dataSheet1[i].Age;
                    sheet1.Cell(i + 2, 3).Value = dataSheet1[i].Email;
                }

                // Add Sheet2
                var sheet2 = workbook.Worksheets.Add("Sheet2");

                // Set the header row for Sheet2
                var headerSheet2 = sheet2.Range("A1:C1");
                headerSheet2.Cell(1, 1).Value = "Product";
                headerSheet2.Cell(1, 2).Value = "Price";
                headerSheet2.Cell(1, 3).Value = "Quantity";

                // Format header for Sheet2: Background color, font style, and borders
                headerSheet2.Style.Fill.BackgroundColor = XLColor.IndianRed; // Background color
                headerSheet2.Style.Font.Bold = true; // Bold font
                headerSheet2.Style.Font.FontColor = XLColor.White; // White font color
                headerSheet2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Center alignment
                headerSheet2.Style.Border.TopBorder = XLBorderStyleValues.Thin; // Top border
                headerSheet2.Style.Border.BottomBorder = XLBorderStyleValues.Thin; // Bottom border
                headerSheet2.Style.Border.LeftBorder = XLBorderStyleValues.Thin; // Left border
                headerSheet2.Style.Border.RightBorder = XLBorderStyleValues.Thin; // Right border

                // Fill Sheet2 with data
                for (int i = 0; i < dataSheet2.Count; i++)
                {
                    sheet2.Cell(i + 2, 1).Value = dataSheet2[i].Product;
                    sheet2.Cell(i + 2, 2).Value = dataSheet2[i].Price;
                    sheet2.Cell(i + 2, 3).Value = dataSheet2[i].Quantity;
                }

                // Save the workbook to a MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    // Return the file to the client for download
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelSampleData.xlsx");
                }
            }
        }
    }
}
