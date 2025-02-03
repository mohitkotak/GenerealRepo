using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GeneralWorkMVC.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult ChangeLanguage(string lang)
        {
            // Validate the language to prevent invalid inputs
            if (string.IsNullOrEmpty(lang) || !(lang == "en" || lang == "fr" || lang == "es"))
            {
                lang = "en"; // Default to English if no valid language is selected
            }

            // Set the language cookie with a 30-day expiration
            var languageCookie = new HttpCookie("Language", lang)
            {
                Expires = DateTime.Now.AddMonths(1) // Cookie will expire in one month
            };
            Response.Cookies.Add(languageCookie);

            return RedirectToAction("Index");
        }
    }
}