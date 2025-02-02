using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace GeneralWorkMVC
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start(object sender, EventArgs e)
        {
            HttpContext context = HttpContext.Current;
            string lang = null;
            if (context != null && context.Session != null)
            {
                // Check if the language is set in session (can also check for cookies here)
                if (Session["Culture"] != null)
                {
                    lang = Session["Culture"].ToString();
                }
            }
            else
            {
                // Set default language (e.g., English) if not set
                lang = "en";
            }

            // Set the culture and UI culture based on the session or default
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(lang);
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(lang);

            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }
    }
}
