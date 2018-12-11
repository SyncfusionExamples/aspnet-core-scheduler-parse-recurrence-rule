using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;

namespace ScheduleSample.Controllers
{
    public class HomeController : Controller
    {   
        public ActionResult Index()
        {  
            return View();
        }
        
        public JsonResult getDates(string dates)
        {
            var recurrenceRule = JsonConvert.DeserializeObject<string>(dates);
            var dateCollection = RecurrenceHelper.GetRecurrenceDateTimeCollection(recurrenceRule, DateTime.Now);            
            return Json(dateCollection, JsonRequestBehavior.AllowGet);
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
    }
}