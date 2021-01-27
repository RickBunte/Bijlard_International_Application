using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BijlandInternationalApplication.Data_and_Classes;
using BijlandInternationalApplication.Models;

namespace BijlandInternationalApplication.Controllers
{
    public class HomeController : Controller
    {
        private static ExcelReader excelReader;
        //I HAVE TO IMPROVE ON THIS!

        public ActionResult Index()
        {
            excelReader = new ExcelReader();
            excelReader.InitializeExcelObject();

            return View(excelReader.GetOrders());
        }

        public ActionResult Order(int id)
        {
            return View(excelReader.GetOrderById(id));
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        
    }
}