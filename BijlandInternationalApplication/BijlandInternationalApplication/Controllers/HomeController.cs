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
        public List<Order> orders;
        private ExcelReader excelReader = new ExcelReader();
        //private readonly string _path = @"D:\GitHub Projects\Portfolio\Bijlard_International_Application\BijlandInternationalApplication\BijlandInternationalApplication\Data and Classes\Excel\SampleData.xlsx";
        private readonly string _path = @"\\Data and Classes\Excel\SampledData.xlsx";

        public ActionResult Index()
        {
            excelReader.InitializeExcelObject(_path);
            orders = TranslateOrdersToList();

            return View(orders);
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

        public List<Order> TranslateOrdersToList()
        {
            List<Order> orders = new List<Order>();
            for(int i = 2; i < excelReader.xlWorksheet.UsedRange.Rows.Count; i++)
            {
                DateTime excelDate = (excelReader.xlWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value;
                string region = (excelReader.xlWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value;
                string rep = (excelReader.xlWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value;
                string item = (excelReader.xlWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value;
                int units = (int)(excelReader.xlWorksheet.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value;
                float price = (float)(excelReader.xlWorksheet.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value;
                Order order = new Order(excelDate, region, rep, item, units, price);
                orders.Add(order);
            }
            return orders;
        }
    }
}