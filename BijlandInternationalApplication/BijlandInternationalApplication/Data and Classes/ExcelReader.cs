using BijlandInternationalApplication.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;

namespace BijlandInternationalApplication.Data_and_Classes
{
    [Serializable]
    public class ExcelReader
    {
        #region Properties
        public Excel.Application xlApp = new Excel.Application();
        public Excel.Workbook xlWorkbook;
        public Excel.Worksheet xlWorksheet;
        private readonly string _path = @"D:\GitHub Projects\Portfolio\Bijlard_International_Application\BijlandInternationalApplication\BijlandInternationalApplication\Data and Classes\Excel\SampleData.xlsx";
        private List<Order> _orders;
        #endregion

        public void InitializeExcelObject()
        {
            xlWorkbook = xlApp.Workbooks.Open(_path);
            xlWorksheet = xlWorkbook.Worksheets.get_Item(2) as Excel.Worksheet;
            _orders = TranslateOrdersToList();
        }

        public ExcelReader() {
        }

        public List<Order> TranslateOrdersToList()
        {
            List<Order> orders = new List<Order>();
            for (int i = 2; i < xlWorksheet.UsedRange.Rows.Count; i++)
            {
                DateTime excelDate = (xlWorksheet.Cells[i, 1] as Excel.Range).Value;
                string region = (xlWorksheet.Cells[i, 2] as Excel.Range).Value;
                string rep = (xlWorksheet.Cells[i, 3] as Excel.Range).Value;
                string item = (xlWorksheet.Cells[i, 4] as Excel.Range).Value;
                int units = (int)(xlWorksheet.Cells[i, 5] as Excel.Range).Value;
                float price = (float)(xlWorksheet.Cells[i, 6] as Excel.Range).Value;
                Order order = new Order(i - 2, excelDate, region, rep, item, units, price);
                orders.Add(order);
            }
            return orders;
        }

        public List<Order> GetOrders()
        {
            return _orders;
        }

        public Order GetOrderById(int id)
        {
            return _orders.Find(order => order.GetId() == id);
        }
    }
}