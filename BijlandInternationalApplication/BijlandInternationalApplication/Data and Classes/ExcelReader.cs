using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;

namespace BijlandInternationalApplication.Data_and_Classes
{
    public class ExcelReader
    {
        #region Properties
        public Excel.Application xlApp = new Excel.Application();
        public Excel.Workbook xlWorkbook;
        public Excel.Worksheet xlWorksheet;
        public Excel.Range range;
        #endregion

        public void InitializeExcelObject()
        {
            xlWorkbook = xlApp.Workbooks.Open(@"~\Excel\SampleData.xlsx");
            xlWorksheet = xlWorkbook.Worksheets.get_Item(1) as Excel.Worksheet;
        }
    }
}