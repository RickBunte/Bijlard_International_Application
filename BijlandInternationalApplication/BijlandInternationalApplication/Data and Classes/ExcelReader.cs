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
        //public Excel.Range range;
        #endregion

        public void InitializeExcelObject(string path)
        {
            xlWorkbook = xlApp.Workbooks.Open(path);
            xlWorksheet = xlWorkbook.Worksheets.get_Item(2) as Excel.Worksheet;
        }

        public ExcelReader() {
        }
    }
}