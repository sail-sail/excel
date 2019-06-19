using System;

using System.Threading.Tasks;

namespace excel
{
    public class Main
    {
        public async Task<object> printExcel(dynamic excelPath)
        {
            Microsoft.Office.Interop.Excel.Application xApp = new Microsoft.Office.Interop.Excel.Application();
            xApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks._Open(excelPath);
            Microsoft.Office.Interop.Excel.Worksheet xSheet = null;
            try
            {
                xApp.Run("Workbook_Open");
            }
            catch (Exception)
            {

            }
            try
            {
                xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.ActiveSheet;
                xSheet.PrintOut(1, 1, 1, false);
            }
            catch (Exception e)
            {

            }
            finally
            {
                xBook.Close(false);
                xApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xApp);
                xSheet = null;
                xBook = null;
                xApp = null;
                GC.Collect();
            }
            return "";
        }
    }
}
