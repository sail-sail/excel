using System;

using System.Threading.Tasks;

namespace excel
{
    public class Main
    {
        public async Task<object> printExcel(dynamic input)
        {
            string excelPath = (string) input.excelPath;
            bool visible = (bool)input.visible;
            bool preview = (bool)input.preview;
            if (preview == true) {
                visible = true;
            }
            string activePrinter = (string)input.activePrinter;
            Microsoft.Office.Interop.Excel.Application xApp = new Microsoft.Office.Interop.Excel.Application();
            xApp.Visible = visible;
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
                xSheet.PrintOut(1, 1, 1, preview, activePrinter);
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
