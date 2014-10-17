using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel_unhide
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Excel.Application app = (Excel.Application) Marshal.GetActiveObject("Excel.Application");
                foreach (Excel.Worksheet worksheet in app.Worksheets) { worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible; }
            }

            catch (Exception e) { Console.WriteLine("Error unhiding sheets. Exception details: {0}", e.Message); Console.ReadKey(); }
        }
    }
}
