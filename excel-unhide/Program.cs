using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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

                foreach (Excel.Worksheet worksheet in app.Worksheets)
                {
                    worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                }
            }

            catch (Exception e) { Console.WriteLine(String.Format("Error unhiding sheets. Exception details: {0}", e.Message)); }
        }
    }
}
