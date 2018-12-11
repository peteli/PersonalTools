using System.Diagnostics;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace peteli.PersonalTools
{
    public class MyAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // get handle of excel application
            Debug.WriteLine("XLL AutoOpen runs");
            
        }

        public void AutoClose()
        {
            // put code here
            Debug.WriteLine("XLL closes");
        }


    }
}
