using System.Diagnostics;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Text;
using System;

namespace peteli.PersonalTools
{
    public class MyAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // get handle of excel application
            Debug.WriteLine("XLL AutoOpen runs");
            InitAppsHooks();
            
        }

        private void InitAppsHooks()
        {
            Application _XlApp = (Application)ExcelDnaUtil.Application;
            _XlApp.WorkbookActivate += DeleteCTP;
            _XlApp.WorkbookDeactivate += DeleteCTP;

        }

        private void DeleteCTP(Workbook Wb)
        {
            CTPManager.DeleteCTP();
        }

        public void AutoClose()
        {
            // put code here
            Debug.WriteLine("XLL closes");
        }

    }
}
