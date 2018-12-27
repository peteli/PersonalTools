using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace peteli.PersonalTools
{
    /////////////// Helper class to manage CTP ///////////////////////////
    internal static class CTPManager
    {
        static CustomTaskPane ctp;

        public static void ShowCTP()
        {
            if (ctp == null)
            {
                // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(UserControlHost), "Corporate Header");
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                ctp.DockPositionStateChange += CTP_DockPositionStateChange;
                ctp.VisibleStateChange += CTP_VisibleStateChange;
                ctp.Visible = true;
            }
            else
            {
                // Just show it again
                ctp.Visible = true;
            }
        }

        public static void DeleteCTP()
        {
            if (ctp != null)
            {
                // Could hide instead, by calling ctp.Visible = false;
                ctp.Delete();
                ctp = null;
            }
        }

        static void CTP_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            //MessageBox.Show("Visibility changed to " + CustomTaskPaneInst.Visible);
        }

        static void CTP_DockPositionStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            //((CTP_rangeexport)ctp.ContentControl).TheLabel.Text = "Moved to " + CustomTaskPaneInst.DockPosition.ToString();
        }
    }
}

//Dialog Properties / Summary Info / (xlDialogProperties, xlDialogSummaryInfo)