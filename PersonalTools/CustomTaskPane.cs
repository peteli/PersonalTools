using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Collections.Generic;

namespace peteli.PersonalTools
{
    // TODO CTPManager to organize more than one ctp
    
    internal static class CustomTaskPaneManager
    {
        internal static Dictionary<string, CustomTaskPane> CustomTaskPanes = new Dictionary<string, CustomTaskPane>();

        internal static void Hide(CustomTaskPane customTaskPane)
        {
            string uniqueControlName = customTaskPane.GetHashCode().ToString();
            if(CustomTaskPanes.ContainsKey(uniqueControlName))
            {
                CustomTaskPanes[uniqueControlName].Delete();
                CustomTaskPanes.Remove(uniqueControlName);
            }
        }

        internal static CustomTaskPane Show(System.Type userControl,string title,MsoCTPDockPosition msoCTPDockPosition=MsoCTPDockPosition.msoCTPDockPositionLeft)
        {
            //throw new NotImplementedException();
            // create new custom task pane with ExcelDna.Integration.CustomUI.CustomTaskPaneFactory
            CustomTaskPane customTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(userControl, "newItem");
            customTaskPane.DockPosition = msoCTPDockPosition;
            customTaskPane.Visible = true;
            customTaskPane.VisibleStateChange += CustomTaskPane_VisibleStateChange;
            string uniqueControlName = customTaskPane.GetHashCode().ToString();
            CustomTaskPanes.Add(uniqueControlName, customTaskPane);
            return customTaskPane;
        }

        private static void CustomTaskPane_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            if (CustomTaskPaneInst.Visible == false)
            {
                Hide(CustomTaskPaneInst);
            }
        }
    }


    /// <summary>
    /// helper class to manage CTP - Custom Task Pane 
    /// </summary>
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
