using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using System.Diagnostics;
using System.Drawing;
using peteli.PersonalTools.guid;

namespace peteli.PersonalTools
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        #region functions
        public override string GetCustomUI(string RibbonID)
        {
            //MessageBox.Show((string)myResources.Resources.Ribbon1);
            return (string)Properties.Resources.customUI;
            /*          return @"
                  <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                  <ribbon>
                    <tabs>
                      <tab id='tab1' label='My Tab'>
                        <group id='group1' label='My Group'>
                          <button id='button1' label='My Button' onAction='OnButtonPressed'/>
                        </group >
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
            */
        }

        #endregion
        #region properties
        public static IRibbonUI AppRibbon { get; protected set; }
        #endregion

        #region callbacks
        #endregion
        #region OnButtonPressed
        public static event EventHandler ButtonPressed;
        // ribbon callback
        public void OnButtonPressed(IRibbonControl control)
        {
         
        }
        // raise event
        protected virtual void OnButtonPressedEvent(EventArgs e)
        {
            EventHandler handler = ButtonPressed;
            handler?.Invoke(this, e);
        }

        public string GetAltText(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetContent(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetDescription(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public bool GetEnabled(IRibbonControl control)
        {
            bool result = true;

            switch (control.Id)
            {
                default:
                    result = true;
                    break;
            }

            return result;
        }

        public string GetHelperText(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetHelperText(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetImageString(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetImageMso(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public int GetItemCount(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemCountString(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public int getItemHeight(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemID(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemID(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public Bitmap GetItemImage(IRibbonControl control, int index)
        {
            throw new NotImplementedException();
        }

        public string GetItemLabel(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemLabel(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemScreentip(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemSupertip(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public int getItemWidth(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetKeytip(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetKeytip(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetLabel(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetLabel(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public bool GetPressed(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetScreentip(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetScreentip(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetSelectedItemID(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public int GetSelectedItemIndex(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetSelectedItemIndex(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public bool GetShowImage(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public bool GetShowLabel(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetSupertip(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetSupertip(IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetTarget(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetText(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetTitle(IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public bool GetVisible(IRibbonControl control)
        {
            bool result = true;

            switch (control.Id)
            {
                default:
                    result = true;
                    break;
            }

            return result;
        }

        public void OnAction(IRibbonControl control)
        {
            Debug.WriteLine(control.Id);

            switch (control.Id)
            {
                case "btnFormatHeaderFooter000":
                    CTPManager.ShowCTP();
                    break;
                case "btnGUID":
                    // do something
                    Debug.WriteLine("case catch {0}", control.Id);
                    XlGUID.Create();
                    break;
                default:
                    Debug.WriteLine("No action is assigned to {0}", control.Id.ToString());
                    break;
            }

        }

        public void OnAction(IRibbonControl control, string itemID, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public void OnChange(IRibbonControl control, string text)
        {
            throw new NotImplementedException();
        }

        public void OnHide(object contextObject)
        {
            throw new NotImplementedException();
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            AppRibbon = ribbon;
            OnLoadEvent(new LoadEventArgs { Ribbon = AppRibbon });
        }

        public void OnShow(object contextObject)
        {
            throw new NotImplementedException();
        }
        #endregion




        #region events
        internal static event EventHandler<LoadEventArgs> Load;
        protected virtual void OnLoadEvent(LoadEventArgs e)
        {
            EventHandler<LoadEventArgs> handler = Load;
            handler?.Invoke(this, e);
        }
        #endregion



        #region event arguments
        public class ButtonPressedEventArgs : EventArgs
        {
            public IRibbonControl ControlId { get; set; }
        }
        public class LoadEventArgs : EventArgs
        {
            public IRibbonUI Ribbon { get; set; }
        }
        #endregion

    }
}