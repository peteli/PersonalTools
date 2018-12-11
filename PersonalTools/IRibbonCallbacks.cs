using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace peteli.PersonalTools
{
    interface IRibbonCallbacks
    {
        #region function
        string GetAltText(IRibbonControl control);
        string GetContent(IRibbonControl control);
        string GetDescription(IRibbonControl control);
        bool GetEnabled(IRibbonControl control);
        string GetHelperText(IRibbonControl control, int itemIndex);
        string GetHelperText(IRibbonControl control);
        //string GetHelperText(IRibbonControl control);
        //IPictureDisp GetImage(IRibbonControl control);
        Bitmap GetImage(IRibbonControl control);
        string GetImageString(IRibbonControl control);
        string GetImageMso(IRibbonControl control);
        int GetItemCount(IRibbonControl control);
        //int GetItemCount(IRibbonControl control);
        //int GetItemCount(IRibbonControl control);
        string GetItemCountString(IRibbonControl control);
        int getItemHeight(IRibbonControl control);
        string GetItemID(IRibbonControl control, int itemIndex);
        //string GetItemID(IRibbonControl control, int itemIndex);
        string GetItemID(IRibbonControl control);
        Bitmap GetItemImage(IRibbonControl control, int index);
        string GetItemLabel(IRibbonControl control, int itemIndex);
        //string GetItemLabel(IRibbonControl control, int itemIndex);
        //string GetItemLabel(IRibbonControl control, int itemIndex);
        string GetItemLabel(IRibbonControl control);
        string GetItemScreentip(IRibbonControl control, int itemIndex);
        string GetItemSupertip(IRibbonControl control, int itemIndex);
        int getItemWidth(IRibbonControl control);
        string GetKeytip(IRibbonControl control);
        string GetKeytip(IRibbonControl control, int itemIndex);
        //string GetKeytip(IRibbonControl control, int itemIndex);
        string GetLabel(IRibbonControl control);
        string GetLabel(IRibbonControl control, int itemIndex);
        //string GetLabel(IRibbonControl control, int itemIndex);
        //string GetLabel(IRibbonControl control, int itemIndex);
        bool GetPressed(IRibbonControl control);
        //bool GetPressed(IRibbonControl control);
        string GetScreentip(IRibbonControl control);
        string GetScreentip(IRibbonControl control, int itemIndex);
        //string GetScreentip(IRibbonControl control, int itemIndex);
        string GetSelectedItemID(IRibbonControl control, int itemIndex);
        int GetSelectedItemIndex(IRibbonControl control);
        string GetSelectedItemIndex(IRibbonControl control, int itemIndex);
        //string GetSelectedItemIndex(IRibbonControl control);
        bool GetShowImage(IRibbonControl control);
        //bool GetShowImage(IRibbonControl control);
        //bool GetShowImage(IRibbonControl control);
        //bool GetShowImage(IRibbonControl control);
        //bool GetShowImage(IRibbonControl control);
        //bool GetShowImage(IRibbonControl control);
        bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //bool GetShowLabel(IRibbonControl control);
        //BackstageGroupStyle GetStyle(IRibbonControl control);
        string GetSupertip(IRibbonControl control);
        string GetSupertip(IRibbonControl control, int itemIndex);
        string GetTarget(IRibbonControl control);
        string GetText(IRibbonControl control);
        string GetTitle(IRibbonControl control);
        bool GetVisible(IRibbonControl control);
        Bitmap LoadImage(string image_id);
        void OnAction(IRibbonControl control);
        //void OnAction(IRibbonControl control);
        //void OnAction(IRibbonControl control);
        void OnAction(IRibbonControl control, string itemID, int itemIndex);
        //void OnAction(IRibbonControl control, string itemID, int itemIndex);
        //void OnAction(IRibbonControl control, string itemID, int itemIndex);
        //void OnAction(IRibbonControl control, string itemID, int itemIndex);
        //void OnAction(IRibbonControl control);
        //void OnAction(IRibbonControl control);
        void OnChange(IRibbonControl control, string text);
        //void OnChange(IRibbonControl control, string text);
        void OnHide(object contextObject);
        void OnLoad(IRibbonUI ribbon);
        void OnShow(object contextObject);
        #endregion

    }
    
}
