using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace peteli.PersonalTools.guid
{
    /// <summary>
    /// puts a GUID in active cell
    /// </summary>
    public static class XlGUID
    {
        #region functions
        public static void Create()
        {
            Range currentCell = _XlApp.ActiveCell;

            try
            {
                currentCell.Value = Guid.NewGuid().ToString();
            }
            catch (System.NullReferenceException e)
            {
                MessageBox.Show(@"Doesn't work here." , sorryFace);
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show(@"Doesn't work here, sorry 'bout that.", sorryFace);
            }
        }
        #endregion

        #region properties
        static ExcelApp _XlApp = (ExcelApp)ExcelDnaUtil.Application;
        const string sorryFace = @"¯\_(ツ)_/¯";
        #endregion
    }
}
