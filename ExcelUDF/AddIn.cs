using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
namespace ExcelCuiHuaJi
{
   public  class AddIn : IExcelAddIn
    {
        public static bool IsFirstRun = true;
        public static Excel.Application ExcelApp;

        public void AutoOpen()
        {
            if (AddIn.IsFirstRun == true)
            {
                IntelliSenseServer.Register();
                AddIn.IsFirstRun = false;
            }

            ExcelApp = ExcelDnaUtil.Application as Excel.Application;
            ExcelApp.SheetActivate += ExcelApp_SheetActivate;
            ExcelApp.SheetSelectionChange += ExcelApp_SheetSelectionChange;
        }

        /// <summary>
        /// 添加单元格右键菜单，可以直达左上角
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        private void ExcelApp_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            var cellsbar = ExcelApp.CommandBars["cell"];
            var control = cellsbar.Controls.Cast<Microsoft.Office.Core.CommandBarControl>().FirstOrDefault(s=>s.Caption== "当前单元格跳转至窗口左上角");
            if (control==null)
            {
                control = cellsbar.Controls.Add(Type: Microsoft.Office.Core.MsoControlType.msoControlButton, Before: 1, Temporary: false);
                control.Caption = "当前单元格跳转至窗口左上角";
                control.OnAction = $"GotoTopWindow";
            }

        }

        private void ExcelApp_SheetActivate(object Sh)
        {
            Excel.Worksheet sht = Sh as Excel.Worksheet;
            Excel.Workbook wkb = sht.Parent;
            try
            {
                if (wkb.Worksheets.Cast<Excel.Worksheet>().Count(s => s.Visible == Excel.XlSheetVisibility.xlSheetVisible) > 1)//当可见工作表大于1个时才起作用
                {
                    Excel.Worksheet shtCatalog = wkb.Worksheets["工作表目录"];
                    var wkbTabsbar = ExcelApp.CommandBars["ply"];
                    wkbTabsbar.Reset();

                    var control = wkbTabsbar.Controls.Add(Type: Microsoft.Office.Core.MsoControlType.msoControlButton, Before: 1, Temporary: true);
                    control.Caption = "跳转至工作表目录";
                    control.OnAction = "GoToCataLogSht";

                }
            }
            catch (Exception)
            {
                return;
            }

        }

        public void AutoClose()
        {
        }

    }

    public  class AddinVoid:AddIn
    {
        public static  void GoToCataLogSht()
        {
            AddIn.ExcelApp.ActiveWorkbook.Worksheets["工作表目录"].Activate();
        }


        public static void GotoTopWindow(string targetAddress)
        {
            ExcelApp.EnableEvents = false;
            ExcelApp.Goto(Common.ExcelApp.ActiveCell, true);
            ExcelApp.EnableEvents = true;
        }


    }
}
