using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelCuiHuaJi.CustomerExtentions
{
    class DnaExtentions
    {
        static public Excel.Range GetCaller()
        {
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            return caller.ToPiaRange();
        }

        static public ExcelReference GetExcelReference(string sheetName, int row, int column)
        {
            return new ExcelReference(row, column, row, column, sheetName);
        }

        static public ExcelReference GetExcelReference(string sheetName, string address)
        {
            Excel.Range rng = (ExcelDnaUtil.Application as Excel.Application).Range[address];
            Excel.Range firstRng = rng[1];
            Excel.Range lastRng = rng[rng.Count];

            return new ExcelReference(firstRng.Row - 1, lastRng.Row - 1, firstRng.Column - 1, lastRng.Column - 1, sheetName);
        }

    }
}