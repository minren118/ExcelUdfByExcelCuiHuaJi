using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCuiHuaJi
{
    public static class ExcelReferenceExtentions
    {
        static public Excel.Range ToPiaRange(this ExcelReference excelReference)
        {
            try
            {
                Excel.Application app = ExcelDnaUtil.Application as Excel.Application;
                string address = XlCall.Excel(XlCall.xlfReftext, excelReference.InnerReferences[0], true).ToString();
                return app.Range[address];
            }
            catch (XlCallException ex)
            {
                throw new Exception("调用方法的特性ExcelFunction的IsMacroType可能为false导致", ex);
            }


        }

        static public ExcelReference Resize(this ExcelReference excelReference, int rows, int columns)
        {
            rows--;
            columns--;
            return new ExcelReference(excelReference.RowFirst, excelReference.RowFirst + rows, excelReference.ColumnFirst, excelReference.ColumnFirst + columns, excelReference.SheetId);
        }

        static public ExcelReference Resize(this ExcelReference excelReference, ExcelReference targetReference)
        {
            return excelReference.Resize(targetReference.GetRows(), targetReference.GetColumns());
        }

        static public int GetRows(this ExcelReference excelReference)
        {
            return excelReference.RowLast - excelReference.RowFirst + 1;
        }

        static public int GetColumns(this ExcelReference excelReference)
        {
            return excelReference.ColumnLast - excelReference.ColumnFirst + 1;
        }

    }
}
