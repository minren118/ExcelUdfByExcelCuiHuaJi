using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using IExcel = Microsoft.Office.Interop.Excel;
namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的超链接地址。Excel催化剂出品，必属精品！")]

        public static string GetCellHyperlinksAddress(
                            [ExcelArgument(Description = "带链接的单元格", AllowReference = true)] object srcRange)
        {
            ExcelReference excelReference = srcRange as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            return excelRange.Hyperlinks[1].Address;
        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的行高。Excel催化剂出品，必属精品！")]

        public static double GetCellRowHeight(
                    [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcRange)
        {
            IExcel.Range excelRange;
            if (srcRange is ExcelMissing)
            {
                IExcel.Application app = ExcelDnaUtil.Application as IExcel.Application;
                excelRange = app.ActiveCell;
            }
            ExcelReference excelReference = srcRange as ExcelReference;
            excelRange = excelReference.ToPiaRange();
            return excelRange.Height;
        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的缩进量。Excel催化剂出品，必属精品！")]

        public static double GetCellIndentLevel(
            [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcRange)
        {
            IExcel.Range excelRange;
            if (srcRange is ExcelMissing)
            {
                IExcel.Application app = ExcelDnaUtil.Application as IExcel.Application;
                excelRange = app.ActiveCell;
            }
            ExcelReference excelReference = srcRange as ExcelReference;
            excelRange = excelReference.ToPiaRange();
            return excelRange.IndentLevel;
        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的列宽。Excel催化剂出品，必属精品！")]

        public static double GetCellColumnWidth(
            [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcRange)
        {
            IExcel.Range excelRange;
            if (srcRange is ExcelMissing)
            {
                IExcel.Application app = ExcelDnaUtil.Application as IExcel.Application;
                excelRange = app.ActiveCell;
            }

            ExcelReference excelReference = srcRange as ExcelReference;
            excelRange = excelReference.ToPiaRange();
            return excelRange.EntireColumn.ColumnWidth;
        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的公式。Excel催化剂出品，必属精品！")]
        public static object GetCellFormular(
                    [ExcelArgument(Description = "带公式的单元格", AllowReference = true)] object srcRange)
        {
            ExcelReference excelReference = srcRange as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }
            else
            {
                return excelRange.Formula;
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的批注信息。Excel催化剂出品，必属精品！")]
        public static object GetCellCommentText(
            [ExcelArgument(Description = "带批注的单元格", AllowReference = true)] object srcRange)
        {
            ExcelReference excelReference = srcRange as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }
            else
            {
                if (excelRange.Comment != null)
                {
                    return excelRange.Comment.Text();
                }
                else
                {
                    return ExcelEmpty.Value;
                }
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的显示文本信息。Excel催化剂出品，必属精品！")]
        public static object GetCellText(
             [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }
            else
            {
                return excelRange.Text;
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的数字格式设置信息。Excel催化剂出品，必属精品！")]
        public static object GetCellNumberFormat(
                [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }
            else
            {
                return excelRange.NumberFormatLocal;
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的填充颜色索引。Excel催化剂出品，必属精品！")]
        public static object GetCellInteriorColor(
          [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            else
            {
                return excelRange.Interior.Color;
            }

        }

        public static object ConvertColorIndexToColor(
                            [ExcelArgument(Description = "输入颜色序号，范围为1-56", AllowReference = true)] int colorIndex)
        {
            if (colorIndex>=1 && colorIndex<=56)
            {
                return Common.ExcelApp.ActiveWorkbook.Colors[colorIndex];
            }
            else
            {
                return ExcelError.ExcelErrorNA;
            }

         


        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的文字颜色索引。Excel催化剂出品，必属精品！")]
        public static object GetCellFontColor(
                [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            else
            {
                return excelRange.Font.Color;
            }

        }




        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的文字是否粗体。Excel催化剂出品，必属精品！")]
        public static object GetCellFontBold(
        [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            else
            {
                return excelRange.Font.Bold;
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的文字是否斜体。Excel催化剂出品，必属精品！")]
        public static object GetCellFontItalic(
                        [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            else
            {
                return excelRange.Font.Italic;
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的文字字号。Excel催化剂出品，必属精品！")]
        public static object GetCellFontSize(
                [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            else
            {
                return excelRange.Font.Size;
            }

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的文字字号。Excel催化剂出品，必属精品！")]
        public static object GetCellFontName(
        [ExcelArgument(Description = "引用单元格", AllowReference = true)] object srcCell)
        {
            ExcelReference excelReference = srcCell as ExcelReference;
            IExcel.Range excelRange = excelReference.ToPiaRange();
            if (excelRange.Cells.Count > 1)
            {
                return ExcelError.ExcelErrorNA;
            }

            else
            {
                return excelRange.Font.Name;
            }

        }


        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取单元格的批注信息。Excel催化剂出品，必属精品！")]
        public static object GetRangeAddress(
                 [ExcelArgument(Description = "输入需获取地址的单元格区域，获取本身地址可省略输入", AllowReference = true)] object srcRange,
                 [ExcelArgument(Description = "是否绝对引用返回引用的行部分，默认为否")] bool isRowAbsolute,
                 [ExcelArgument(Description = "是否绝对引用返回引用的列部分，默认为否")] bool isColumnAbsolute)
        {
            IExcel.Range excelRange = null;
            if (srcRange is ExcelMissing)
            {
                excelRange = Common.ExcelApp.ActiveCell;
            }
            else
            {
                ExcelReference excelReference = srcRange as ExcelReference;
                if (excelReference != null)
                {
                    excelRange = excelReference.ToPiaRange();
                }
                else
                {
                    return ExcelError.ExcelErrorRef;
                }
            }
            return excelRange.Address[isRowAbsolute, isColumnAbsolute];

        }


        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取当前工作表名称。Excel催化剂出品，必属精品！")]
        public static object GetCurrentWorkSheetName()
        {
            return Common.ExcelApp.ActiveSheet.Name;
        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取当前工作薄名称。Excel催化剂出品，必属精品！")]
        public static object GetCurrentWorkBookName()
        {
            return Common.ExcelApp.ActiveWorkbook.Name;

        }

        [ExcelFunction(Category = "Excel对象属性", IsVolatile = true, IsMacroType = true, Description = "获取当前工作薄全路径。Excel催化剂出品，必属精品！")]
        public static object GetCurrentWorkBookFullPath()
        {
            return Common.ExcelApp.ActiveWorkbook.FullName;

        }







    }
}
