using ExcelDna.Integration;
using Microsoft.International.Converters.PinYinConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.IO;


namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "辅助函数", Description = "传入参数为多多列时，需要使用此函数来引用多列。Excel催化剂出品，必属精品！")]
        public static object FZGetMultiColRange(
    [ExcelArgument(Description = "分组列区域1，仅能选取一列")] object[,] groupRange1,
    [ExcelArgument(Description = "分组列区域2，仅能选取一列")] object[,] groupRange2,
    [ExcelArgument(Description = "分组列区域3，仅能选取一列")] object[,] groupRange3,
    [ExcelArgument(Description = "分组列区域4，仅能选取一列")] object[,] groupRange4
    )
        {
            //当不止选了一列时，返回错误
            if (groupRange1.GetLength(1) != 1 || groupRange2.GetLength(1) != 1 || groupRange3.GetLength(1) != 1 || groupRange4.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0Length1 = groupRange1.GetLength(0);
            int arrDim0Length2 = groupRange2.GetLength(0);
            int arrDim0Length3 = groupRange3.GetLength(0);
            int arrDim0Length4 = groupRange4.GetLength(0);

            //当行的数量不同时，也不能计算，返回出错
            if ((arrDim0Length1 != arrDim0Length2 && arrDim0Length2 != 1) || (arrDim0Length1 != arrDim0Length3 && arrDim0Length3 != 1) || (arrDim0Length1 != arrDim0Length4 && arrDim0Length4 != 1))
            {
                throw new ArgumentException("参数出错");
            }

            List<string> result = new List<string>();

            for (int i = 0; i < arrDim0Length1; i++)
            {
                StringBuilder sb = new StringBuilder();

                sb.Append(groupRange1[i, 0].ToString() + "|");

                if (arrDim0Length2 != 1)
                {
                    sb.Append(groupRange2[i, 0].ToString() + "|");
                }
                if (arrDim0Length3 != 1)
                {
                    sb.Append(groupRange3[i, 0].ToString() + "|");
                }
                if (arrDim0Length4 != 1)
                {
                    sb.Append(groupRange4[i, 0].ToString() + "|");
                }
                result.Add(sb.ToString().TrimEnd('|'));

            }
            object[,] arrResult = new object[result.Count, 1];
            for (int i = 0; i < result.Count; i++)
            {
                arrResult[i, 0] = result[i];
            }

            return arrResult;
        }

    }
}
