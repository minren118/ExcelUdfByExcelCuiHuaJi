using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "序列函数", Description = "选定区域内容重复，如选定区域内容为1,2,3，重复3次结果为123123123。Excel催化剂出品，必属精品！")]
        public static object XL重复选定区域(
            [ExcelArgument(Description = "选定要重复的区域")] object selectRange,
            [ExcelArgument(Description = "若选定的区域为多列时，是先按行还是按列排列，默认为按列，TRUE为按行")] bool IsByRows,
            [ExcelArgument(Description = "重复的次数")] int repeatTimes,
            [ExcelArgument(Description = "返回多值时是按行排列还是按列排列，输入H为按行横向，输入L为按列纵向")] string optAlignHorL
            )
        {
            List<object> listSrc = new List<object>();
            if (selectRange is object[,])
            {
                object[,] srcArr = selectRange as object[,];
                if (IsByRows)
                {
                    for (int i = 0; i < srcArr.GetLength(1); i++)
                    {
                        for (int j = 0; j < srcArr.GetLength(0); j++)//0为行，后面的arr写法是srcArr[j,i]
                        {
                            listSrc.Add(srcArr[j,i]);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < srcArr.GetLength(0); i++)
                    {
                        for (int j = 0; j < srcArr.GetLength(1); j++)
                        {
                            listSrc.Add(srcArr[i, j]);
                        }
                    }
                }
            }
            else if (selectRange is object[])
            {
                object[] srcArr = selectRange as object[];
                for (int i = 0; i < srcArr.GetLength(0); i++)
                {
                    listSrc.Add(srcArr[i]);
                }
            }
            else
            {
                return ExcelError.ExcelErrorGettingData;
            }

            List<object> result = new List<object>();
            for (int i = 0; i < repeatTimes; i++)
            {
                result.AddRange(listSrc);
            }

            return Common.ReturnDataArray(result.ToArray(),optAlignHorL);

        }




        [ExcelFunction(Category = "序列函数", IsMacroType = true, Description = "生成重复循环序列，从输入函数公式开始向下填充AAABBBCCC这样结构的数据序列。Excel催化剂出品，必属精品！")]
        public static object XL重复循环列字母(
                       [ExcelArgument(Description = "开始列字母")] string firstColChar,
                       [ExcelArgument(Description = "重复的次数")] int repeatTimes,
                       [ExcelArgument(Description = "递增或递减的步长，默认为1，如为2时，效果为AAACCCEEE")] object loopStep,
                       [ExcelArgument(Description = "结束列字母，为空时填充至当前数据区域下当前列的最后一个单元格")] object lastColChar
                                        )
        {
            List<string> listColChars = GetColChars();

            int firstIndex = listColChars.IndexOf(firstColChar);
            object[] resultIndexs;
            if (lastColChar is ExcelMissing)
            {
                resultIndexs = GetArrOfRepeat(firstIndex, repeatTimes, loopStep, lastColChar);
            }
            else
            {
                int lastIndex = listColChars.IndexOf(lastColChar.ToString());
                resultIndexs = GetArrOfRepeat(firstIndex, repeatTimes, loopStep, lastIndex);
            }
            return Common.ReturnDataArray(resultIndexs.Select(s => listColChars[Convert.ToInt32(s)]).ToArray(), "L");

        }

        private static List<string> GetColChars()
        {
            List<string> listColChars = new List<string>();
            listColChars.AddRange(GetListA());
            listColChars.AddRange(GetListAA());
            return listColChars;
        }

        [ExcelFunction(Category = "序列函数", IsMacroType = true, Description = "生成重复循环序列，从输入函数公式开始向下填充111222333这样结构的数据序列。Excel催化剂出品，必属精品！")]
        public static object XL重复循环整数(
           [ExcelArgument(Description = "开始序号")] int firstIndex,
           [ExcelArgument(Description = "重复的次数")] int repeatTimes,
           [ExcelArgument(Description = "递增或递减的步长，默认为1，如为2时，效果为111333555")] object loopStep,
           [ExcelArgument(Description = "结束序号，为空时填充至当前数据区域下当前列的最后一个单元格")] object lastIndex
           )
        {
            object[] arr = GetArrOfRepeat(firstIndex, repeatTimes, loopStep, lastIndex);
            return Common.ReturnDataArray(arr, "L");
        }

        private static object[] GetArrOfRepeat(int firstIndex, int repeatTimes, object loopStep, object lastIndex)
        {
            int arrNum;
            int step = GetStep(loopStep);

            if (lastIndex is ExcelMissing)
            {
                arrNum = GetArrNum();
            }
            else
            {
                if (firstIndex> Convert.ToInt32(lastIndex) && step>0)
                {
                    step = step * -1;
                }
                arrNum = (Math.Abs(Convert.ToInt32(lastIndex) - firstIndex) / Math.Abs(step) + 1)  * repeatTimes;//防止有负数出现，加上绝对值
            }

            object[] arr = new object[arrNum];
            int iloop = 0;

            while (iloop < arr.Length)
            {
                for (int j = 0; j < repeatTimes; j++)
                {

                    if (iloop < arr.Length)
                    {
                        arr[iloop] = firstIndex;
                        iloop++;
                    }

                }
                firstIndex = firstIndex + step;
            }

            return arr;
        }


        [ExcelFunction(Category = "序列函数", IsMacroType = true, Description = "生成间隔循环序列，从输入函数公式开始向下填充123123123这样结构的数据序列。Excel催化剂出品，必属精品！")]
        public static object XL间隔循环整数(
           [ExcelArgument(Description = "开始序号")] int firstIndex,
           [ExcelArgument(Description = "结束序号")] int lastIndex,
           [ExcelArgument(Description = "递增或递减的步长，默认为1，如为2时，效果为135135135")] object loopStep,
           [ExcelArgument(Description = "重复的次数，为空时填充至当前数据区域下当前列的最后一个单元格")] object repeatTimes

   )
        {
            object[] arr = GetArrOfInteval(firstIndex, lastIndex, loopStep, repeatTimes);
            return Common.ReturnDataArray(arr, "L");

        }


        [ExcelFunction(Category = "序列函数", IsMacroType = true, Description = "生成间隔循环序列，从输入函数公式开始向下填充ABCABCABC这样结构的数据序列。Excel催化剂出品，必属精品！")]
        public static object XL间隔循环列字母(
          [ExcelArgument(Description = "开始序号")] string firstColChar,
          [ExcelArgument(Description = "结束序号")] string lastColChar,
          [ExcelArgument(Description = "递增或递减的步长，默认为1，如为2时，效果为ACEACEACE")] object loopStep,
          [ExcelArgument(Description = "重复的次数，为空时填充至当前数据区域下当前列的最后一个单元格")] object repeatTimes
          )

        {
            List<string> listColChars = GetColChars();
            int firstIndex = listColChars.IndexOf(firstColChar);
            int lastIndex = listColChars.IndexOf(lastColChar);

            object[] resultIndexs = GetArrOfInteval(firstIndex, lastIndex, loopStep, repeatTimes);
            return Common.ReturnDataArray(resultIndexs.Select(s => listColChars[Convert.ToInt32(s)]).ToArray(), "L");
        }


        private static object[] GetArrOfInteval(int firstIndex, int lastIndex, object loopStep, object repeatTimes)
        {
            int arrNum;

            int step = GetStep(loopStep);
            if (step >0 && firstIndex > lastIndex)
            {
                step = step*-1;
            }

            if (repeatTimes is ExcelMissing)
            {
                arrNum = GetArrNum() / Math.Abs(step);
            }
            else
            {
                arrNum = (Math.Abs(lastIndex - firstIndex)/Math.Abs(step) + 1)   * Convert.ToInt32(repeatTimes);
            }

            object[] arr = new object[arrNum];
            int iloop = 0;

            while (iloop < arr.Length)
            {
                int firstLoop = firstIndex;

                while ((firstLoop >= lastIndex && step < 0) || (firstLoop <= lastIndex && step > 0))
                {
                    if (iloop < arr.Length)
                    {
                        arr[iloop] = firstLoop;
                        iloop++;
                    }
                    firstLoop = firstLoop + step;
                }
            }

            return arr;
        }

        private static int GetArrNum()
        {
            int arrNum;
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            IExcel.Range range = caller.ToPiaRange();
            int firstRowIndex = range.Row;
            IExcel.Range currentRegion = range.CurrentRegion;
            arrNum = currentRegion.Rows.Count + currentRegion.Row - 1 - firstRowIndex + 1;
            return arrNum;
        }

        private static int GetStep(object loopStep)
        {
            int step;
            if (loopStep is ExcelMissing)
            {
                step = 1;
            }
            else
            {
                step = Convert.ToInt32(loopStep);
            }

            return step;
        }

        private static List<string> GetListAA()
        {
            List<string> listAA = new List<string>();
            for (int i = 0; i < 26; i++)
            {
                for (int j = 0; j < 26; j++)
                {
                    int index1 = 65 + i;
                    int index2 = 65 + j;
                    listAA.Add(new string(new char[] { (char)index1, (char)index2 }));
                }
            }

            return listAA;
        }

        private static string[] GetListA()
        {
            return Enumerable.Range(65, 26).Select(s => (char)s).Select(s => s.ToString()).ToArray();
        }

    }

}
