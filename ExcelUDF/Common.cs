using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
namespace ExcelCuiHuaJi
{
    class Common
    {
        public static Microsoft.Office.Interop.Excel.Application ExcelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

       

        public static void AddValueToList(object sourceRange, ref List<object> listAll)
        {
            if (sourceRange is object[])
            {
                listAll.AddRange(sourceRange as object[]);
            }
            else if (sourceRange is object[,])
            {
                var arr = sourceRange as object[,];
                for (int i = 0; i < arr.GetLength(0); i++)
                {

                    for (int j = 0; j < arr.GetLength(1); j++)
                    {
                        listAll.Add(arr[i, j]);
                    }
                }
            }

        }
        public static List<string> GetSplitStringList(object lookupValues)
        {
            List<string> listLookupValues = new List<string>();
            if (lookupValues is string)
            {
                listLookupValues.AddRange(lookupValues.ToString().Split(new char[] { ',', '，' }).Select(s=>s.Trim()));
            }
            else if (lookupValues is object[,])
            {
                object[,] arr = lookupValues as object[,];
                for (int i = 0; i < arr.GetLength(0); i++)
                {
                    for (int j = 0; j < arr.GetLength(1); j++)
                    {
                        listLookupValues.Add(arr[i, j].ToString().Trim());
                    }
                }
            }

            return listLookupValues;
        }


        public static void ChangeNumberFormat(string numberFormatString)
        {
            object caller = XlCall.Excel(XlCall.xlfCaller);
            if (caller is ExcelReference)
            {
                    ExcelAsyncUtil.QueueAsMacro(delegate
                    {
                        // Set the formatting of the function caller
                        using (new ExcelEchoOffHelper())
                        using (new ExcelSelectionHelper((ExcelReference)caller))
                        {
                            XlCall.Excel(XlCall.xlcFormatNumber, numberFormatString);
                        }
                    });
            }
        }




        public static object RunMacro(object oApp, object[] oRunArgs)
        {
           return oApp.GetType().InvokeMember("Run",
                System.Reflection.BindingFlags.Default |
                System.Reflection.BindingFlags.InvokeMethod,
                null, oApp, oRunArgs);
        }

        public static bool IsMissOrEmpty(object srcPara)
        {
            if (srcPara is ExcelMissing || srcPara is ExcelEmpty || string.IsNullOrWhiteSpace(srcPara.ToString()))
            {
                return true;

            }
            else
            {
                return false;
            }
        }

        public static double TransNumberPara(object srcPara)
        {
            double number;
            NumberStyles style = NumberStyles.Any;
            CultureInfo culture = CultureInfo.CurrentCulture;
            if (srcPara is bool && (bool)srcPara == true)
            {
                return 1;
            }
            else if (double.TryParse(srcPara.ToString(), style, culture, out number))
            {
                return number;
            }
            else
            {
                return 0;
            }
        }


        public static bool TransBoolPara(object srcPara)
        {
            double number;
            NumberStyles style = NumberStyles.Any;
            CultureInfo culture = CultureInfo.CurrentCulture;

            if (srcPara is bool && (bool)srcPara == true)
            {
                return true;
            }
            else if (double.TryParse(srcPara.ToString(), style, culture, out number))
            {
                if (number!=0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
                
            }
            else if (srcPara is string && string.IsNullOrEmpty((string)srcPara) != true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static object ReturnDataArray(object[] srcArrData, string optAlignHorL)
        {

            int resultCount = srcArrData.Count();

            if (Common.IsMissOrEmpty(optAlignHorL) || optAlignHorL.Equals("H", StringComparison.CurrentCultureIgnoreCase) == false)
            {
                optAlignHorL = "L";
            }
            else
            {
                optAlignHorL = "H";
            }
            //直接用从下标为0开始的数组也可以
            if (optAlignHorL == "L")
            {
                object[,] resultArr = new object[resultCount, 1];
                for (int i = 0; i < resultCount; i++)
                {
                    resultArr[i, 0] = srcArrData[i];
                }
                //return resultArr;
                return ArrayResizer.Resize(resultArr);
            }

            else
            {
                //int[] myLengthsArray = new int[2] {1, resultCount };
                //int[] myBoundsArray = new int[2] { 1, 1 };
                //Array resultArr = Array.CreateInstance(typeof(object), myLengthsArray, myBoundsArray);
                //for (int i = resultArr.GetLowerBound(0); i <= resultArr.GetUpperBound(0); i++)
                //    for (int j = resultArr.GetLowerBound(1); j <= resultArr.GetUpperBound(1); j++)
                //    {
                //        int[] myIndicesArray = new int[2] { i, j };
                //        resultArr.SetValue(subfolders[ilist], myIndicesArray);
                //        ilist++;
                //    }
                //横排时，直接用一维数组就可以识别到
                object[,] resultArr = new object[1, resultCount];
                for (int i = 0; i < resultCount; i++)
                {
                    resultArr[0,i] = srcArrData[i];
                }
                return ArrayResizer.Resize(resultArr);
                //return srcArrData;
                //object[,] resultArr = new object[1, resultCount];
                //for (int i = 0; i < resultCount; i++)
                //{


                //}
                //return resultArr;
            }

        }
    }
}
