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


        [ExcelFunction(Category = "查找引用增强", Description = "类似INDEX+MATCH套路，查找值在交叉表结构的数据区域中对应的值。Excel催化剂出品，必属精品！")]
        public static object CZYY查找引用INDEX(
                 [ExcelArgument(Description = "查找值的列区域，当有多列作为查找值的列时，需使用辅助函数【FZGetMultiColRange】输入")] object[,] lookupValueToRow,
                 [ExcelArgument(Description = "引用匹配列区域，当有多列作为引用匹配列列时，需使用辅助函数【FZGetMultiColRange】输入")] object[,] lookupValueToCol,
                 [ExcelArgument(Description = "引用匹配列所在数据区域内的返回列")] object[,] referenceRange
               )

        {
            ValidParas(lookupValueToCol, lookupValueToRow);

            List<(int Index, object lookupValueToRow, object lookupValueToCol)> lookupValueList = new List<(int Index, object lookupValueToRow, object lookupValueToCol)>();

            for (int i = 0; i < lookupValueToRow.GetLength(0); i++)
            {
                lookupValueList.Add((i, lookupValueToRow[i, 0], lookupValueToCol[i, 0]));
            }

            List<object> listReferenceRow = new List<object>();
            for (int i = 0; i < referenceRange.GetLength(0); i++)
            {
                listReferenceRow.Add(referenceRange[i, 0]);
            }

            List<object> listReferenceCol = new List<object>();
            for (int j = 0; j < referenceRange.GetLength(1); j++)
            {
                listReferenceCol.Add(referenceRange[0, j]);
            }

            var grpLookupValues = lookupValueList.GroupBy(s => (s.lookupValueToRow, s.lookupValueToCol));
            Dictionary<int, object> result = new Dictionary<int, object>();


            foreach (var grpLookupValue in grpLookupValues)
            {
                object referenceValue = null;
                object lookupValueRow = grpLookupValue.Key.lookupValueToRow;
                int rowIndex = 0;
                if (lookupValueRow is string)
                {
                    rowIndex = listReferenceRow.Select(s => s.ToString()).ToList().IndexOf(lookupValueRow.ToString());
                    if (rowIndex > 0)
                    {
                        referenceValue = GetReferenceValue(referenceRange, listReferenceCol, grpLookupValue, rowIndex);
                    }
                }
                else if (lookupValueRow is double)
                {
                    rowIndex = listReferenceRow.ToList().IndexOf(lookupValueRow);
                    if (rowIndex > 0)
                    {
                        referenceValue = GetReferenceValue(referenceRange, listReferenceCol, grpLookupValue, rowIndex);
                    }
                }

                foreach (var item in grpLookupValue)
                {
                    if (referenceValue != null)
                    {
                        result.Add(item.Index, referenceValue);
                    }
                    else
                    {
                        result.Add(item.Index, "");
                    }
                }
            }

            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        private static object GetReferenceValue(object[,] referenceRange, List<object> listReferenceCol, IGrouping<(object lookupValueToRow, object lookupValueToCol), (int Index, object lookupValueToRow, object lookupValueToCol)> grpLookupValue, int rowIndex)
        {
            int colIndex = 0;
            object lookupValueCol = grpLookupValue.Key.lookupValueToCol;
            if (lookupValueCol is string)
            {

                colIndex = listReferenceCol.Select(s => s.ToString()).ToList().IndexOf(lookupValueCol.ToString());

            }
            else if (lookupValueCol is double)
            {
                colIndex = listReferenceCol.ToList().IndexOf(lookupValueCol);
            }

            if (colIndex > 0)
            {
                return referenceRange[rowIndex, colIndex];
            }
            else
            {
                return null;
            }

        }


        [ExcelFunction(Category = "查找引用增强", Description = "类似Vlookup的用法，查找某列的值在相对引用区域的对应的返回值。Excel催化剂出品，必属精品！")]
        public static object CZYY反向模糊查找引用LOOKUP(
                 [ExcelArgument(Description = "查找值的列区域，当有多列作为查找值的列时，需使用辅助函数【FZGetMultiColRange】输入")] object[,] lookupValueRange,
                 [ExcelArgument(Description = "引用匹配列区域，当有多列作为引用匹配列列时，需使用辅助函数【FZGetMultiColRange】输入")] object[,] referenceRange,
                 [ExcelArgument(Description = "是否要遍历引用列所有内容返回所有结果，默认只返回首次符合条件的结果")] bool isLookupAll,
                 [ExcelArgument(Description = "当遍历引用列所有内容返回所有结果时，多个结果间的分隔符")] object splitString
               )
        {
            string splitStr;
            if (splitString is ExcelMissing)
            {
                splitStr = ",";
            }
            else
            {
                splitStr = splitString.ToString();
            }

            List<(int Index, string LookupValue)> lookupValueList = new List<(int Index, string LookupValue)>();
            for (int i = 0; i < lookupValueRange.GetLength(0); i++)
            {
                var lookupValue = lookupValueRange[i, 0];
                if (lookupValue == ExcelEmpty.Value)
                {
                    lookupValueList.Add((i, ""));
                }
                else
                {
                    lookupValueList.Add((i, lookupValueRange[i, 0].ToString()));
                }
            }

            var grpLookupValues = lookupValueList.GroupBy(s => s.LookupValue);
            Dictionary<int, string> result = new Dictionary<int, string>();
            foreach (var grpLookupValue in grpLookupValues)
            {
                string referenceValue = GetReferenceValue(grpLookupValue.Key, referenceRange, isLookupAll, splitStr);

                foreach (var item in grpLookupValue)
                {
                    result.Add(item.Index, referenceValue);
                }
            }
            if (result.Count>1)
            {
                return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
            }
            else
            {
                return result[0];
            }
            
        }

        private static string GetReferenceValue(string lookupValue, object[,] referenceRange, bool isLookupAll, string splitStr)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < referenceRange.GetLength(0); i++)
            {
                string referenceValue = referenceRange[i, 0].ToString();
                if (lookupValue.Contains(referenceValue))
                {
                    result.Add(referenceValue);
                    if (isLookupAll == false)
                    {
                        break;
                    }
                }
            }

            return string.Join(splitStr, result);//当list一个元素都没有返回string.empty，当只有一个值时，返回本值，完全符合预期不用if判断
        }




        [ExcelFunction(Category = "查找引用增强", Description = "类似Vlookup的用法，查找某列的值在相对引用区域的对应的返回值。Excel催化剂出品，必属精品！")]
        public static object CZYY查找引用LOOKUP(
                  [ExcelArgument(Description = "查找值的列区域，当有多列作为查找值的列时，需使用辅助函数【FZGetMultiColRange】输入")] object[,] lookupValueRange,
                  [ExcelArgument(Description = "引用匹配列区域，当有多列作为引用匹配列列时，需使用辅助函数【FZGetMultiColRange】输入")] object[,] referenceRange,
                  [ExcelArgument(Description = "引用匹配列所在数据区域内的返回列")] object[,] returnValueRange,
                  [ExcelArgument(Description = "是否模糊匹配，默认为否精确匹配，传入TRUE为模糊匹配")] bool IsFuzzyMatching,
                  [ExcelArgument(Description = "当是否降序匹配，第4参数为模糊匹配时，默认为否升序匹配，传入TRUE为降序匹配")] bool IsDescFuzzyMatching
                )
        {
            ValidParas(lookupValueRange, referenceRange, returnValueRange);

            var referecedataList = GetReferenceData(referenceRange, returnValueRange);

            List<(int Index, object LookupValue)> lookupValueList = new List<(int Index, object LookupValue)>();
            for (int i = 0; i < lookupValueRange.GetLength(0); i++)
            {
                lookupValueList.Add((i, lookupValueRange[i, 0]));
            }

            var grpLookupValues = lookupValueList.GroupBy(s => s.LookupValue);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grpLookupValue in grpLookupValues)
            {
                (object ReferenceField, object ReturnValueField) referenceValue = (null, null);

                if (IsFuzzyMatching)
                {

                    if (IsDescFuzzyMatching)
                    {
                        referecedataList = referecedataList.OrderByDescending(s => GetValueOfOrderField(s.ReferenceField)).ToList();

                        if (grpLookupValue.Key is string)
                        {
                            referenceValue = referecedataList.TakeWhile(s => GetValueOfOrderField(s.ReferenceField.ToString()) >= GetValueOfOrderField(grpLookupValue.Key.ToString())).LastOrDefault();
                        }
                        else if (grpLookupValue.Key is double)
                        {
                            referenceValue = referecedataList.Where(s => s.ReferenceField is double).TakeWhile(s => Convert.ToDouble(s.ReferenceField) >= Convert.ToDouble(grpLookupValue.Key)).LastOrDefault();
                        }
                    }
                    else
                    {
                        referecedataList = referecedataList.OrderBy(s => GetValueOfOrderField(s.ReferenceField)).ToList();

                        if (grpLookupValue.Key is string)
                        {
                            referenceValue = referecedataList.TakeWhile(s => GetValueOfOrderField(s.ReferenceField.ToString()) <= GetValueOfOrderField(grpLookupValue.Key.ToString())).LastOrDefault();
                        }
                        else if (grpLookupValue.Key is double)
                        {
                            referenceValue = referecedataList.Where(s => s.ReferenceField is double).TakeWhile(s => Convert.ToDouble(s.ReferenceField) <= Convert.ToDouble(grpLookupValue.Key)).LastOrDefault();
                        }
                    }


                }
                else
                {
                    if (grpLookupValue.Key is string)
                    {
                        referenceValue = referecedataList.FirstOrDefault(s => s.ReferenceField.ToString() == grpLookupValue.Key.ToString());
                    }
                    else if (grpLookupValue.Key is double)
                    {
                        referenceValue = referecedataList.Where(s => s.ReferenceField is double).FirstOrDefault(s => Convert.ToDouble(s.ReferenceField) == Convert.ToDouble(grpLookupValue.Key));
                    }
                }


                foreach (var item in grpLookupValue)
                {
                    if (referenceValue.ReferenceField != null)
                    {
                        result.Add(item.Index, referenceValue.ReturnValueField);
                    }
                    else
                    {
                        result.Add(item.Index, "");
                    }
                }
            }

            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }

        private static void ValidParas(object[,] lookupValueRange, object[,] referenceRange, object[,] returnValueRange)
        {
            if (lookupValueRange.GetLength(1) != 1 || referenceRange.GetLength(1) != 1 || returnValueRange.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0lookupValueRange = lookupValueRange.GetLength(0);
            int arrDim0referenceRange = referenceRange.GetLength(0);
            int arrDim0returnValueRange = returnValueRange.GetLength(0);

            if (arrDim0referenceRange != arrDim0returnValueRange)
            {
                throw new ArgumentException("参数出错");
            }
        }

        private static void ValidParas(object[,] para1, object[,] para2)
        {
            if (para1.GetLength(1) != 1 || para2.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0para1 = para1.GetLength(0);
            int arrDim0para2 = para2.GetLength(0);

            if (arrDim0para1 != arrDim0para2)
            {
                throw new ArgumentException("参数出错");
            }
        }

        private static List<(object ReferenceField, object ReturnValueField)> GetReferenceData(object[,] referenceRange, object[,] returnValueRange)
        {
            List<(object ReferenceField, object ReturnValueField)> srcDatas = new List<(object ReferenceField, object ReturnValueField)>();

            int arr0Length = referenceRange.GetLength(0);
            for (int i = 0; i < arr0Length; i++)
            {
                (object ReferenceField, object ReturnValueField) row = (null, null);
                row.ReferenceField = referenceRange[i, 0];
                row.ReturnValueField = returnValueRange[i, 0];
                srcDatas.Add(row);
            }
            return srcDatas;
        }

    }
}
