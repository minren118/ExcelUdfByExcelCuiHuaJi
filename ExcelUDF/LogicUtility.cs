using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "逻辑判断_日期", Description = "是否为日期，EXCEL可识别的日期为从1900-01-00至9999-12-31，即数字为0至2958465之间。Excel催化剂出品，必属精品！")]
        public static bool IsDate(
            [ExcelArgument(Description = "输入的值")] object input)
        {
            double number;
            NumberStyles style = NumberStyles.Any;
            CultureInfo culture = CultureInfo.CurrentCulture;

            if (Common.IsMissOrEmpty(input))
            {
                return false;
            }
            if (double.TryParse(input.ToString(), style, culture, out number))
            {
                //1900年1月1日至9999-12-31日之间
                if (number >= 0 && number < 2958466)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        [ExcelFunction(Category = "逻辑判断_日期", Description = "是否为日期，EXCEL可识别的日期为从1900-01-00至9999-12-31，即数字为0至2958465之间。Excel催化剂出品，必属精品！")]
        public static object IsDateBetween(
            [ExcelArgument(Description = "输入的值")] object input,
             [ExcelArgument(Description = "开始日期")] object optstartDate,
              [ExcelArgument(Description = "结束日期")] object optEndDate
            )
        {
            double number;
            NumberStyles style = NumberStyles.Any;
            CultureInfo culture = CultureInfo.CurrentCulture;

            if (Common.IsMissOrEmpty(input))
            {
                return false;
            }

            if (Common.IsMissOrEmpty(optstartDate))
            {
                optstartDate = new DateTime(1899, 12, 31);
            }

            if (Common.IsMissOrEmpty(optEndDate))
            {
                optEndDate = DateTime.MaxValue;
            }

            double dbStartDate = 0;
            double dbEndDate = 0;
            DateTime dtStartDate = DateTime.MinValue;
            DateTime dtEndDate = DateTime.MinValue;

            if (!DateTime.TryParse(optstartDate.ToString(), out dtStartDate))
            {
                if (double.TryParse(optstartDate.ToString(), out dbStartDate))
                {
                    if (dbStartDate > 0 && dbStartDate < 2958466)
                    {
                        dtStartDate = new DateTime(1899, 12, 30).AddDays(dbStartDate);
                    }
                }
            }

            if (!DateTime.TryParse(optEndDate.ToString(), out dtEndDate))
            {

                if (double.TryParse(optEndDate.ToString(), out dbEndDate))
                {
                    if (dbEndDate > 0 && dbEndDate < 2958466)
                    {
                        dtEndDate = new DateTime(1899, 12, 30).AddDays(dbEndDate);
                    }
                }
            }

            //把传入的开始、结束日期转化为double格式，因EXCEL输入的日期为double格式
            //检验输入的开始、结束日期是否在EXCEL的识别日期范围内1900-1-0至9999-12-31
            //用dateTime的addDays方法把EXCEL日期转化为.net日期

            if (double.TryParse(input.ToString(), style, culture, out number))
            {
                //1900年1月1日至9999-12-31日之间
                DateTime dt = new DateTime(1899, 12, 30).AddDays(number);
                if (number >= 0 && number <= 2958466 && dt >= dtStartDate && dt <= dtEndDate)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        [ExcelFunction(Category = "逻辑判断_文本", Description = "判断查找字符是否在源字符串内，如果存在返回true，否则返回false，示例：lookupValue:AB,sourceString:ABCDE,AB在ABCDE中，结果返回true。Excel催化剂出品，必属精品！")]
        public static bool IsTextContains(
            [ExcelArgument(Description = "查找字符串，如ABCDE")] string sourceString,
            [ExcelArgument(Description = "查找条件，如AB")] string lookupValue,
            [ExcelArgument(Description = "是否区分大小写，默认为FALSE不区分，TRUE为区分大小写")] bool isCaseSensitive
            )
        {
            if (isCaseSensitive)
            {
                return sourceString.Contains(lookupValue);
            }
            else
            {
                return sourceString.ToUpper().Contains(lookupValue.ToUpper());
            }

        }

        [ExcelFunction(Category = "逻辑判断_文本", Description = "判断查找字符是否在源字符串内，如果存在返回true，否则返回false，示例：lookupValue:AB,sourceString:ABCDE,AB在ABCDE中，结果返回true。Excel催化剂出品，必属精品！")]
        public static bool IsTextContainsWithMultiLookupValues(
            [ExcelArgument(Description = "查找字符串，如ABCDE")] string sourceString,
            [ExcelArgument(Description = "多个查找条件，可以引用多个连续单元格或以英文逗号分隔的一个字符串，如A,B")] object lookupValues,
            [ExcelArgument(Description = "是否区分大小写，默认为FALSE不区分，TRUE为区分大小写")]  bool isCaseSensitive
            )
        {
            List<string> listLookupValues = Common.GetSplitStringList(lookupValues);
            if (isCaseSensitive)
            {
                return listLookupValues.Any(s => sourceString.Contains(s));
            }
            else
            {
                return listLookupValues.Any(s => sourceString.ToUpper().Contains(s.ToUpper()));
            }

        }


        [ExcelFunction(Category = "逻辑判断_文本", Description = "判断字符是否在指定的字符串集合内，如果存在返回true，否则返回false，示例：sourceStrings=AB,CD,E,lookupValue=CD,strSplit=',',CD在｛AB，CD，E｝的集合中，返回true。Excel催化剂出品，必属精品！")]
        public static bool IsTextContainsWithSplit(
        [ExcelArgument(Description = "查找字符串集合")] string sourceStrings,
        [ExcelArgument(Description = "查找条件")] string lookupValue,
        [ExcelArgument(Description = "查找字符串集合内用于分割的字符，注意中英文符号要与查找字符串集合一致,若传入多个分隔符，使用|隔开。")] string strSplit,
        [ExcelArgument(Description = "是否区分大小写，默认为FALSE不区分，TRUE为区分大小写")]  bool isCaseSensitive
                )
        {
            string[] strsplits;
            //当传入的strsplit是以|结尾或开头的，就当作一个字符串处理，不进行strsplit的分隔
            if (strSplit.Trim() == "|")
            {
                strsplits = new string[] { strSplit.Trim() };
            }
            else
            {
                strsplits = strSplit.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();
            }
            if (isCaseSensitive)
            {
                return sourceStrings.Split(strsplits, StringSplitOptions.RemoveEmptyEntries).Equals(lookupValue);
            }
            else
            {
                return sourceStrings.Split(strsplits, StringSplitOptions.RemoveEmptyEntries).Select(s => s.ToUpper()).Equals(lookupValue.ToUpper());
            }


        }

        [ExcelFunction(Category = "逻辑判断_区域", Description = "判断查找区域内查找是否有查找条件相同的，如果存在返回true，否则返回false，示例：lookupValue:ABC,sourceRange:A1：A3｛ABC，BCD，EDF｝,ABC在sourceRange中，结果返回true。Excel催化剂出品，必属精品！")]
        public static object IsRangeContains(
                [ExcelArgument(Description = "查找的区域如：A1：A3｛ABC，BCD，EDF}", AllowReference = true)] object[,] sourceRange,
                [ExcelArgument(Description = "查找条件，如B1的值ABC")] object lookupValue,
                [ExcelArgument(Description = "是否模糊匹配，传入0为false,非0为true,默认是false，即为精确匹配，如AB与｛ABC，BCD，EDF｝不匹配")] bool IsFuzzyMatch
                )
        {

            if (!(lookupValue is ExcelMissing))
            {
                for (int i = 0; i < sourceRange.GetLength(0); i++)
                {
                    for (int j = 0; j < sourceRange.GetLength(1); j++)
                    {
                        if (IsFuzzyMatch == true)
                        {
                            //模糊匹配一定是用字符串匹配
                            if (sourceRange[i, j].ToString().Contains(lookupValue.ToString()))
                            {
                                return true;
                            }
                        }
                        else
                        {
                            //精确匹配用的是数字或字符串，让object自己判断
                            if (sourceRange[i, j].Equals(lookupValue))
                            {
                                return true;
                            }
                        }
                    }
                }
                return false;

            }
            else
            {
                return ExcelError.ExcelErrorValue;
            }
        }

        [ExcelFunction(Category = "逻辑判断_区域", Description = "判断查找区域内查找是否有查找条件相同的，如果存在返回true，否则返回false，示例：lookupValue:ABC,sourceRange:A1：A3｛ABC，BCD，EDF｝,ABC在sourceRange中，结果返回true。Excel催化剂出品，必属精品！")]
        public static object IsRangeContainsWithMultiLookupValues(
              [ExcelArgument(Description = "查找的区域如：A1：A3｛ABC，BCD，EDF}", AllowReference = true)] object[,] sourceRange,
              [ExcelArgument(Description = "多个查找条件，可以引用多个连续单元格或以英文逗号分隔的一个字符串，如A,B")] object lookupValues,
              [ExcelArgument(Description = "是否模糊匹配，传入0为false,非0为true,默认是false，即为精确匹配，如AB与｛ABC，BCD，EDF｝不匹配")] bool IsFuzzyMatch
              )
        {

            List<string> listLookupValues = Common.GetSplitStringList(lookupValues);

            for (int i = 0; i < sourceRange.GetLength(0); i++)
            {
                for (int j = 0; j < sourceRange.GetLength(1); j++)
                {
                    if (IsFuzzyMatch == true)
                    {
                        //模糊匹配一定是用字符串匹配
                        if (listLookupValues.Any(s => sourceRange[i, j].ToString().Contains(s)))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        //精确匹配用的是数字或字符串，让object自己判断
                        if (listLookupValues.Any(s=>s== sourceRange[i, j].ToString()))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;

        }

        [ExcelFunction(Category = "逻辑判断_区域", Description = "判断查找区域内查找是否有至少2个满足查找条件的值，如果存在返回true，否则返回false，示例：lookupValue:ABC,sourceRange:A1：A3｛ABC，ABC，EDF｝,ABC在sourceRange中出现2次，结果返回true。Excel催化剂出品，必属精品！")]
        public static object IsRangeContainsDuplicatedValue(
        [ExcelArgument(Description = "查找的区域如：A1：A3｛ABC，ABC，EDF｝")] object[,] sourceRange,
        [ExcelArgument(Description = "查找条件，如B1的值ABC")] object lookupValue,
        [ExcelArgument(Description = "是否模糊匹配，传入0为false,非0为true,默认是false，即为精确匹配，如AB与｛ABC，BCD，EDF｝不匹配")] bool isFuzzyMatch
        )
        {
            byte DuplicatedCount = 0;
            if (!(lookupValue is ExcelMissing))
            {
                for (int i = 0; i < sourceRange.GetLength(0); i++)
                {
                    for (int j = 0; j < sourceRange.GetLength(1); j++)
                    {
                        if (isFuzzyMatch == true)
                        {
                            //模糊匹配一定是用字符串匹配
                            if (sourceRange[i, j].ToString().Contains(lookupValue.ToString()))
                            {
                                DuplicatedCount++;
                                if (DuplicatedCount > 1)
                                {
                                    return true;
                                }

                            }
                        }
                        else
                        {
                            //精确匹配用的是数字或字符串，让object自己判断
                            if (sourceRange[i, j].Equals(lookupValue))
                            {
                                DuplicatedCount++;
                                if (DuplicatedCount > 1)
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                return false;

            }
            else
            {
                return ExcelError.ExcelErrorValue;
            }

        }


        [ExcelFunction(Category = "逻辑判断_文本", Description = "判断字符串是否以某字符开头。Excel催化剂出品，必属精品！")]
        public static bool IsTextStartsWith(
            [ExcelArgument(Description = "输入待查找的字符串")] string inputStr,
            [ExcelArgument(Description = "需要查找的开头部分的字符串值")] string value
)
        {
            return inputStr.StartsWith(value);
        }

        [ExcelFunction(Category = "逻辑判断_文本", Description = "判断字符串是否以某字符开头。Excel催化剂出品，必属精品！")]
        public static bool IsTextEndsWith(
            [ExcelArgument(Description = "输入待查找的字符串")] string inputStr,
            [ExcelArgument(Description = "需要查找的结尾部分的字符串值")] string value
)
        {
            //inputStr.
            return inputStr.EndsWith(value);
        }



    }
}
