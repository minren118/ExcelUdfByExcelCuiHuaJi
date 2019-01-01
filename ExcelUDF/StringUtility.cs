using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using IExcel = Microsoft.Office.Interop.Excel;
namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {


        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取指定字符。Excel催化剂出品，必属精品！")]
        public static object WB提取指定字符(
            [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
            [ExcelArgument(Description = "查找用于替换的指定字符，多个字符用逗号隔开")] string matchString,
            [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
        )
        {

            var patterns = matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length > 1).ToList();
            string patSingle = string.Join("", matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length == 1));
            patterns.Add($"[{patSingle}]+");
            var pattern = string.Join("|", patterns);


            return RegMatchValue(inputString, pattern, splitStr);
        }
        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取中文。Excel催化剂出品，必属精品！")]
        public static object WB提取中文(
            [ExcelArgument(Description = "待分割的字符串")] string inputString,
            [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
                 )

        {
            string pattern = "[\u4e00-\u9fa5]+";
            return RegMatchValue(inputString, pattern, splitStr);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取数字。Excel催化剂出品，必属精品！")]
        public static object WB提取数字(
                [ExcelArgument(Description = "待分割的字符串")] string inputString,
                [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
         )

        {
            string pattern = @"[0-9][0-9,]*\.[0-9]+|[0-9][0-9,]*";
            return RegMatchValue(inputString, pattern, splitStr);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "提取英文。Excel催化剂出品，必属精品！")]
        public static object WB提取英文(
        [ExcelArgument(Description = "待分割的字符串")] string inputString,
        [ExcelArgument(Description = "当提取多个结果时，结果之间的间隔符")] string splitStr
                )

        {
            string pattern = @"[a-zA-Z]+";
            return RegMatchValue(inputString, pattern, splitStr);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换英文。Excel催化剂出品，必属精品！")]
        public static object WB替换英文(
                [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
                [ExcelArgument(Description = "查找到的字符串替换为此字符串，默认不输入为替换为空")] string replaceString
        )

        {
            string pattern = @"[a-zA-Z]+";
            return RegReplaceValue(inputString, pattern, replaceString);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换中文。Excel催化剂出品，必属精品！")]
        public static object WB替换中文(
        [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
        [ExcelArgument(Description = "查找到的字符串替换为此字符串，默认不输入为替换为空")] string replaceString
                )
        {
            string pattern = @"[\u4e00-\u9fa5]+";
            return RegReplaceValue(inputString, pattern, replaceString);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换中文。Excel催化剂出品，必属精品！")]
        public static object WB替换数字(
            [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
            [ExcelArgument(Description = "查找到的字符串替换为此字符串，默认不输入为替换为空")] string replaceString
             )
        {
            string pattern = @"[0-9][0-9,]*\.[0-9]+|[0-9][0-9,]*";
            return RegReplaceValue(inputString, pattern, replaceString);
        }

        [ExcelFunction(Category = "文本处理_提取替换", Description = "替换指定字符。Excel催化剂出品，必属精品！")]
        public static object WB替换指定字符(
            [ExcelArgument(Description = "待查找替换的字符串")] string inputString,
            [ExcelArgument(Description = "查找用于替换的指定字符，多个字符用逗号隔开")] string matchString,
            [ExcelArgument(Description = "查找到的字符串替换为此字符串，默认不输入为替换为空")] string replaceString
                )
        {

            var patterns = matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length > 1).ToList();
            string patSingle = string.Join("", matchString.Split(new char[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Where(s => s.Length == 1));
            patterns.Add($"[{patSingle}]+");
            var pattern = string.Join("|", patterns);

            return RegReplaceValue(inputString, pattern, replaceString);
        }

        private static string RegReplaceValue(string inputString, string pattern, string replaceString)
        {
            RegexOptions options = RegexOptions.Multiline;
            return Regex.Replace(inputString, pattern, replaceString, options);
        }

        private static string RegMatchValue(string inputString, string pattern, string splitStr)
        {
            RegexOptions options = RegexOptions.Multiline;
            MatchCollection matches = Regex.Matches(inputString, pattern, options);
            return string.Join(splitStr, matches.Cast<Match>().Where(s => !string.IsNullOrEmpty(s.Value)));
        }

        [ExcelFunction(Category = "文本处理", Description = "字符串去重。Excel催化剂出品，必属精品！")]
        public static object TextDistinctChar(
            [ExcelArgument(Description = "待分割的字符串")] string inputString
                )

        {
            return string.Join("", inputString.ToArray().Distinct());
        }

        [ExcelFunction(Category = "文本处理", Description = "字符串反转。Excel催化剂出品，必属精品！")]
        public static object TextReverse(
                [ExcelArgument(Description = "待分割的字符串")] string inputString
        )

        {
            return string.Join("", inputString.ToArray().Reverse());
        }

        [ExcelFunction(Category = "文本处理", Description = "字符串排序。Excel催化剂出品，必属精品！")]
        public static object TextOrder(
        [ExcelArgument(Description = "待分割的字符串")] string inputString,
        [ExcelArgument(Description = "是否降序排列，默认为升序，TRUE为降序，FALSE为升序")] bool isDesc
            )

        {
            if (isDesc)
            {
                return string.Join("", inputString.ToArray().OrderByDescending(s => s));
            }
            else
            {
                return string.Join("", inputString.ToArray().OrderBy(s => s));
            }

        }
        [ExcelFunction(Category = "文本处理", Description = "根据传入的分割字符，对字符串进行分割操作，返回多值。Excel催化剂出品，必属精品！")]
        public static object TextSplits(
            [ExcelArgument(Description = "待分割的字符串")] string inputString,
            [ExcelArgument(Description = "输入分隔字符串，可以引用多个连续单元格或以英文逗号分隔的一个字符串")] object splitValues,
             [ExcelArgument(Description = "返回多值时是按行排列还是按列排列，输入H为按行横向，输入L为按列纵向")] string optAlignHorL
            )
        {
            var splitList = Common.GetSplitStringList(splitValues);
            var result = inputString.Split(splitList.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            return Common.ReturnDataArray(result, optAlignHorL);
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "根据传入的分割字符，对字符串进行分割操作，返回第N个值。Excel催化剂出品，必属精品！")]
        public static object TextSplit(
            [ExcelArgument(Description = "待分割的字符串")] string inputString,
            [ExcelArgument(Description = "输入分隔字符串，可以引用多个连续单元格或以英文逗号分隔的一个字符串")] object splitValues,
             [ExcelArgument(Description = "返回第几个分割后的值，从1开始")] int returnNum
             )
        {
            //inputString.
            var splitList = Common.GetSplitStringList(splitValues);
            var result = inputString.Split(splitList.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            if (returnNum <= result.Length && returnNum > 0)
            {
                return result[returnNum - 1];
            }
            else
            {
                return "";
            }
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "根据传入的待清洗字符，对其时行前后的字符清除。Excel催化剂出品，必属精品！")]
        public static object TextTrim(
            [ExcelArgument(Description = "待前后清除内容的字符串")] string inputString,
            [ExcelArgument(Description = "输入需要清除的字符串，可以引用多个连续单元格或以英文逗号分隔的单个字符串")] object trimValues
            )
        {
            var trimList = Common.GetSplitStringList(trimValues);
            return inputString.Trim(trimList.Select(s => Convert.ToChar(s.Trim().Substring(0, 1))).ToArray());
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "根据传入的待清洗字符，对其时行前面的字符清除。Excel催化剂出品，必属精品！")]
        public static object TextTrimStart(
            [ExcelArgument(Description = "待前面清除内容的字符串")] string inputString,
            [ExcelArgument(Description = "输入需要清除的字符串，可以引用多个连续单元格或以英文逗号分隔的单个字符串")] object trimValues
                 )
        {

            var trimList = Common.GetSplitStringList(trimValues);
            return inputString.TrimStart(trimList.Select(s => Convert.ToChar(s.Trim().Substring(0, 1))).ToArray());
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "根据传入的待清洗字符，对其时行前面的字符清除。Excel催化剂出品，必属精品！")]
        public static object TextTrimEnd(
            [ExcelArgument(Description = "待前面清除内容的字符串")] string inputString,
            [ExcelArgument(Description = "输入需要清除的字符串，可以引用多个连续单元格或以英文逗号分隔的一个ASCII字符串")] object trimValues
                )
        {
            var trimList = Common.GetSplitStringList(trimValues);
            return inputString.TrimEnd(trimList.Select(s => Convert.ToChar(s.Trim().Substring(0, 1))).ToArray());
        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "左侧填充指定个数字符。Excel催化剂出品，必属精品！")]
        public static object TextPadLeft(
            [ExcelArgument(Description = "待填充字符")] string inputString,
            [ExcelArgument(Description = "需要填充的单个字符")] object padStr,
            [ExcelArgument(Description = "返回结果字符串的总位数")] int strLen
         )

        {
            return inputString.PadLeft(strLen, Convert.ToChar(padStr.ToString().Trim().Substring(0, 1)));

        }

        [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "右侧填充指定个数字符。Excel催化剂出品，必属精品！")]
        public static object TextPadRight(
            [ExcelArgument(Description = "待填充字符")] string inputString,
            [ExcelArgument(Description = "需要填充的单个字符")] object padStr,
            [ExcelArgument(Description = "返回结果字符串的总位数")] int strLen
                 )
        {
            return inputString.PadRight(strLen, Convert.ToChar(padStr.ToString().Trim().Substring(0, 1)));
        }
        [ExcelFunction(Category = "文本处理",IsThreadSafe =true, Description = "字符串拼接函数，StringJoinRange：传入的需要写入目标单元格的数据区域，strSplit：用来分割的字符。Excel催化剂出品，必属精品！")]
        //AllowReference = true把传进来的EXCEL区域识别出来。
        public static string StringJoin(
            [ExcelArgument(Description = "输入要拼接的字符串区域")] object StringJoinRange,
            [ExcelArgument(Description = "输入分隔字符串")] string strSplit,
            [ExcelArgument(Description = "包含着在拼接的字符的部分，如双引号、单引号包含着，默认不输为不需要包含的字符")] string strSurround)
        {

            List<object> valuesArr = new List<object>();
            Common.AddValueToList(StringJoinRange, ref valuesArr);
            return string.Join(strSplit, valuesArr.Select(s=> strSurround+s.ToString()+ strSurround));
        }

    [ExcelFunction(Category = "文本处理", IsThreadSafe = true, Description = "字符串拼接函数，在查找的区域查找对应条件下的值，最终用分隔符拼接起来，类似系统函数SUMIF、COUNTIF。Excel催化剂出品，必属精品！")]
    public static object StringJoinIf(
            [ExcelArgument(Description = "查找的区域，若有多列引用，请使用FZGetMultiColRange函数输入")] object[,] lookupRange,
            [ExcelArgument(Description = "用于验证查找区域是否符合的条件")] string criteria,
            [ExcelArgument(Description = "要拼接的字符串区域，请使用FZGetMultiColRange函数输入")] object[,] StringJoinRange,
            [ExcelArgument(Description = "输入分隔字符串")] string strSplit,
            [ExcelArgument(Description = "是否精确匹配，默认为否")] bool isExactMatch = false)
    {
        //输入不规范，不是单列
        if (lookupRange.GetLength(0) != StringJoinRange.GetLength(0) || lookupRange.GetLength(1) != 1 || StringJoinRange.GetLength(1) != 1)
        {
            return ExcelError.ExcelErrorNA;
        }

        List<string> list = new List<string>();

        for (int i = 0; i < lookupRange.GetLength(0); i++)
        {
            if (isExactMatch)
            {
                if (lookupRange.GetValue(i, 0).ToString() == criteria)
                {
                    list.Add(StringJoinRange.GetValue(i, 0).ToString());
                }

            }
            else
            {
                if (lookupRange.GetValue(i, 0).ToString().Contains(criteria))
                {
                    list.Add(StringJoinRange.GetValue(i, 0).ToString());
                }
            }
        }

        return string.Join(strSplit, list);

    }

}
}
