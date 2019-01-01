using ExcelDna.Integration;
using Microsoft.International.Converters.PinYinConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "中文相关", Description = "繁体转简体。Excel催化剂出品，必属精品！")]
        public static object TraditionalChineseToSimplified(
                                                    [ExcelArgument(Description = "输入繁体中文字符")] string inputString
                                                    )
        {

            return ChineseConverter.Convert(inputString, ChineseConversionDirection.TraditionalToSimplified);
        }



        [ExcelFunction(Category = "中文相关", Description = "简体转繁体。Excel催化剂出品，必属精品！")]
        public static object SimplifiedChieseToTraditional(
                                            [ExcelArgument(Description = "输入简体中文字符")] string inputString
                                            )
        {

            return ChineseConverter.Convert(inputString, ChineseConversionDirection.SimplifiedToTraditional);
        }



        [ExcelFunction(Category = "中文相关", Description = "数字转换为多个单元格存放的效果，财务用途。Excel催化剂出品，必属精品！")]
        public static object NumberConvertToMultiCells(
            [ExcelArgument(Description = "需要拆分的原始数字")] double inputNumber,
            [ExcelArgument(Description = "拆分的总列数，含角分。拆分到亿为单位的话为11")] int colNum,
            [ExcelArgument(Description = "拆分的总列数，含角分。拆分到亿为单位的话为11")] bool hasLeadingZero
            )
        {
            if (colNum == 0)
            {
                colNum = 11;
            }
            string numberString = inputNumber.ToString();
            if (numberString.Contains("."))//含小数点
            {

                numberString = numberString.Split('.')[0] + numberString.Split('.')[1].PadRight(2, '0');
            }
            if (hasLeadingZero)
            {
                numberString = numberString.PadLeft(colNum, '0');

                numberString = numberString.Substring(numberString.Length - colNum, colNum);
            }
            else
            {
                numberString = numberString.PadLeft(colNum, ' ');
                numberString = numberString.Substring(numberString.Length - colNum, colNum);
            }


            return Common.ReturnDataArray(numberString.ToArray().Select(s => s.ToString().Replace(" ", "")).ToArray(), "H");
        }




        [ExcelFunction(Category = "中文相关", Description = "数字转换为大写中文金额。Excel催化剂出品，必属精品！")]
        public static string NumberConvertToChineseCapitalAmount(
            [ExcelArgument(Description = "传入需要转换大写的数字")] double inputNumber)
        {

            // 大写数字数组
            string[] num = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
            // 数量单位数组，个位数为空
            string[] unit = { "", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿", "拾", "佰", "仟", "兆" };
            string d = inputNumber.ToString();
            string zs = string.Empty;// 整数
            string xs = string.Empty;// 小数
            int i = d.IndexOf(".");
            string str = string.Empty;
            if (i > -1)
            {
                // 仅考虑两位小数
                zs = d.Substring(0, i);
                xs = d.Substring(i + 1, d.Length - i - 1);
                str = "元";
                if (xs.Length == 1)
                    str = str + xs + "角";
                else if (xs.Length == 2)
                    str = str + xs.Substring(0, 1) + "角" + xs.Substring(1, 1) + "分";
            }
            else
            {
                zs = d;
                str = "元整";
            }
            // 处理整数部分
            if (!string.IsNullOrEmpty(zs))
            {
                i = 0;
                // 从整数部分个位数起逐一添加单位
                foreach (char s in zs.Reverse())
                {
                    str = s.ToString() + unit[i] + str;
                    i++;
                }
            }
            // 将阿拉伯数字替换成中文大写数字
            for (int m = 0; m < 10; m++)
            {
                str = str.Replace(m.ToString(), num[m]);
            }
            // 替换零佰、零仟、零拾之类的字符
            str = Regex.Replace(str, "[零]+仟", "零");
            str = Regex.Replace(str, "[零]+佰", "零");
            str = Regex.Replace(str, "[零]+拾", "零");
            str = Regex.Replace(str, "[零]+亿", "亿");
            str = Regex.Replace(str, "[零]+万", "万");
            str = Regex.Replace(str, "[零]+", "零");
            str = Regex.Replace(str, "亿[万|仟|佰|拾]+", "亿");
            str = Regex.Replace(str, "万[仟|佰|拾]+", "万");
            str = Regex.Replace(str, "仟[佰|拾]+", "仟");
            str = Regex.Replace(str, "佰拾", "佰");
            str = Regex.Replace(str, "[零]+元", "元");
            str = Regex.Replace(str, "[零]+元整", "元整");
            return str;


        }
        [ExcelFunction(Category = "中文相关", Description = "中文转全拼。Excel催化剂出品，必属精品！")]
        public static string ChineseConvertToPinYinAllSpell(
            [ExcelArgument(Description = "转入需要转换拼音的中文字符串")] string inputChineseChar,
            [ExcelArgument(Description = "中文拼音间的间隔符")] string separateString
            )
        {
            return PingYinHelper.ConvertToAllSpell(inputChineseChar, separateString);
        }

        [ExcelFunction(Category = "中文相关", Description = "中文转首字母拼单。Excel催化剂出品，必属精品！")]
        public static string ChineseConvertToPinYinFirstSpell([ExcelArgument(Description = "转入需要转换拼音的中文字符串")] string inputChineseChar)
        {
            return PingYinHelper.GetFirstSpell(inputChineseChar);
        }


        [ExcelFunction(Category = "中文相关", Description = "大写中文金额转换为数字。Excel催化剂出品，必属精品！")]
        public static decimal ChineseCapitalAmountConvertToNumber([ExcelArgument(Description = "传入需要转换大写的数字")] string inputChinseUpperMoney)
        {

            string beforeYiString;
            string afterYiString;
            decimal result = 0;

            inputChinseUpperMoney = inputChinseUpperMoney.Replace("整", "");
            if (inputChinseUpperMoney == "元") //零元整
            {
                return 0;
            }
            else if (inputChinseUpperMoney.Contains("亿"))//有亿
            {
                beforeYiString = inputChinseUpperMoney.Split(new string[] { "亿" }, StringSplitOptions.RemoveEmptyEntries)[0];

                if (beforeYiString.Contains("兆"))
                {
                    string beforeZhao = beforeYiString.Split(new string[] { "兆" }, StringSplitOptions.RemoveEmptyEntries)[0];
                    result = ConvertNameToSmall(beforeZhao) * 1000000000000;
                    string afterZhao = beforeYiString.Split(new string[] { "兆" }, StringSplitOptions.RemoveEmptyEntries)[1] + "元";
                    result = result + CalculateWithoutYi(afterZhao, 0) * 100000000;
                }
                else
                {
                    result = CalculateWithoutYi(beforeYiString, result) * 100000000;
                }

                afterYiString = inputChinseUpperMoney.Split(new string[] { "亿" }, StringSplitOptions.RemoveEmptyEntries)[1];

                result = CalculateWithoutYi(afterYiString, result);
            }
            else
            {
                afterYiString = inputChinseUpperMoney;
                result = CalculateWithoutYi(afterYiString, result);
            }
            return result;
        }

        /// <summary>
        /// 转换含万字不含亿的部分
        /// </summary>
        /// <param name="afterYiString"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        private static decimal CalculateWithoutYi(string afterYiString, decimal result)
        {
            if (afterYiString.Contains("万"))
            {

                string beforWan = afterYiString.Split(new string[] { "万" }, StringSplitOptions.RemoveEmptyEntries)[0] + "元";
                result = result + GetQianYiXia(beforWan) * 10000;

                string afterWan = afterYiString.Split(new string[] { "万" }, StringSplitOptions.RemoveEmptyEntries)[1];
                result = result + GetQianYiXia(afterWan);
            }
            else
            {
                return GetQianYiXia(afterYiString);
            }

            return result;
        }


        private static int ConvertNameToSmall(string str)
        {
            int number = 0;
            switch (str)
            {
                case "零": number = 0; break;
                case "壹": number = 1; break;
                case "贰": number = 2; break;
                case "叁": number = 3; break;
                case "肆": number = 4; break;
                case "伍": number = 5; break;
                case "陆": number = 6; break;
                case "柒": number = 7; break;
                case "捌": number = 8; break;
                case "玖": number = 9; break;
                default: break;
            }
            return number;
        }

        /// <summary>
        /// 千以下的拼接字符串计算数字求和
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static decimal GetQianYiXia(string str)
        {
            var arr = str.ToArray().Select(s => s.ToString()).ToArray();
            string resultString = string.Empty;
            for (int i = 0; i < arr.Count(); i++)
            {
                if ("零壹贰叁肆伍陆柒捌玖".Contains(arr[i]))
                {
                    if (string.IsNullOrEmpty(resultString))
                    {
                        resultString = ConvertNameToSmall(arr[i]).ToString();
                    }
                    else
                    {
                        resultString = $"{resultString}+{ConvertNameToSmall(arr[i]).ToString()}";
                    }
                    ;
                }
                else
                {
                    switch (arr[i])
                    {
                        case "仟":
                            resultString = $"{resultString}*1000";
                            break;
                        case "佰":
                            resultString = $"{resultString}*100";
                            break;
                        case "拾":
                            resultString = $"{resultString}*10";
                            break;
                        case "元":
                        case "圆":
                            resultString = $"{resultString}*1";
                            break;
                        case "角":
                            resultString = $"{resultString}*0.1";
                            break;
                        case "分":
                            resultString = $"{resultString}*0.01";
                            break;
                    }
                }
            }
            return decimal.Parse(new System.Data.DataTable().Compute(resultString, "").ToString());

        }


    }

    public class PingYinHelper
    {
        private static Encoding gb2312 = Encoding.GetEncoding("GB2312");

        /// <summary>
        /// 汉字转全拼
        /// </summary>
        /// <param name="strChinese"></param>
        /// <returns></returns>
        internal static string ConvertToAllSpell(string strChinese, string splitString)
        {
            try
            {
                if (strChinese.Length != 0)
                {
                    List<string> list = new List<string>();
                    for (int i = 0; i < strChinese.Length; i++)
                    {
                        var chr = strChinese[i];
                        list.Add(GetSpell(chr));
                    }

                    return string.Join(splitString, list);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("全拼转化出错！" + e.Message);
            }

            return string.Empty;
        }

        /// <summary>
        /// 汉字转首字母
        /// </summary>
        /// <param name="strChinese"></param>
        /// <returns></returns>
        internal static string GetFirstSpell(string strChinese)
        {
            //NPinyin.Pinyin.GetInitials(strChinese)  有Bug  洺无法识别
            //return NPinyin.Pinyin.GetInitials(strChinese);

            try
            {
                if (strChinese.Length != 0)
                {
                    StringBuilder fullSpell = new StringBuilder();
                    for (int i = 0; i < strChinese.Length; i++)
                    {
                        var chr = strChinese[i];
                        fullSpell.Append(GetSpell(chr)[0]);
                    }

                    return fullSpell.ToString().ToUpper();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("首字母转化出错！" + e.Message);
            }

            return string.Empty;
        }

        private static string GetSpell(char chr)
        {
            var coverchr = NPinyin.Pinyin.GetPinyin(chr);

            bool isChineses = ChineseChar.IsValidChar(coverchr[0]);
            if (isChineses)
            {
                ChineseChar chineseChar = new ChineseChar(coverchr[0]);
                foreach (string value in chineseChar.Pinyins)
                {
                    if (!string.IsNullOrEmpty(value))
                    {
                        return value.Remove(value.Length - 1, 1);
                    }
                }
            }

            return coverchr;

        }
    }
}
