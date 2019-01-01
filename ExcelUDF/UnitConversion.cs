using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "单位换算", Description = "不同进制数间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS进制转换(
            [ExcelArgument(Description = "输入待转换的值")] string input,
            [ExcelArgument(Description = "输入值的进制数")] int fromType,
            [ExcelArgument(Description = "需要转换的进制数")] int toType
            )
        {

            string output = input;
            switch (fromType)
            {
                case 2:
                    output = ConvertGenericBinaryFromBinary(input, toType);
                    break;
                case 8:
                    output = ConvertGenericBinaryFromOctal(input, toType);
                    break;
                case 10:
                    output = ConvertGenericBinaryFromDecimal(input, toType);
                    break;
                case 16:
                    output = ConvertGenericBinaryFromHexadecimal(input, toType);
                    break;
                default:
                    break;
            }
            return output;

        }


        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS颜色RGB转换Html(
                    [ExcelArgument(Description = "输入R值，范围0-255")] object inputR,
                    [ExcelArgument(Description = "输入G值，范围0-255")] object inputG,
                    [ExcelArgument(Description = "输入B值，范围0-255")] object inputB
                     )
        {
            try
            {
                int R = Convert.ToInt32(inputR.ToString().Trim());
                int G = Convert.ToInt32(inputG.ToString().Trim());
                int B = Convert.ToInt32(inputB.ToString().Trim());

                return ColorTranslator.ToHtml(Color.FromArgb(255, R, G, B));

            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }



        }

        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS颜色RGB转换Ole(
            [ExcelArgument(Description = "输入R值，范围0-255")] object inputR,
            [ExcelArgument(Description = "输入G值，范围0-255")] object inputG,
            [ExcelArgument(Description = "输入B值，范围0-255")] object inputB
             )
        {
            try
            {
                int R = Convert.ToInt32(inputR.ToString().Trim());
                int G = Convert.ToInt32(inputG.ToString().Trim());
                int B = Convert.ToInt32(inputB.ToString().Trim());

                return ColorTranslator.ToOle(Color.FromArgb(255, R, G, B));
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS颜色Ole转换RGB(
                [ExcelArgument(Description = "输入Ole值，OFFICE软件的Color属性")] int inputOle
                     )
        {
            try
            {
               Color color= ColorTranslator.FromOle(inputOle);
                return $"{color.R},{color.G},{color.B}"; 
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS颜色Ole转换Html(
        [ExcelArgument(Description = "输入Ole值，OFFICE软件的Color属性")] int inputOle
             )
        {
            try
            {
                Color color = ColorTranslator.FromOle(inputOle);
                
                return ColorTranslator.ToHtml(Color.FromArgb(255,color.R,color.G,color.B));
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS颜色Html转RGB(
                [ExcelArgument(Description = "输入网页Html格式颜色值，由#开头")] string inputHtmlColor
                         )
        {
            try
            {
                Color color = ColorTranslator.FromHtml(inputHtmlColor);
                return $"{color.R},{color.G},{color.B}";
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "单位换算", Description = "不同颜色间表示方法间的转换。Excel催化剂出品，必属精品！")]
        public static object DWHS颜色Html转Ole(
        [ExcelArgument(Description = "输入网页Html格式颜色值，由#开头")] string inputHtmlColor
                 )
        {
            try
            {
                Color color = ColorTranslator.FromHtml(inputHtmlColor);
                return ColorTranslator.ToOle(color);
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorNA;
            }
        }


        /// <summary>
        /// 从二进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromBinary(string input, int toType)
        {
            switch (toType)
            {
                case 8:
                    //先转换成十进制然后转八进制
                    input = Convert.ToString(Convert.ToInt32(input, 2), 8);
                    break;
                case 10:
                    input = Convert.ToInt32(input, 2).ToString();
                    break;
                case 16:
                    input = Convert.ToString(Convert.ToInt32(input, 2), 16);
                    break;
                default:
                    break;
            }
            return input;
        }

        /// <summary>
        /// 从八进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromOctal(string input, int toType)
        {
            switch (toType)
            {
                case 2:
                    input = Convert.ToString(Convert.ToInt32(input, 8), 2);
                    break;
                case 10:
                    input = Convert.ToInt32(input, 8).ToString();
                    break;
                case 16:
                    input = Convert.ToString(Convert.ToInt32(input, 8), 16);
                    break;
                default:
                    break;
            }
            return input;
        }

        /// <summary>
        /// 从十进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromDecimal(string input, int toType)
        {
            string output = "";
            int intInput = Convert.ToInt32(input);
            switch (toType)
            {
                case 2:
                    output = Convert.ToString(intInput, 2);
                    break;
                case 8:
                    output = Convert.ToString(intInput, 8);
                    break;
                case 16:
                    output = Convert.ToString(intInput, 16);
                    break;
                default:
                    output = input;
                    break;
            }
            return output;
        }

        /// <summary>
        /// 从十六进制转换成其他进制
        /// </summary>
        /// <param name="input"></param>
        /// <param name="toType"></param>
        /// <returns></returns>
        private static string ConvertGenericBinaryFromHexadecimal(string input, int toType)
        {
            switch (toType)
            {
                case 2:
                    return Convert.ToString(Convert.ToInt32(input, 16), 2);

                case 8:
                    return Convert.ToString(Convert.ToInt32(input, 16), 8);

                case 10:
                    return Convert.ToInt32(input, 16).ToString();

                default:
                    return string.Empty;
            }
        }


        [ExcelFunction(Category = "单位换算", Description = "字符转ASCII编号。Excel催化剂出品，必属精品！")]
        public static object DWHS字符转ASCCII码(
             [ExcelArgument(Description = "输入所要查找的单个字符")] string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                return ExcelError.ExcelErrorNA;
            }

        }

        [ExcelFunction(Category = "单位换算", Description = "ASCII编转字符。Excel催化剂出品，必属精品！")]
        public static object DWHSASCCII转字符(
             [ExcelArgument(Description = "输入0-255之间的整数ASCII码")] int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                return ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "单位换算", Description = "Unix timestamp转普通日期。Excel催化剂出品，必属精品！")]
        public static object DWHSUnixTimestampToDatetime(
           [ExcelArgument(Description = "输入UnixTimestamp")] Int64 inputUnixTimestamp)

        {
            if (inputUnixTimestamp.ToString().Length == 10)
            {
                inputUnixTimestamp = inputUnixTimestamp * 1000;
            }
            System.DateTime time = System.DateTime.MinValue;
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            time = startTime.AddMilliseconds(inputUnixTimestamp);
            Common.ChangeNumberFormat("yyyy-mm-dd hh:mm:ss");
            return time;

        }

        [ExcelFunction(Category = "单位换算", Description = "普通日期转Unix timestamp。Excel催化剂出品，必属精品！")]
        public static object DWHSDatetimeToUnixTimestamp(
            [ExcelArgument(Description = "输入UnixTimestamp")] DateTime inputDateTime,
             [ExcelArgument(Description = "是否精确到秒，TRUE为秒，FALSE为毫秒")] bool isSecond
            )
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1, 0, 0, 0, 0));
            //intResult = (time- startTime).TotalMilliseconds;
            long unixTime = (inputDateTime.Ticks - startTime.Ticks) / 10000;            //除10000调整为13位
            Common.ChangeNumberFormat("0");
            return isSecond ? unixTime / 1000 : unixTime;
        }

        [ExcelFunction(Category = "单位换算", Description = "数字转万为单位。Excel催化剂出品，必属精品！")]
        public static object DWHSNumberConverToWan(
           [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
           [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits,
           [ExcelArgument(Description = "是否需要带上万字样的数字格式")] bool isNumberFormatWan
           )
        {
            double ratio = 0.0001;
            if (isNumberFormatWan)
            {
                string numString = num_digits is ExcelMissing ? (new string('0', 2)) : (new string('0', Convert.ToInt32(num_digits)));
                Common.ChangeNumberFormat($"#,##0.{numString}万;-#,##0.{numString}万");
            }
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "数字转亿为单位。Excel催化剂出品，必属精品！")]
        public static object DWHSNumberConverToYi(
           [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
           [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits,
           [ExcelArgument(Description = "是否需要带上万字样的数字格式")] bool isNumberFormatYi)
        {
            double ratio = 0.00000001;
            if (isNumberFormatYi)
            {
                string numString = num_digits is ExcelMissing ? (new string('0', 2)) : (new string('0', Convert.ToInt32(num_digits)));
                Common.ChangeNumberFormat($"#,##0.{numString}亿;-#,##0.{numString}亿");
            }
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }





        [ExcelFunction(Category = "单位换算", Description = "美国加仑转升。Excel催化剂出品，必属精品！")]
        public static object DWHS美加仑转升(
           [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
           [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 3.785;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "英国加仑转升。Excel催化剂出品，必属精品！")]
        public static object DWHS英加仑转升(
           [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
           [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 4.546;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "升转汤勺。Excel催化剂出品，必属精品！")]
        public static object DWHS升转汤勺(
           [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
           [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 66.67;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "升转调羹。Excel催化剂出品，必属精品！")]
        public static object DWHS升转调羹(
               [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
               [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 200;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "千米转英里。Excel催化剂出品，必属精品！")]
        public static object DWHS千米转英里(
       [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
       [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.621;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "米转英尺。Excel催化剂出品，必属精品！")]
        public static object DWHS米转英尺(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 3.281;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "米转码。Excel催化剂出品，必属精品！")]
        public static object DWHS米转英码(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 1.094;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }


        [ExcelFunction(Category = "单位换算", Description = "米转英寸。Excel催化剂出品，必属精品！")]
        public static object DWHS米转英寸(
               [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
               [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 39.37;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }


        [ExcelFunction(Category = "单位换算", Description = "海里转英里。Excel催化剂出品，必属精品！")]
        public static object DWHS海里转英里(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 1.1516;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "海里转千米。Excel催化剂出品，必属精品！")]
        public static object DWHS海里转千米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 1.852;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "汤勺转毫升。Excel催化剂出品，必属精品！")]
        public static object DWHS汤勺转毫升(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 15;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "调羹转毫升。Excel催化剂出品，必属精品！")]
        public static object DWHS调羹转毫升(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 5;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "英国液量盎司转毫升。Excel催化剂出品，必属精品！")]
        public static object DWHS英液量盎司转毫升(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 28.41;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "美国液量盎司转毫升。Excel催化剂出品，必属精品！")]
        public static object DWHS美液量盎司转毫升(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 29.57;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }


        [ExcelFunction(Category = "单位换算", Description = "英里转千米。Excel催化剂出品，必属精品！")]
        public static object DWHS英里转千米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 1.6093;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "英尺转米。Excel催化剂出品，必属精品！")]
        public static object DWHS英尺转米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.3048;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "英寸转米。Excel催化剂出品，必属精品！")]
        public static object DWHS英寸转米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.0254;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "中国寸转米。Excel催化剂出品，必属精品！")]
        public static object DWHS中国寸转米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.0333;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "中国尺转米。Excel催化剂出品，必属精品！")]
        public static object DWHS中国尺转米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.3333;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }


        [ExcelFunction(Category = "单位换算", Description = "码转米。Excel催化剂出品，必属精品！")]
        public static object DWHS码转米(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.9144;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "公顷转平方公里。Excel催化剂出品，必属精品！")]
        public static object DWHS公顷转平方公里(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.01;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "中国亩转公顷。Excel催化剂出品，必属精品！")]
        public static object DWHS中国亩转公顷(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {

            double ratio = 0.0667;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "中国亩转公顷。Excel催化剂出品，必属精品！")]
        public static object DWHS摄氏度转华氏度(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 33.8;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "中国亩转公顷。Excel催化剂出品，必属精品！")]
        public static object DWHS华氏度转摄氏度(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 0.0296;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "贵金属的金衡盎司转克。Excel催化剂出品，必属精品！")]
        public static object DWHS金衡盎司转克(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 31.10;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "日常普通常规的常衡盎司转克。Excel催化剂出品，必属精品！")]
        public static object DWHS常衡盎司转克(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 28.35;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "日常普通常规的常衡磅转克。Excel催化剂出品，必属精品！")]
        public static object DWHS常衡磅转克(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 453.59;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        [ExcelFunction(Category = "单位换算", Description = "贵金属的金衡磅转克。Excel催化剂出品，必属精品！")]
        public static object DWHS金衡磅转克(
            [ExcelArgument(Description = "输入要转换的数字")] double inputNumber,
            [ExcelArgument(Description = "需要保留小数点位数，省略不进行小数位数四舍五入")] object num_digits)
        {
            double ratio = 373.24;
            return ConVertByRatio(ratio, inputNumber, num_digits);
        }

        private static object ConVertByRatio(double ratio, double inputNumber, object num_digits)
        {
            if (num_digits is ExcelMissing)
            {
                return ratio * inputNumber;
            }
            else
            {
                return Math.Round((decimal)(ratio * inputNumber), Convert.ToInt32(num_digits), MidpointRounding.AwayFromZero);
            }
        }
    }

}
