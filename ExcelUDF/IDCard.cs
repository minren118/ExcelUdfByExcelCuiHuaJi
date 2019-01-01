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
        private static XElement element = XElement.Load(new StringReader(Resource1.行政区划数据库));
        [ExcelFunction(Category = "身份证相关",IsThreadSafe =true, Description = "输入身份证号码（15位或18位，其他数字返回错误），获取地区信息。Excel催化剂出品，必属精品！")]
        public static object SFZGetAreaInfo(
        [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input,
        [ExcelArgument(Description = "是否拆开横向存放多个单元格数组公式返回结果，默认为false,只返回一个单元格的值")] bool isSplit
                            )
        {
            try
            {
                var idCard = SFZConvert15To18(input);

                if (idCard is ExcelError.ExcelErrorValue)
                {
                    return ExcelError.ExcelErrorValue;
                }
                else
                {
                    string areacode = idCard.ToString().Substring(0, 6);
                    var xNameRow = XName.Get("Row", "urn:schemas-microsoft-com:office:spreadsheet");
                    var xNameData = XName.Get("Data", "urn:schemas-microsoft-com:office:spreadsheet");

                    var row = element.Descendants(xNameRow).FirstOrDefault(s=>s.Descendants(xNameData).FirstOrDefault().Value == areacode);
                    if (row == null)//当找不到时，找市级
                    {
                        areacode = areacode.Substring(0, 4) + "00";
                        row = element.Descendants(xNameRow).FirstOrDefault(s => s.Descendants(xNameData).FirstOrDefault().Value == areacode);
                    }
                    if (row == null)////当找不到时，找省级
                    {
                        areacode = areacode.Substring(0, 2) + "0000";
                        row = element.Descendants(xNameRow).FirstOrDefault(s => s.Descendants(xNameData).FirstOrDefault().Value == areacode);
                    }

                    //返回结果
                    if (row != null)
                    {
                        var areaInfo = row.Descendants(xNameData).Skip(1).FirstOrDefault().Value;

                        if (isSplit)
                        {
                            return Common.ReturnDataArray(areaInfo.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries), "H");
                        }
                        else
                        {
                            return areaInfo;
                        }
                    }
                    else
                    {
                        return "地区信息未能匹配";
                    }
                }
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorValue;
            }
        }

        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "输入身份证号码（15位或18位，其他数字返回错误），获取所属生肖。Excel催化剂出品，必属精品！")]
        public static object SFZGetShengXiao(
                [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input
                    )
        {
            try
            {
                var birthDay = SFZGetBirthday(input);
                if (birthDay is ExcelError.ExcelErrorValue)
                {
                    return ExcelError.ExcelErrorValue;
                }
                else
                {
                    return RiQiGetShengXiao(Convert.ToDateTime(birthDay));
                }

            }
            catch (Exception)
            {

                return ExcelError.ExcelErrorValue;
            }
        }

        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "输入身份证号码（15位或18位，其他数字返回错误），获取所属干支年份。Excel催化剂出品，必属精品！")]
        public static object SFZGetGanZhiYear(
        [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input
            )
        {
            try
            {

                var birthDay = SFZGetBirthday(input);
                if (birthDay is ExcelError.ExcelErrorValue)
                {
                    return ExcelError.ExcelErrorValue;
                }
                else
                {
                    return RiQiGetGanZhiYear(Convert.ToDateTime(birthDay));
                }

            }
            catch (Exception)
            {

                return ExcelError.ExcelErrorValue;
            }
        }


        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "输入身份证号码（15位或18位，其他数字返回错误），获取所属星座。Excel催化剂出品，必属精品！")]
        public static object SFZGetXingZuo(
        [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input
                            )
        {

            try
            {
                var birthDay = SFZGetBirthday(input);
                if (birthDay is ExcelError.ExcelErrorValue)
                {
                    return ExcelError.ExcelErrorValue;
                }

                return RiQiGetXingZuo(Convert.ToDateTime(birthDay));

            }
            catch (Exception)
            {

                return ExcelError.ExcelErrorValue;
            }

        }


        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "输入身份证号码（15位或18位，其他数字返回错误），获取性别。Excel催化剂出品，必属精品！")]
        public static object SFZGetSex(
                [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input
                                    )
        {

            var idCard = SFZConvert15To18(input);

            if (idCard is ExcelError.ExcelErrorValue)
            {
                return ExcelError.ExcelErrorValue;
            }
            else
            {
                string sexStr = idCard.ToString().Substring(16, 1);
                return int.Parse(sexStr) % 2 == 0 ? "女" : "男";

            }
        }

        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "输入身份证号码（15位或18位，其他数字返回错误），获取当前年龄，过了生日才算一年。Excel催化剂出品，必属精品！")]
        public static object SFZGetCurrentAge(
         [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input
                                            )
        {
            var birthDay = SFZGetBirthday(input);
            if (birthDay is ExcelError.ExcelErrorValue)
            {
                return ExcelError.ExcelErrorValue;
            }
            else
            {
                Common.ChangeNumberFormat("0");//生日日期调用变成日期格式，转回来
                return RiQiGetAge(Convert.ToDateTime(birthDay));
            }
        }


        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "输入身份证号码（15位或18位，其他数字返回错误），取出生日期。Excel催化剂出品，必属精品！")]
        public static object SFZGetBirthday(
                 [ExcelArgument(Description = "输入身份证号码，15位或18位，其他数字返回错误")] object input
                                                    )
        {
            var IDCardNo = SFZConvert15To18(input.ToString());
            if (IDCardNo is ExcelError.ExcelErrorValue)
            {
                return ExcelError.ExcelErrorValue;
            }
            else
            {
                string birthDayString = IDCardNo.ToString().Substring(6, 8);
                Common.ChangeNumberFormat("yyyy-mm-dd");
                return DateTime.ParseExact(birthDayString, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
            }
        }



        [ExcelFunction(Category = "身份证相关", IsThreadSafe = true, Description = "15位身份证号转18位，输入非15位或18数字返回错误.Excel催化剂出品，必属精品！")]
        public static object SFZConvert15To18(
           [ExcelArgument(Description = "输入15位的身份证号码，输入非15或18位数字返回错误")] object input)
        {
            string oldIDCard = input.ToString().Trim();

            if (oldIDCard.Length == 15)
            {
                int iS = 0;
                //加权因子常数 
                int[] iW = new int[] { 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2 };
                //校验码常数 
                string LastCode = "10X98765432";
                //新身份证号 
                string newIDCard;

                newIDCard = oldIDCard.Substring(0, 6);
                //填在第6位及第7位上填上‘1’，‘9’两个数字 
                newIDCard += "19";

                newIDCard += oldIDCard.Substring(6, 9);

                //进行加权求和 
                for (int i = 0; i < 17; i++)
                {
                    iS += int.Parse(newIDCard.Substring(i, 1)) * iW[i];
                }

                //取模运算，得到模值 
                int iY = iS % 11;
                //从LastCode中取得以模为索引号的值，加到身份证的最后一位，即为新身份证号。 
                newIDCard += LastCode.Substring(iY, 1);
                return newIDCard;
            }
            else if (oldIDCard.Length == 18 && Int64.TryParse(oldIDCard.Substring(0, 17), out var number))
            {
                return oldIDCard;
            }
            else
            {
                return ExcelError.ExcelErrorValue;
            }


        }

    }
}
