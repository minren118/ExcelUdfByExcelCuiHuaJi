using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "日期相关", Description = "输入公历年份及节气名称返回对应年份下的节气所在日期。Excel催化剂出品，必属精品！")]
        public static object RiQiGetDateByJieQiName(
        [ExcelArgument(Description = "节气对应的公历年")] int JieQiYear,
        [ExcelArgument(Description = "节气名称")] string JieQiName
        )

        {
            var jieQis = Properties.Settings.Default.JieQi;

            foreach (var item in jieQis)
            {
                DateTime jieQiDate = DateTime.FromOADate(double.Parse(item.Split(',')[0]));
                int year = jieQiDate.Year;
                if (year == JieQiYear && item.Split(',')[1] == JieQiName)
                {
                    Common.ChangeNumberFormat("yyyy-mm-dd");
                    return jieQiDate;
                }
            }
            return ExcelError.ExcelErrorValue;
        }

        [ExcelFunction(Category = "日期相关", Description = "输入公历日期，返回对应的节气区间。Excel催化剂出品，必属精品！")]
        public static object RiQiGetJieQiNameByInputDate(
                [ExcelArgument(Description = "输入公历日期，仅限1901至2050年间的日期")] DateTime inputDate
                                                        )
        {
            if (inputDate < new DateTime(1901, 2, 4) || inputDate > new DateTime(2050, 12, 22))
            {
                return ExcelError.ExcelErrorValue;
            }
            var jieQis = Properties.Settings.Default.JieQi;
            var query = from s in jieQis.Cast<string>()
                        let jieQiDate = DateTime.FromOADate(double.Parse(s.Split(',')[0]))
                        let jieQiName = s.Split(',')[1]
                        select new { JieQiName = jieQiName, JieQiDate = jieQiDate };


            var queryEqual = query.FirstOrDefault(s => s.JieQiDate == inputDate);
            if (queryEqual != null)
            {
                return queryEqual.JieQiName;
            }
            else
            {
                var queryGreaterThan = query.SkipWhile(s => s.JieQiDate < inputDate);
                var queryFirst = queryGreaterThan.FirstOrDefault();
                var greaterDate = queryFirst.JieQiDate;
                var greaterName = queryFirst.JieQiName;
                var lowerName = query.TakeWhile(s => s.JieQiDate < greaterDate).LastOrDefault().JieQiName;
                return lowerName + "-" + greaterName;
            }


        }


        [ExcelFunction(Category = "日期相关", Description = "输入公历日期，返回农历。Excel催化剂出品，必属精品！")]
        public static string RiQiGetChineseLunisolarDateFromSolarDate(
           [ExcelArgument(Description = "公历日期")] DateTime solarDate)
        {
            System.Globalization.ChineseLunisolarCalendar cal = new System.Globalization.ChineseLunisolarCalendar();
            int year = cal.GetYear(solarDate);
            int month = cal.GetMonth(solarDate);
            int day = cal.GetDayOfMonth(solarDate);
            int leapMonth = cal.GetLeapMonth(year);
            return string.Format("农历{0}{1}（{2}）年{3}{4}月{5}{6}"
                                , "甲乙丙丁戊己庚辛壬癸"[(year - 4) % 10]
                                , "子丑寅卯辰巳午未申酉戌亥"[(year - 4) % 12]
                                , "鼠牛虎兔龙蛇马羊猴鸡狗猪"[(year - 4) % 12]
                                , month == leapMonth ? "闰" : ""
                                , "无正二三四五六七八九十冬腊"[leapMonth > 0 && leapMonth <= month ? month - 1 : month]
                                , "初十廿三"[day / 10]
                                , "日一二三四五六七八九"[day % 10]
                                );
        }

        [ExcelFunction(Category = "日期相关", Description = "输入中国农历日期，返回公历。Excel催化剂出品，必属精品！")]
        public static DateTime RiQiGetSolarDateFromChineseLunisolarDate(
                             [ExcelArgument(Description = "中国农历日期")] DateTime chineseLunisolarDate,
                              [ExcelArgument(Description = "当月是否是闰月")] bool isRunYue
                             )

        {

            System.Globalization.ChineseLunisolarCalendar cal = new System.Globalization.ChineseLunisolarCalendar();

            int year = chineseLunisolarDate.Year;
            int month = chineseLunisolarDate.Month;
            int day = chineseLunisolarDate.Day;
            int leapMonth = cal.GetLeapMonth(year);
            if (leapMonth > 0)
            {
                if (month > leapMonth - 1)
                {
                    month++;
                }
                else if (month == leapMonth - 1 && isRunYue)
                {
                    month++;
                }
            }

            var solarDate = cal.ToDateTime(year, month, day, 0, 0, 0, 0);
            Common.ChangeNumberFormat("yyyy-mm-dd");
            return solarDate;
        }


        [ExcelFunction(Category = "日期相关", Description = "输入公历日期，返回生肖。Excel催化剂出品，必属精品！")]
        public static string RiQiGetShengXiao(
           [ExcelArgument(Description = "公历日期")] DateTime solarDate)
        {
            System.Globalization.ChineseLunisolarCalendar cal = new System.Globalization.ChineseLunisolarCalendar();
            int year = cal.GetYear(solarDate);
            return "鼠牛虎兔龙蛇马羊猴鸡狗猪"[(year - 4) % 12].ToString();

        }

        [ExcelFunction(Category = "日期相关", Description = "输入公历日期，返回干支年份。Excel催化剂出品，必属精品！")]
        public static string RiQiGetGanZhiYear(
           [ExcelArgument(Description = "公历日期")] DateTime solarDate)
        {
            System.Globalization.ChineseLunisolarCalendar cal = new System.Globalization.ChineseLunisolarCalendar();
            int year = cal.GetYear(solarDate);

            char gan = "甲乙丙丁戊己庚辛壬癸"[(year - 4) % 10];
            char zhi = "子丑寅卯辰巳午未申酉戌亥"[(year - 4) % 12];
            return new string(new char[] { gan, zhi });

        }

        [ExcelFunction(Category = "日期相关", Description = "输入公历日期，返回农历。Excel催化剂出品，必属精品！")]
        public static object RiQiGetXingZuo(
           [ExcelArgument(Description = "公历日期")] DateTime solarDate)
        {
            var result = GetXingZuoInfo();
            foreach (var item in result)
            {
                int year = solarDate.Year;
                DateTime startDate = new DateTime(year, item.startDate.Month, item.startDate.Day);
                DateTime endDate = new DateTime(year, item.endDate.Month, item.endDate.Day);

                if (solarDate >= startDate && solarDate <= endDate)
                {
                    return item.XingZuo;
                }
            }
            return ExcelError.ExcelErrorValue;
        }

        [ExcelFunction(Category = "日期相关", Description = "输入公历日期，返回年龄或工龄，不足一年部分舍去。Excel催化剂出品，必属精品！")]
        public static object RiQiGetAge(
           [ExcelArgument(Description = "公历日期，出生日期或工作开始日期")] DateTime solarDate)
        {
            DateTime now = DateTime.Now;
            int age = now.Year - solarDate.Year;
            if (now.Month < solarDate.Month || (now.Month == solarDate.Month && now.Day < solarDate.Day))
            {
                age--;
            }
            return age;
        }

        private static IEnumerable<(string XingZuo, DateTime startDate, DateTime endDate)> GetXingZuoInfo()
        {
            string source = "白羊座,43180,43210;金牛座,43211,43241;双子座,43242,43272;巨蟹座,43273,43303;狮子座,43304,43335;处女座,43336,43366;天秤座,43367,43396;天蝎座,43397,43426;射手座,43427,43455;魔羯座,43456,43465;水瓶座,43121,43150;双鱼座,43151,43179;魔羯座,43101,43120";
            var sourceSplits = source.Split(';');
            return sourceSplits.Select(s => (s.Split(',')[0], DateTime.FromOADate(double.Parse(s.Split(',')[1])), DateTime.FromOADate(double.Parse(s.Split(',')[2]))));
        }
    }
}
