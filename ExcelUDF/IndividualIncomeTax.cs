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


        [ExcelFunction(Category = "个人所得税", Description = "计算年终奖个人所得税，旧版3500起征的。Excel催化剂出品，必属精品！")]
        public static object GS计算年终奖个人所得税3500(
            [ExcelArgument(Description = "年终奖金额")] double yearAward,
            [ExcelArgument(Description = "税前扣除社保等专项附加后的计税收入金额")] double income,
            [ExcelArgument(Description = "起征点，一般为3500，外籍为4800，默认不填为3500")] double baseLine = 3500
         )
        {
            if (baseLine == 0)
            {
                baseLine = 3500;
            }

            List<(decimal Income, decimal Rate, decimal Deduction)> listRateClassInfo = GetRateClassInfos(Properties.Settings.Default.GS3500);

            var taxableIncome = Convert.ToDecimal(yearAward - (baseLine > income ? baseLine - income : 0));
            var rateInfo = listRateClassInfo.LastOrDefault(s => taxableIncome / 12 >= s.Income);
            return taxableIncome * rateInfo.Rate - rateInfo.Deduction;
        }

        [ExcelFunction(Category = "个人所得税", Description = "计算工资个人所得税，旧版3500起征的。Excel催化剂出品，必属精品！")]
        public static object GS计算工资个人所得税3500(
            [ExcelArgument(Description = "税前扣除社保等专项附加后的计税收入金额")] double income,
            [ExcelArgument(Description = "起征点，一般为3500，外籍为4800，默认不填为3500")] double baseLine = 3500
                 )
        {
            if (baseLine == 0)
            {
                baseLine = 3500;
            }
            return CalculateRate(income, baseLine, Properties.Settings.Default.GS3500);

        }



        [ExcelFunction(Category = "个人所得税", Description = "计算工资个人所得税，新版5000起征的。Excel催化剂出品，必属精品！")]
        public static object GS计算工资个人所得税5000(
            [ExcelArgument(Description = "税前扣除社保等专项附加后的计税收入金额")] double income
         )

        {

            double baseLine = 5000;
            return CalculateRate(income, baseLine, Properties.Settings.Default.GS5000);
        }

        [ExcelFunction(Category = "个人所得税", Description = "根据个税金额反算税前收入，新版5000起征的。Excel催化剂出品，必属精品！")]
        public static object GS根据个税金额反算税前收入3500(
                    [ExcelArgument(Description = "个人所得税缴纳金额")] double tax,
            [ExcelArgument(Description = "起征点，一般为3500，外籍为4800，默认不填为3500")] double baseLine = 3500
                 )
        {

            if (baseLine == 0)
            {
                baseLine = 3500;
            }

            return GetIncomeBeforTaxByTax(tax, baseLine, Properties.Settings.Default.GS3500);

        }

        [ExcelFunction(Category = "个人所得税", Description = "根据个税金额反算税前收入，新版5000起征的。Excel催化剂出品，必属精品！")]
        public static object GS根据个税金额反算税前收入5000(
          [ExcelArgument(Description = "个人所得税缴纳金额")] double tax
         )
        {
            return GetIncomeBeforTaxByTax(tax, 5000, Properties.Settings.Default.GS5000);

        }

        [ExcelFunction(Category = "个人所得税", Description = "根据税后收入反算税前收入，旧版3500起征的。Excel催化剂出品，必属精品！")]
        public static object GS根据税后收入反算税前收入3500(
            [ExcelArgument(Description = "税后收入金额")] double incomeAfterTax,
            [ExcelArgument(Description = "起征点，一般为3500，外籍为4800，默认不填为3500")] double baseLine = 3500
         )
        {
            if (baseLine == 0)
            {
                baseLine = 3500;
            }

            return GetIncomeBeforTaxByIncomeAT(incomeAfterTax, baseLine, Properties.Settings.Default.GS3500);

        }

        [ExcelFunction(Category = "个人所得税", Description = "根据税后收入反算税前收入，新版5000起征的。Excel催化剂出品，必属精品！")]
        public static object GS根据税后收入反算税前收入5000(
            [ExcelArgument(Description = "税后收入金额")] double incomeAfterTax
 )
        {
            return GetIncomeBeforTaxByIncomeAT(incomeAfterTax, 5000, Properties.Settings.Default.GS5000);

        }

        private static object GetIncomeBeforTaxByTax(double tax, double baseLine, System.Collections.Specialized.StringCollection rateClassInfoTable)
        {

            var rateClassInfos = GetRateClassInfos(rateClassInfoTable);
            var matchItem = rateClassInfos.Select(s => new { rateClassInfo = s, MaxTax =GetMaxTax(s, rateClassInfos) })
                                         .FirstOrDefault(s =>  Convert.ToDecimal(tax) <= s.MaxTax );
            return (Convert.ToDecimal(tax) + matchItem.rateClassInfo.Deduction) / matchItem.rateClassInfo.Rate + Convert.ToDecimal(baseLine);
        }

        private static decimal GetMaxTax((decimal Income, decimal Rate, decimal Deduction) s, List<(decimal Income, decimal Rate, decimal Deduction)> rateClassInfos)
        {
            //rateClassInfos.FirstOrDefault(t=>t.Income> s.Income).Income 这里取的是下一个减3500后的计税金额
            var nextRateCalssInfo = rateClassInfos.FirstOrDefault(t => t.Income > s.Income);
            //当最后一级时，返回无数据，income是默认值0时
            if (nextRateCalssInfo.Income == 0)
            {
                return 999999999;
            }
            else
            {
                return nextRateCalssInfo.Income * nextRateCalssInfo.Rate - nextRateCalssInfo.Deduction;
            }
        }

        private static object GetIncomeBeforTaxByIncomeAT(double incomeAfterTax, double baseLine, System.Collections.Specialized.StringCollection rateClassInfoTable)
        {
            var rateClassInfos = GetRateClassInfos(rateClassInfoTable);

            var matchItem = rateClassInfos.Select(s => new { rateClassInfo = s, MaxIncomeAT = GetMaxIncomeAT(baseLine, s, rateClassInfos) })
                                         .FirstOrDefault(s => Convert.ToDecimal(incomeAfterTax) <= s.MaxIncomeAT);

            return (Convert.ToDecimal(incomeAfterTax) - matchItem.rateClassInfo.Rate * Convert.ToDecimal(baseLine) - matchItem.rateClassInfo.Deduction) / (1 - matchItem.rateClassInfo.Rate);

        }

        private static decimal GetMaxIncomeAT(double baseLine, (decimal Income, decimal Rate, decimal Deduction) s, List<(decimal Income, decimal Rate, decimal Deduction)> rateClassInfos)
        {
            //rateClassInfos.FirstOrDefault(t=>t.Income> s.Income).Income 这里取的是下一个减3500后的计税金额
            var nextRateCalssInfo = rateClassInfos.FirstOrDefault(t => t.Income > s.Income);
            //当最后一级时，返回无数据，income是默认值0时
            if (nextRateCalssInfo.Income == 0)
            {
                return 999999999;
            }
            else
            {
                return nextRateCalssInfo.Income + Convert.ToDecimal(baseLine) - (nextRateCalssInfo.Income * nextRateCalssInfo.Rate - nextRateCalssInfo.Deduction);
            }

        }

        private static object CalculateRate(double income, double baseLine, System.Collections.Specialized.StringCollection rateClassInfos)
        {
            List<(decimal Income, decimal Rate, decimal Deduction)> listRateClassInfo = GetRateClassInfos(rateClassInfos);
            var taxableIncome = Convert.ToDecimal(income - baseLine);
            var rateInfo = listRateClassInfo.LastOrDefault(s => taxableIncome >= s.Income);
            return taxableIncome * rateInfo.Rate - rateInfo.Deduction;
        }

        private static List<(decimal Income, decimal Rate, decimal Deduction)> GetRateClassInfos(System.Collections.Specialized.StringCollection rateClassInfos)
        {
            List<(decimal Income, decimal Rate, decimal Deduction)> listRateClassInfo = new List<(decimal Income, decimal Rate, decimal Deduction)>();

            foreach (var item in rateClassInfos)
            {
                listRateClassInfo.Add((decimal.Parse(item.Split(',')[0]), decimal.Parse(item.Split(',')[1]), decimal.Parse(item.Split(',')[2])));
            }

            return listRateClassInfo;
        }
    }
}
