using ExcelDna.Integration;
using Microsoft.International.Converters.PinYinConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "规划求解类", Description = "分组凑数，从源数据列中，抽取出指定的项目组合，使其求和数最大限度接近分组的大小。Excel催化剂出品，必属精品！")]
        public static object CouShuWithGroupFromOrTools(
                                                   [ExcelArgument(Description = "需要分组的原始数据单元格区域，精度为最多4位小数点，多于4位将截断")] object[] srcRange,
                                                   [ExcelArgument(Description = "限定组的上限的单元格区域，可选多个单元格代表分多个组，组的大小可不相同，尽量较难组合的放最上面优先对其组合")] object[] groupeRange
                                                    )


        {

            int scaleNum = GetScaleNum(srcRange);

            KnapsacksService.KnapsacksServiceSoapClient client = new KnapsacksService.KnapsacksServiceSoapClient();


            KnapsacksService.ArrayOfLong values = new KnapsacksService.ArrayOfLong();

            values.AddRange(srcRange.Select(s => Convert.ToDouble(s)).Select(t => Convert.ToInt64(t * Math.Pow(10, scaleNum))));

            KnapsacksService.ArrayOfLong capacities = new KnapsacksService.ArrayOfLong();
            capacities.AddRange(groupeRange.Where(s => s != ExcelEmpty.Value).Select(t => Convert.ToDouble(t)).Select(r => Convert.ToInt64(r * Math.Pow(10, scaleNum))));

            KnapsacksService.ArrayOfAnyType results = client.GetGroupIdsByKnapsacks(values, capacities,scaleNum);

            return Common.ReturnDataArray(results.Select(s => s).ToArray(), "L");

        }
        [ExcelFunction(Category = "规划求解类", Description = "分组凑数，从源数据列中，抽取出指定的项目组合，使其求和数最大限度接近分组的大小。大量参考EH香川群子代码，Excel催化剂出品，必属精品！")]
        public static object CouShuWithGroupFromEH(
                                           [ExcelArgument(Description = "需要分组的原始数据单元格区域，精度为最多4位小数点，多于4位将截断")] object[] srcRange,
                                           [ExcelArgument(Description = "限定组的上限的单元格区域，可选多个单元格代表分多个组，组的大小可不相同，尽量较难组合的放最上面优先对其组合")] object[] groupeRange
                                            )
        {
            int scaleNum = GetScaleNum(srcRange);

            object[] values = srcRange.Select(s => Convert.ToDouble(s)).Select(t => Convert.ToInt64(t * Math.Pow(10, scaleNum))).Select(z => (object)z).ToArray();
            object[] capacities = groupeRange.Where(s => s != ExcelEmpty.Value).Select(t => Convert.ToDouble(t)).Select(r => Convert.ToInt64(r * Math.Pow(10, scaleNum))).Select(z=>(object)z).ToArray();
            var result = Common.ExcelApp.Run("GetGroupIds", values, capacities);

            return Common.ReturnDataArray(result, "L"); 
        }


        private static int GetScaleNum(object[] srcRange)
        {
            int scaleNum = 0;
            bool isNotInt = srcRange.Select(s => s.ToString()).Any(t => !IsInteger(t));
            if (isNotInt)
            {
                if (srcRange.Select(s => s.ToString()).Any(t => !IsNumber(t, 32, 2)))
                {
                    scaleNum = 4;
                }
                else
                {
                    scaleNum = 2;
                }
            }

            return scaleNum;
        }

        private static bool IsInteger(string s)
        {
            string pattern = @"^\d*$";
            return Regex.IsMatch(s, pattern);
        }
        /// <summary>
        /// 判断一个字符串是否为合法数字(0-32整数)
        /// </summary>
        /// <param name="s">字符串</param>
        /// <returns></returns>
        private static bool IsNumber(string s)
        {
            return IsNumber(s, 32, 0);
        }
        /// <summary>
        /// 判断一个字符串是否为合法数字(指定整数位数和小数位数)
        /// </summary>
        /// <param name="s">字符串</param>
        /// <param name="precision">整数位数</param>
        /// <param name="scale">小数位数</param>
        /// <returns></returns>
        private static bool IsNumber(string s, int precision, int scale)
        {
            if ((precision == 0) && (scale == 0))
            {
                return false;
            }
            string pattern = @"(^\d{1," + precision + "}";
            if (scale > 0)
            {
                pattern += @"\.\d{0," + scale + "}$)|"+ pattern;//|或者条件，它用了一次的pattern，小数点位数是小于等于的作用
            }
            pattern += "$)";
            return Regex.IsMatch(s, pattern);
        }

    }
}
