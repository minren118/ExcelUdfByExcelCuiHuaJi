using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他交集区域，求对区域的值求交集返回数据集。Excel催化剂出品，必属精品！")]
        public static object RangeValuesDistinct(
         [ExcelArgument(Description = "源区域")] object srcRange,
         [ExcelArgument(Description = "是否保留空值，默认是不保留的，保留输入True，不保留输入False")] bool optIsRetainNull,
         [ExcelArgument(Description = "返回结果按行还是按列返回，默认不填按列返回，传入H按行返回")] string optAlignHorL)
        {

            var arr = RangeValuesDistinctArray(srcRange, optIsRetainNull);
            return Common.ReturnDataArray(arr, optAlignHorL);
        }

        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他交集区域，求对区域的值求交集返回数据集。Excel催化剂出品，必属精品！")]
        public static object[] RangeValuesDistinctArray(
         [ExcelArgument(Description = "源区域")] object srcRange,
         [ExcelArgument(Description = "是否保留空值，默认是不保留的，保留输入True，不保留输入False")] bool optIsRetainNull
            )
        {
            List<object> listSrcCells = new List<object>();
            Common.AddValueToList(srcRange, ref listSrcCells);

            if (optIsRetainNull)
            {
                return listSrcCells.Select(s=>s is ExcelEmpty?"":s).Distinct().ToArray();
            }
            else
            {
                return listSrcCells.Where(s => s != ExcelEmpty.Value).Distinct().ToArray();
            }
        }





        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他交集区域，求对区域的值求交集返回数据集。Excel催化剂出品，必属精品！")]
        public static object RangeIntersectValue(
           [ExcelArgument(Description = "返回结果按行还是按列返回，默认不填按列返回，传入H按行返回")] string optAlignHorL,
           [ExcelArgument(Description = "返回结果集是否去重复项，默认不输入为否不去重处理，True为去重，False为不去重")] bool optIsDistinctResult,
           [ExcelArgument(Description = "源区域")] object srcRange,
           [ExcelArgument(Description = "和源区域取交集区域1")] object intersectRange1,
           [ExcelArgument(Description = "和源区域取交集区域2")] object intersectRange2,
           [ExcelArgument(Description = "和源区域取交集区域3")] object intersectRange3,
           [ExcelArgument(Description = "和源区域取交集区域4")] object intersectRange4,
           [ExcelArgument(Description = "和源区域取交集区域5")] object intersectRange5,
           [ExcelArgument(Description = "和源区域取交集区域6")] object intersectRange6,
           [ExcelArgument(Description = "和源区域取交集区域7")] object intersectRange7,
           [ExcelArgument(Description = "和源区域取交集区域8")] object intersectRange8)
        {
            var arr = RangeIntersectValueArray(optIsDistinctResult, srcRange, intersectRange1, intersectRange2, intersectRange3, intersectRange4, intersectRange5, intersectRange6, intersectRange7, intersectRange8);

            return Common.ReturnDataArray(arr, optAlignHorL);
        }


        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他交集区域，求对区域的值求交集返回数据集。Excel催化剂出品，必属精品！")]
        public static object[] RangeIntersectValueArray(
               [ExcelArgument(Description = "返回结果集是否去重复项，默认不输入为否不去重处理，True为去重，False为不去重")] bool optIsDistinctResult,
               [ExcelArgument(Description = "源区域")] object srcRange,
               [ExcelArgument(Description = "和源区域取交集区域1")] object intersectRange1,
               [ExcelArgument(Description = "和源区域取交集区域2")] object intersectRange2,
               [ExcelArgument(Description = "和源区域取交集区域3")] object intersectRange3,
               [ExcelArgument(Description = "和源区域取交集区域4")] object intersectRange4,
               [ExcelArgument(Description = "和源区域取交集区域5")] object intersectRange5,
               [ExcelArgument(Description = "和源区域取交集区域6")] object intersectRange6,
               [ExcelArgument(Description = "和源区域取交集区域7")] object intersectRange7,
               [ExcelArgument(Description = "和源区域取交集区域8")] object intersectRange8)
        {

            List<object> listIntersectRange = new List<object>();
            List<object> listSrcRange = new List<object>();

            if (srcRange is ExcelMissing)
            {
                throw new Exception();
            }
            else
            {
                Common.AddValueToList(srcRange, ref listSrcRange);
            }

            AddValueToListOfCompareValues(intersectRange1, intersectRange2, intersectRange3, intersectRange4, intersectRange5, intersectRange6, intersectRange7, intersectRange8, ref listIntersectRange);

            object[] result;
            if (optIsDistinctResult == true)
            {
                result = listSrcRange.Intersect(listIntersectRange).ToArray();
            }
            else
            {
                result = listSrcRange.Where(s => listIntersectRange.Contains(s)).ToArray();
            }

            return result;

        }
        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他并集区域，求对区域的值求并集返回数据集。Excel催化剂出品，必属精品！")]
        public static object RangeUnionValue(
         [ExcelArgument(Description = "返回结果按行还是按列返回，默认不填按列返回，传入H按行返回")] string optAlignHorL,
         [ExcelArgument(Description = "返回结果集是否去重复项，默认不输入为否不去重处理，True为去重，False为不去重")] bool optIsDistinctResult,
         [ExcelArgument(Description = "源区域")] object srcRange,
        [ExcelArgument(Description = "和源区域取并集区域1")] object unionRange1,
        [ExcelArgument(Description = "和源区域取并集区域2")] object unionRange2,
        [ExcelArgument(Description = "和源区域取并集区域3")] object unionRange3,
        [ExcelArgument(Description = "和源区域取并集区域4")] object unionRange4,
        [ExcelArgument(Description = "和源区域取并集区域5")] object unionRange5,
        [ExcelArgument(Description = "和源区域取并集区域6")] object unionRange6,
        [ExcelArgument(Description = "和源区域取并集区域7")] object unionRange7,
        [ExcelArgument(Description = "和源区域取并集区域8")] object unionRange8)
        {
            var result = RangeUnionValueArray(optAlignHorL, optIsDistinctResult, srcRange, unionRange1, unionRange2, unionRange3, unionRange4, unionRange5, unionRange6, unionRange7, unionRange8);
            return Common.ReturnDataArray(result, optAlignHorL);
        }


        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他并集区域，求对区域的值求并集返回数据集。Excel催化剂出品，必属精品！")]
        public static object[] RangeUnionValueArray(
         [ExcelArgument(Description = "返回结果按行还是按列返回，默认不填按列返回，传入H按行返回")] string optAlignHorL,
         [ExcelArgument(Description = "返回结果集是否去重复项，默认不输入为否不去重处理，True为去重，False为不去重")] bool optIsDistinctResult,
         [ExcelArgument(Description = "源区域")] object srcRange,
        [ExcelArgument(Description = "和源区域取并集区域1")] object unionRange1,
        [ExcelArgument(Description = "和源区域取并集区域2")] object unionRange2,
        [ExcelArgument(Description = "和源区域取并集区域3")] object unionRange3,
        [ExcelArgument(Description = "和源区域取并集区域4")] object unionRange4,
        [ExcelArgument(Description = "和源区域取并集区域5")] object unionRange5,
        [ExcelArgument(Description = "和源区域取并集区域6")] object unionRange6,
        [ExcelArgument(Description = "和源区域取并集区域7")] object unionRange7,
        [ExcelArgument(Description = "和源区域取并集区域8")] object unionRange8)
        {
            List<object> listUnionRange = new List<object>();
            List<object> listSrcRange = new List<object>();

            if (srcRange is ExcelMissing)
            {
                throw new Exception();
            }
            else
            {
                Common.AddValueToList(srcRange, ref listSrcRange);
            }

            AddValueToListOfCompareValues(unionRange1, unionRange2, unionRange3, unionRange4, unionRange5, unionRange6, unionRange7, unionRange8, ref listUnionRange);

            if (optIsDistinctResult == true)
            {
                return listSrcRange.Union(listUnionRange).ToArray();
            }
            else
            {
                return listSrcRange.Concat(listUnionRange).ToArray();
            }
        }


        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他补集区域，求对区域的值求补集返回数据集。Excel催化剂出品，必属精品！")]
        public static object RangeExceptValue(
           [ExcelArgument(Description = "返回结果按行还是按列返回，默认不填按列返回，传入H按行返回")] string optAlignHorL,
           [ExcelArgument(Description = "返回结果集是否去重复项，默认不输入为否不去重处理，True为去重，False为不去重")] bool optIsDistinctResult,
         [ExcelArgument(Description = "源区域")] object srcRange,
        [ExcelArgument(Description = "和源区域取补集区域1")] object exceptRange1,
        [ExcelArgument(Description = "和源区域取补集区域2")] object exceptRange2,
        [ExcelArgument(Description = "和源区域取补集区域3")] object exceptRange3,
        [ExcelArgument(Description = "和源区域取补集区域4")] object exceptRange4,
        [ExcelArgument(Description = "和源区域取补集区域5")] object exceptRange5,
        [ExcelArgument(Description = "和源区域取补集区域6")] object exceptRange6,
        [ExcelArgument(Description = "和源区域取补集区域7")] object exceptRange7,
        [ExcelArgument(Description = "和源区域取补集区域8")] object exceptRange8)
        {

            var result = RangeExceptValueArray(optIsDistinctResult, srcRange, exceptRange1, exceptRange2, exceptRange3, exceptRange4, exceptRange5, exceptRange6, exceptRange7, exceptRange8);

            return Common.ReturnDataArray(result, optAlignHorL);
        }

        [ExcelFunction(Category = "区域集合处理", Description = "区域值根据源区域和其他补集区域，求对区域的值求补集返回数据集。Excel催化剂出品，必属精品！")]
        public static object[] RangeExceptValueArray(
         [ExcelArgument(Description = "返回结果集是否去重复项，默认不输入为否不去重处理，True为去重，False为不去重")] bool optIsDistinctResult,
         [ExcelArgument(Description = "源区域")] object srcRange,
        [ExcelArgument(Description = "和源区域取补集区域1")] object exceptRange1,
        [ExcelArgument(Description = "和源区域取补集区域2")] object exceptRange2,
        [ExcelArgument(Description = "和源区域取补集区域3")] object exceptRange3,
        [ExcelArgument(Description = "和源区域取补集区域4")] object exceptRange4,
        [ExcelArgument(Description = "和源区域取补集区域5")] object exceptRange5,
        [ExcelArgument(Description = "和源区域取补集区域6")] object exceptRange6,
        [ExcelArgument(Description = "和源区域取补集区域7")] object exceptRange7,
        [ExcelArgument(Description = "和源区域取补集区域8")] object exceptRange8)
        {
            List<object> listExceptRange = new List<object>();
            List<object> listSrcRange = new List<object>();

            if (srcRange is ExcelMissing)
            {
                throw new Exception();
            }
            else
            {
                Common.AddValueToList(srcRange, ref listSrcRange);
            }

            AddValueToListOfCompareValues(exceptRange1, exceptRange2, exceptRange3, exceptRange4, exceptRange5, exceptRange6, exceptRange7, exceptRange8, ref listExceptRange);

            if (optIsDistinctResult == true)
            {
                return listSrcRange.Except(listExceptRange).ToArray();
            }
            else
            {
                return listSrcRange.Where(s => !listExceptRange.Contains(s)).ToArray();
            }

        }
        private static void AddValueToListOfCompareValues(object intersectRange1, object intersectRange2, object intersectRange3, object intersectRange4, object intersectRange5, object intersectRange6, object intersectRange7, object intersectRange8, ref List<object> listIntersectRange)
        {
            if (!(intersectRange1 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange1, ref listIntersectRange);
            }
            if (!(intersectRange2 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange2, ref listIntersectRange);
            }
            if (!(intersectRange3 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange3, ref listIntersectRange);
            }
            if (!(intersectRange4 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange4, ref listIntersectRange);
            }
            if (!(intersectRange5 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange5, ref listIntersectRange);
            }
            if (!(intersectRange6 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange6, ref listIntersectRange);
            }
            if (!(intersectRange7 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange7, ref listIntersectRange);
            }
            if (!(intersectRange8 is ExcelMissing))
            {
                Common.AddValueToList(intersectRange8, ref listIntersectRange);
            }
        }



    }
}
