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

        [ExcelFunction(Category = "分组计算", Description = "分组最小值，实现的效果类似MINIF函数，但效率性能更高。Excel催化剂出品，必属精品！")]
        public static object FZJS分组最小值(
                  [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
                  [ExcelArgument(Description = "求最小值列区域")] object[,] minRange
                )
        {
            object[,] rankRange2 = new object[1, 1];
            object[,] rankRange3 = new object[1, 1];

            List<(int Index, object GrpFiled, object CalField)> queryList = GetSrcData(groupRange, minRange);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, double> result = new Dictionary<int, double>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    result.Add(item.Index, grp.Where(s => s.CalField is double).Min(s => Convert.ToDouble(s.CalField)));
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        [ExcelFunction(Category = "分组计算", Description = "分组最大值，实现的效果类似MAXIF函数，但效率性能更高。Excel催化剂出品，必属精品！")]
        public static object FZJS分组最大值(
          [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
          [ExcelArgument(Description = "求最大值列区域")] object[,] maxRange
        )

        {
            object[,] rankRange2 = new object[1, 1];
            object[,] rankRange3 = new object[1, 1];

            List<(int Index, object GrpFiled, object CalField)> queryList = GetSrcData(groupRange, maxRange);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, double> result = new Dictionary<int, double>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    result.Add(item.Index, grp.Where(s => s.CalField is double).Max(s => Convert.ToDouble(s.CalField)));
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        [ExcelFunction(Category = "分组计算", Description = "分组平均值，实现的效果类似AVERAGEIF函数，但效率性能更高。Excel催化剂出品，必属精品！")]
        public static object FZJS分组平均值(
            [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
            [ExcelArgument(Description = "求平均值列区域")] object[,] avgRange
          )

        {
            object[,] rankRange2 = new object[1, 1];
            object[,] rankRange3 = new object[1, 1];

            List<(int Index, object GrpFiled, object CalField)> queryList = GetSrcData(groupRange, avgRange);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, double> result = new Dictionary<int, double>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    result.Add(item.Index, grp.Where(s => s.CalField is double).Average(s => Convert.ToDouble(s.CalField)));
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }

        [ExcelFunction(Category = "分组计算", Description = "分组求和，实现的效果类似SUMIF函数，但效率性能更高。Excel催化剂出品，必属精品！")]
        public static object FZJS分组求和(
      [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object groupRange,
      [ExcelArgument(Description = "求和列区域")] object[,] sumRange
                  )
        {
            object[,] rankRange2 = new object[1, 1];
            object[,] rankRange3 = new object[1, 1];

            List<(int Index, object GrpFiled, object CalField)> queryList = GetSrcData(groupRange, sumRange);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, double> result = new Dictionary<int, double>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    result.Add(item.Index, grp.Where(s => s.CalField is double).Sum(s => Convert.ToDouble(s.CalField)));
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        [ExcelFunction(Category = "分组计算", Description = "分组美式排名。Excel催化剂出品，必属精品！")]
        public static object FZJS分组美式排名(
         [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object groupRange,
         [ExcelArgument(Description = "排名列区域")] object[,] rankRange,
         [ExcelArgument(Description = "排名列区域是否按降序排列，默认FALSE为从大到小排名，TRUE时为从小到大排名")] bool isRankAsc
                     )
        {
            object[,] rankRange2 = new object[1, 1];
            object[,] rankRange3 = new object[1, 1];

            List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, rankRange, rankRange2, rankRange3);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, int> result = new Dictionary<int, int>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    if (rankRange.GetLength(0) != 1)
                    {
                        if (isRankAsc)
                        {
                            int rankIndex = grp.Where(s => s.OrderField1 < item.OrderField1).Count() + 1;
                            result.Add(item.Index, rankIndex);
                        }
                        else
                        {
                            int rankIndex = grp.Where(s => s.OrderField1 > item.OrderField1).Count() + 1;
                            result.Add(item.Index, rankIndex);
                        }

                    }
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }

        [ExcelFunction(Category = "分组计算", Description = "分组中式排名。Excel催化剂出品，必属精品！")]
        public static object FZJS分组中式排名(
                  [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object groupRange,
                  [ExcelArgument(Description = "排名列区域")] object[,] rankRange,
                  [ExcelArgument(Description = "排名列区域是否按降序排列，默认FALSE为从大到小排名，TRUE时为从小到大排名")] bool isRankAsc
              )
        {
            object[,] rankRange2 = new object[1, 1];
            object[,] rankRange3 = new object[1, 1];

            List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, rankRange, rankRange2, rankRange3);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, int> result = new Dictionary<int, int>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    if (rankRange.GetLength(0) != 1)
                    {
                        if (isRankAsc)
                        {
                            int rankIndex = grp.Where(s => s.OrderField1 < item.OrderField1).Select(s => s.OrderField1).Distinct().Count() + 1;
                            result.Add(item.Index, rankIndex);
                        }
                        else
                        {
                            int rankIndex = grp.Where(s => s.OrderField1 > item.OrderField1).Select(s => s.OrderField1).Distinct().Count() + 1;
                            result.Add(item.Index, rankIndex);
                        }

                    }
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        [ExcelFunction(Category = "分组计算", Description = "分组求和累计。Excel催化剂出品，必属精品！")]
        public static object FZJS分组求和累计(
            [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
            [ExcelArgument(Description = "求和区域")] object[,] sumRange,
            [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
            [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
            [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
            [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
            [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
            [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
             )
        {

            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, sumRange, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select((s, index) => new { Item = s, Index = index }))
                {
                    int itemIndex = item.Index;
                    var sumValue = grp.Where((s, grpIndx) => s.ReturnField is double && grpIndx <= itemIndex).Sum(s => Convert.ToDouble(s.ReturnField));
                    result.Add(item.Item.Index, sumValue);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");
        }



        [ExcelFunction(Category = "分组计算", Description = "分组求末元素。Excel催化剂出品，必属精品！")]
        public static object FZJS分组求末元素(
            [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
            [ExcelArgument(Description = "返回值区域")] object[,] returnRange,
            [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
            [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
            [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
            [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
            [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
            [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
            )
        {

            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, returnRange, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select((s, index) => new { Item = s, Index = index }))
                {
                    int itemIndex = item.Index;
                    object lagValue = grp.Reverse().Select(s => s.ReturnField).FirstOrDefault();
                    result.Add(item.Item.Index, lagValue == null ? "" : lagValue);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");
        }



        [ExcelFunction(Category = "分组计算", Description = "分组求首元素。Excel催化剂出品，必属精品！")]
        public static object FZJS分组求首元素(
      [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
      [ExcelArgument(Description = "返回值区域")] object[,] returnRange,
      [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
      [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
      [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
      [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
      [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
      [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
                  )
        {

            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, returnRange, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select((s, index) => new { Item = s, Index = index }))
                {
                    int itemIndex = item.Index;
                    object lagValue = grp.Select(s => s.ReturnField).FirstOrDefault();
                    result.Add(item.Item.Index, lagValue == null ? "" : lagValue);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");
        }


        [ExcelFunction(Category = "分组计算", Description = "分组求下一元素。Excel催化剂出品，必属精品！")]
        public static object FZJS分组求下一元素(
              [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
              [ExcelArgument(Description = "返回值区域")] object[,] returnRange,
              [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
              [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
              [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
              [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
              [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
              [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
                          )
        {

            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, returnRange, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select((s, index) => new { Item = s, Index = index }))
                {
                    int itemIndex = item.Index;
                    object lagValue = grp.Where((s, index) => index == itemIndex + 1).Select(s => s.ReturnField).FirstOrDefault();
                    result.Add(item.Item.Index, lagValue == null ? "" : lagValue);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");
        }



        [ExcelFunction(Category = "分组计算", Description = "分组求上一元素。Excel催化剂出品，必属精品！")]
        public static object FZJS分组求上一元素(
              [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
              [ExcelArgument(Description = "返回值区域")] object[,] returnRange,
              [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
              [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
              [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
              [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
              [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
              [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
                          )
        {

            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, returnRange, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select((s, index) => new { Item = s, Index = index }))
                {
                    int itemIndex = item.Index;
                    object lagValue = grp.Where((s, index) => index == itemIndex - 1).Select(s => s.ReturnField).FirstOrDefault();
                    result.Add(item.Item.Index, lagValue == null ? "" : lagValue);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");
        }

        [ExcelFunction(Category = "分组计算", Description = "分组计数去重，实现的效果类似COUNTIF函固定首行绝对引用，但效率性能更高。Excel催化剂出品，必属精品！")]
        public static object FZJS分组计数去重(
                [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
                [ExcelArgument(Description = "去重统计计数列区域")] object[,] distinctCountRange
                            )
        {
            var queryList = GetSrcData(groupRange, distinctCountRange);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    result.Add(item.Index, grp.Where(s => s.calculateField != ExcelEmpty.Value).Select(s => s.calculateField).Distinct().Count());
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");

        }




        [ExcelFunction(Category = "分组计算", Description = "分组计数，实现的效果类似COUNTIF函数，但效率性能更高。Excel催化剂出品，必属精品！")]
        public static object FZJS分组计数(
                    [ExcelArgument(Description = "分组列区域1，仅能选取一列")] object[,] groupRange1,
                    [ExcelArgument(Description = "分组列区域2，仅能选取一列")] object[,] groupRange2,
                    [ExcelArgument(Description = "分组列区域3，仅能选取一列")] object[,] groupRange3,
                    [ExcelArgument(Description = "分组列区域4，仅能选取一列")] object[,] groupRange4
                                )
        {
            object[,] groupArrs = (object[,])FZJS分组列合并(groupRange1, groupRange2, groupRange3, groupRange4);
            int arrDim0Length = groupArrs.GetLength(0);
            object[] srcGrpDatas = new object[arrDim0Length];
            for (int i = 0; i < arrDim0Length; i++)
            {
                srcGrpDatas[i] = groupArrs[i, 0];
            }

            var grps = srcGrpDatas.Select((s, index) => new { Index = index, Value = s }).GroupBy(s => s.Value);
            Dictionary<int, int> result = new Dictionary<int, int>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select(s => s))
                {
                    result.Add(item.Index, grp.Count());
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }

        [ExcelFunction(Category = "分组计算", Description = "分组字符拼接，生成组内字符拼接列的字符拼接成一个字符串，Excel催化剂出品，必属精品！")]
        public static object FZJS分组字符拼接(
                [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object[,] groupRange,
                [ExcelArgument(Description = "拼接字符串的列区域")] object[,] joinString,
                [ExcelArgument(Description = "拼接字符串的分隔符")] object splitString,
                [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
                [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
                [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
                [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
                [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
                [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
                            )
        {
            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, joinString, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            string splitStr = string.Empty;
            if (splitString !=ExcelMissing.Value)
            {
                splitStr = splitString.ToString();
            }


            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, object> result = new Dictionary<int, object>();
            foreach (var grp in grps)
            {
                foreach (var item in grp)
                {
                    var joinStrings = grp.Select(t => t.ReturnField==null?"": t.ReturnField);
                    result.Add(item.Index, string.Join(splitStr, joinStrings));
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => s.Value).ToArray(), "L");
        }

        [ExcelFunction(Category = "分组计算", Description = "分组序号，生成组内不重复的递增序号，Excel催化剂出品，必属精品！")]
        public static object FZJS分组序号(
                 [ExcelArgument(Description = "分组列区域，当有多列作为分组条件时，需使用【FZGetMultiColRange】函数输入")] object groupRange,
                 [ExcelArgument(Description = "排序列区域1")] object[,] orderRange1,
                 [ExcelArgument(Description = "排序列区域1是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange1Desc,
                 [ExcelArgument(Description = "排序列区域2")] object[,] orderRange2,
                 [ExcelArgument(Description = "排序列区域2是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange2Desc,
                 [ExcelArgument(Description = "排序列区域3")] object[,] orderRange3,
                 [ExcelArgument(Description = "排序列区域3是否按降序排列，默认为升序FALSE，TRUE时为降序")] bool isorderRange3Desc
                             )
        {
            List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList = GetSrcData(groupRange, orderRange1, orderRange2, orderRange3);
            queryList = GetOrderList(queryList, orderRange1, isorderRange1Desc, orderRange2, isorderRange2Desc, orderRange3, isorderRange3Desc);

            var grps = queryList.GroupBy(s => s.GrpFiled);
            Dictionary<int, int> result = new Dictionary<int, int>();
            foreach (var grp in grps)
            {
                foreach (var item in grp.Select((s, index) => new { Item = s, Index = index }))
                {
                    result.Add(item.Item.Index, item.Index + 1);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        [ExcelFunction(Category = "分组计算", Description = "分组列合并，用于分组列有多列时，需要使用此函数来引用多列。Excel催化剂出品，必属精品！")]
        public static object FZJS分组列合并(
            [ExcelArgument(Description = "分组列区域1，仅能选取一列")] object[,] groupRange1,
            [ExcelArgument(Description = "分组列区域2，仅能选取一列")] object[,] groupRange2,
            [ExcelArgument(Description = "分组列区域3，仅能选取一列")] object[,] groupRange3,
            [ExcelArgument(Description = "分组列区域4，仅能选取一列")] object[,] groupRange4
            )
        {

            //当不止选了一列时，返回错误
            if (groupRange1.GetLength(1) != 1 || groupRange2.GetLength(1) != 1 || groupRange3.GetLength(1) != 1 || groupRange4.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0Length1 = groupRange1.GetLength(0);
            int arrDim0Length2 = groupRange2.GetLength(0);
            int arrDim0Length3 = groupRange3.GetLength(0);
            int arrDim0Length4 = groupRange4.GetLength(0);

            //当行的数量不同时，也不能计算，返回出错
            if ((arrDim0Length1 != arrDim0Length2 && arrDim0Length2 != 1) || (arrDim0Length1 != arrDim0Length3 && arrDim0Length3 != 1) || (arrDim0Length1 != arrDim0Length4 && arrDim0Length4 != 1))
            {
                throw new ArgumentException("参数出错");
            }

            List<string> result = new List<string>();

            for (int i = 0; i < arrDim0Length1; i++)
            {
                StringBuilder sb = new StringBuilder();

                sb.Append(groupRange1[i, 0].ToString() + "|");

                if (arrDim0Length2 != 1)
                {
                    sb.Append(groupRange2[i, 0].ToString() + "|");
                }
                if (arrDim0Length3 != 1)
                {
                    sb.Append(groupRange3[i, 0].ToString() + "|");
                }
                if (arrDim0Length4 != 1)
                {
                    sb.Append(groupRange4[i, 0].ToString() + "|");
                }
                result.Add(sb.ToString().TrimEnd('|'));

            }
            object[,] arrResult = new object[result.Count, 1];
            for (int i = 0; i < result.Count; i++)
            {
                arrResult[i, 0] = result[i];
            }

            return arrResult;
        }

        private static List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> GetOrderList(List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList, object[,] orderRange1, bool isorderRange1Desc, object[,] orderRange2, bool isorderRange2Desc, object[,] orderRange3, bool isorderRange3Desc)
        {
            if (orderRange1.GetLength(0) != 1 && orderRange2.GetLength(0) == 1 && orderRange3.GetLength(0) == 1)//一列排序列
            {
                if (isorderRange1Desc)
                {
                    return queryList.OrderByDescending(s => s.OrderField1).ToList();
                }
                else
                {
                    return queryList.OrderBy(s => s.OrderField1).ToList();
                }
            }
            else if (orderRange1.GetLength(0) != 1 && orderRange2.GetLength(0) != 1 && orderRange3.GetLength(0) == 1)//两列排序列
            {
                if (isorderRange1Desc)
                {
                    if (isorderRange2Desc)
                    {
                        return queryList.OrderByDescending(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ToList();
                    }
                    else
                    {
                        return queryList.OrderByDescending(s => s.OrderField1).ThenBy(s => s.OrderField2).ToList();
                    }
                }
                else
                {
                    if (isorderRange2Desc)
                    {
                        return queryList.OrderBy(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ToList();
                    }
                    else
                    {
                        return queryList.OrderBy(s => s.OrderField1).ThenBy(s => s.OrderField2).ToList();
                    }
                }
            }

            else if (orderRange1.GetLength(0) != 1 && orderRange2.GetLength(0) != 1 && orderRange3.GetLength(0) != 1)//3列排序列
            {
                if (isorderRange1Desc)
                {
                    if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }

                    }
                    else//if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }
                    }
                }
                else//if (isorderRange1Desc)
                {
                    if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }
                    }
                    else//if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }
                    }
                }
            }

            return new List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)>();
        }

        private static List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)> GetOrderList(List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)> queryList, object[,] orderRange1, bool isorderRange1Desc, object[,] orderRange2, bool isorderRange2Desc, object[,] orderRange3, bool isorderRange3Desc)
        {
            if (orderRange1.GetLength(0) != 1 && orderRange2.GetLength(0) == 1 && orderRange3.GetLength(0) == 1)//一列排序列
            {
                if (isorderRange1Desc)
                {
                    return queryList.OrderByDescending(s => s.OrderField1).ToList();
                }
                else
                {
                    return queryList.OrderBy(s => s.OrderField1).ToList();
                }
            }
            else if (orderRange1.GetLength(0) != 1 && orderRange2.GetLength(0) != 1 && orderRange3.GetLength(0) == 1)//两列排序列
            {
                if (isorderRange1Desc)
                {
                    if (isorderRange2Desc)
                    {
                        return queryList.OrderByDescending(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ToList();
                    }
                    else
                    {
                        return queryList.OrderByDescending(s => s.OrderField1).ThenBy(s => s.OrderField2).ToList();
                    }
                }
                else
                {
                    if (isorderRange2Desc)
                    {
                        return queryList.OrderBy(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ToList();
                    }
                    else
                    {
                        return queryList.OrderBy(s => s.OrderField1).ThenBy(s => s.OrderField2).ToList();
                    }
                }
            }

            else if (orderRange1.GetLength(0) != 1 && orderRange2.GetLength(0) != 1 && orderRange3.GetLength(0) != 1)//3列排序列
            {
                if (isorderRange1Desc)
                {
                    if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }

                    }
                    else//if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderByDescending(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }
                    }
                }
                else//if (isorderRange1Desc)
                {
                    if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenByDescending(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }
                    }
                    else//if (isorderRange2Desc)
                    {
                        if (isorderRange3Desc)
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenByDescending(s => s.OrderField3).ToList();
                        }
                        else
                        {
                            return queryList.OrderBy(s => s.OrderField1).ThenBy(s => s.OrderField2).ThenBy(s => s.OrderField3).ToList();
                        }
                    }
                }
            }

            return new List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)>();
        }


        private static List<(int Index, object GrpFiled, object calculateField)> GetSrcData(object groupRange, object[,] calculateRange)
        {

            if (calculateRange.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0groupRange = 0;
            List<object> grpFields = new List<object>();
            if (groupRange is object[])
            {
                var arrGrp = groupRange as object[];
                arrDim0groupRange = arrGrp.GetLength(0);
                grpFields.AddRange(arrGrp);
            }
            else if (groupRange is object[,])
            {
                var arrGrp = groupRange as object[,];
                arrDim0groupRange = (groupRange as object[,]).GetLength(0);
                for (int i = 0; i < arrGrp.GetLength(0); i++)
                {
                    grpFields.Add(arrGrp[i, 0]);
                }
            }

            int arrDim0CalRange = calculateRange.GetLength(0);



            if (arrDim0groupRange != arrDim0CalRange && arrDim0groupRange != 1)//分组列可以为空，但计算列不可以，同时如果非空时两列的数量要相等
            {
                throw new ArgumentException("参数出错");
            }


            List<(int Index, object GrpFiled, object CalField)> srcDatas = new List<(int Index, object GrpFiled, object CalField)>();

            for (int i = 0; i < arrDim0CalRange; i++)
            {
                (int Index, object GrpFiled, object CalField) row = (0, null, 0);

                row.Index = i;
                if (arrDim0groupRange != 1)
                {
                    row.GrpFiled = grpFields[i];
                }
                row.CalField = calculateRange[i, 0];
                srcDatas.Add(row);
            }
            return srcDatas;
        }



        private static List<(int Index, object GrpFiled, object ReturnField, decimal orderField1, decimal orderField2, decimal orderField3)> GetSrcData(object[,] groupRange, object[,] returnRange, object[,] orderRange1, object[,] orderRange2, object[,] orderRange3)
        {
            if (orderRange1.GetLength(1) != 1 || orderRange2.GetLength(1) != 1 || orderRange3.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }


            int arrDim0groupRange = groupRange.GetLength(0);
            int arrReturn0Length = returnRange.GetLength(0);
            int arrDim0Length1 = orderRange1.GetLength(0);
            int arrDim0Length2 = orderRange2.GetLength(0);
            int arrDim0Length3 = orderRange3.GetLength(0);

            //必须有return列，其他列都可以省略
            if ((arrReturn0Length != arrDim0Length1 && arrDim0Length1 != 1) || (arrReturn0Length != arrDim0Length2 && arrDim0Length2 != 1) || (arrReturn0Length != arrDim0Length3 && arrDim0Length3 != 1))
            {
                throw new ArgumentException("参数出错");
            }


            List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)> srcDatas = new List<(int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3)>();

            for (int i = 0; i < arrReturn0Length; i++)
            {
                (int Index, object GrpFiled, object ReturnField, decimal OrderField1, decimal OrderField2, decimal OrderField3) row = (0, null, null, 0, 0, 0);

                row.Index = i;
                row.ReturnField = returnRange[i, 0];

                if (arrDim0groupRange != 1)
                {
                    row.GrpFiled = groupRange[i, 0];
                }

                if (arrDim0Length1 != 1)
                {
                    row.OrderField1 = GetValueOfOrderField(orderRange1[i, 0]);
                }
                if (arrDim0Length2 != 1)
                {
                    row.OrderField2 = GetValueOfOrderField(orderRange2[i, 0]);
                }
                if (arrDim0Length3 != 1)
                {
                    row.OrderField3 = GetValueOfOrderField(orderRange3[i, 0]);
                }

                srcDatas.Add(row);

            }
            return srcDatas;
        }


        private static List<(int Index, object GrpFiled, decimal orderField1, decimal orderField2, decimal orderField3)> GetSrcData(object groupRange, object[,] orderRange1, object[,] orderRange2, object[,] orderRange3)
        {
            if (orderRange1.GetLength(1) != 1 || orderRange2.GetLength(1) != 1 || orderRange3.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0groupRange = 0;
            List<object> grpFields = new List<object>();
            if (groupRange is object[])
            {
                var arrGrp = groupRange as object[];
                arrDim0groupRange = arrGrp.GetLength(0);
                grpFields.AddRange(arrGrp);
            }
            else if (groupRange is object[,])
            {
                var arrGrp = groupRange as object[,];
                arrDim0groupRange = (groupRange as object[,]).GetLength(0);
                for (int i = 0; i < arrGrp.GetLength(0); i++)
                {
                    grpFields.Add(arrGrp[i, 0]);
                }
            }

            int arrDim0Length1 = orderRange1.GetLength(0);
            int arrDim0Length2 = orderRange2.GetLength(0);
            int arrDim0Length3 = orderRange3.GetLength(0);

            if ((arrDim0groupRange != arrDim0Length1 && arrDim0Length1 != 1) || (arrDim0groupRange != arrDim0Length2 && arrDim0Length2 != 1) || (arrDim0groupRange != arrDim0Length3 && arrDim0Length3 != 1))
            {
                throw new ArgumentException("参数出错");
            }


            List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)> srcDatas = new List<(int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3)>();

            for (int i = 0; i < grpFields.Count; i++)
            {
                (int Index, object GrpFiled, decimal OrderField1, decimal OrderField2, decimal OrderField3) row = (0, null, 0, 0, 0);

                row.Index = i;
                row.GrpFiled = grpFields[i];

                if (arrDim0Length1 != 1)
                {
                    row.OrderField1 = GetValueOfOrderField(orderRange1[i, 0]);
                }
                if (arrDim0Length2 != 1)
                {
                    row.OrderField2 = GetValueOfOrderField(orderRange2[i, 0]);
                }
                if (arrDim0Length3 != 1)
                {
                    row.OrderField3 = GetValueOfOrderField(orderRange3[i, 0]);
                }

                srcDatas.Add(row);

            }
            return srcDatas;
        }


        private static decimal GetValueOfOrderField(object value)
        {
            if (value is string)
            {
                byte[] array = System.Text.Encoding.Unicode.GetBytes(value as string);
                decimal strSum = 0;
                int scaleMax = 15;
                int iScale = 0;
                int loopLength = array.Length > 6 ? 6 : array.Length;
                for (int i = 0; i < loopLength; i = i + 2)
                {
                    int strInt = (int)(array[i]) + (int)(array[i + 1]);
                    strSum = strSum + Convert.ToDecimal(strInt * Math.Pow(10, (scaleMax - 3 * iScale)));
                    iScale++;
                }
                return decimal.MaxValue / 10 + strSum; //先把double缩小10倍，可以相加新的文本数字;
            }
            else if (value is ExcelError)
            {
                return decimal.MinValue + 1;
            }
            else if (value is bool)
            {
                return decimal.MaxValue - 1 + Convert.ToDecimal(value);//如果是逻辑值的话，Excel里逻辑值比字符还要大
            }
            else if (value is ExcelEmpty || value is null)
            {
                return decimal.MinValue;
            }
            else
            {
                return Convert.ToDecimal(value);
            }

        }
    }
}
