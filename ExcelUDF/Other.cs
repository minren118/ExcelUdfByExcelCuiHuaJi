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

        [ExcelFunction(Category = "其他高频函数",IsThreadSafe =true, Description = "不分组下的单列中国式排名。Excel催化剂出品，必属精品！")]

        public static object PM中国式排名(
                 [ExcelArgument(Description = "排名列区域")] object[,] rankRange,
                 [ExcelArgument(Description = "排名列区域3是否按降序排列，默认FALSE为从大到小排名，TRUE时为从小到大排名")] bool isRankAsc
             )
        {

            List<(int Index, decimal RankField)> queryList = GetSrcData(rankRange);

            Dictionary<int, int> result = new Dictionary<int, int>();
            foreach (var item in queryList)
            {

                if (isRankAsc)
                {
                    int rankIndex = queryList.Where(s => s.RankField < item.RankField).Select(s => s.RankField).Distinct().Count() + 1;
                    result.Add(item.Index, rankIndex);
                }
                else
                {
                    int rankIndex = queryList.Where(s => s.RankField > item.RankField).Select(s => s.RankField).Distinct().Count() + 1;
                    result.Add(item.Index, rankIndex);
                }
            }
            return Common.ReturnDataArray(result.OrderBy(s => s.Key).Select(s => (object)s.Value).ToArray(), "L");
        }


        private static List<(int Index, decimal calculateField)> GetSrcData(object[,] calculateRange)
        {

            if (calculateRange.GetLength(1) != 1)
            {
                throw new ArgumentException("参数出错");
            }

            int arrDim0CalRange = calculateRange.GetLength(0);


            List<(int Index, decimal CalField)> srcDatas = new List<(int Index, decimal CalField)>();

            for (int i = 0; i < arrDim0CalRange; i++)
            {
                (int Index, decimal CalField) row = (0, 0);

                row.Index = i;
                row.CalField = GetValueOfOrderField(calculateRange[i, 0]);
                srcDatas.Add(row);
            }
            return srcDatas;
        }


    }
}
