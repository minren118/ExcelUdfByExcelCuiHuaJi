using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using static ExcelDna.Integration.XlCall;
namespace ExcelCuiHuaJi
{

    public partial class ExcelUDF
    {
        [ExcelFunction(Category = "文件文件夹相关", Description = "根据给定的路径字符，合成文件路径，可以末尾无需输入路径分隔符。Excel催化剂出品，必属精品！")]
        public static object PathCombine(
   [ExcelArgument(Description = @"从左往右拼接的文件夹/文件名路径，最后的\可有可没有")] string path1,
   [ExcelArgument(Description = @"从左往右拼接的文件夹/文件名路径径，最后的\可有可没有")] string path2,
   [ExcelArgument(Description = @"从左往右拼接的文件夹/文件名路径，最后的\可有可没有")] string path3,
   [ExcelArgument(Description = @"从左往右拼接的文件夹/文件名路径，最后的\可有可没有")] string path4,
   [ExcelArgument(Description = @"从左往右拼接的文件夹/文件名路径，最后的\可有可没有")] string path5)
        {
            List<string> paths = new List<string>();
            if (string.IsNullOrEmpty(path1.Trim()))
            {
                return string.Empty;
            }
            else
            {
                paths.Add(path1);
            }
            if (!string.IsNullOrEmpty(path2.Trim()))
            {
                paths.Add(path2);
            }
            if (!string.IsNullOrEmpty(path3.Trim()))
            {
                paths.Add(path3);
            }
            if (!string.IsNullOrEmpty(path4.Trim()))
            {
                paths.Add(path4);
            }
            if (!string.IsNullOrEmpty(path5.Trim()))
            {
                paths.Add(path5);
            }

            return Path.Combine(paths.ToArray());
        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "获取上一级的文件夹全路径。Excel催化剂出品，必属精品！")]
        public static string GetDirectoryName(
                                    [ExcelArgument(Description = "传入一个的文件或文件夹全路径字符串")] string srcFullpath)
        {
            return Path.GetDirectoryName(srcFullpath);

        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "获取文件大小，单位KB。Excel催化剂出品，必属精品！")]
        public static object GetFileSize(
                            [ExcelArgument(Description = "传入一个的文件全路径字符串")] string srcFullpath)
        {

            FileInfo fileInfo = new FileInfo(srcFullpath);
            if (fileInfo.Exists)
            {
                Common.ChangeNumberFormat("#,##0");
                return fileInfo.Length / 1024;
            }
            else
            {
                return "文件不存在";
            }

        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "判断传入的文件或文件 夹路径是否是真实存在。Excel催化剂出品，必属精品！")]
        public static bool IsFileOrDirExist(
            [ExcelArgument(Description = "传入一个文件或文件夹全名字符串")] string srcFullpath)
        {
            if (File.Exists(srcFullpath))
            {
                return true;
            }
            else
            {
                return Directory.Exists(srcFullpath);
            }

        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "获取文件或文件夹创建时间。Excel催化剂出品，必属精品！")]
        public static object GetFileOrDirCreateTime(
            [ExcelArgument(Description = "传入一个文件或文件夹路径")] string fileOrdirPath)
        {
            Common.ChangeNumberFormat("yyyy-mm-dd HH:MM:SS");
            if (File.Exists(fileOrdirPath))
            {
                return File.GetCreationTime(fileOrdirPath);
            }
            else if (Directory.Exists(fileOrdirPath))
            {
                return Directory.GetCreationTime(fileOrdirPath);
            }
            else
            {
                return "文件或文件夹不存在";
            }
        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "获取文件或文件夹属性。Excel催化剂出品，必属精品！")]
        public static object GetFileOrDirAttributes(
            [ExcelArgument(Description = "传入一个文件或文件夹路径")] string fileOrdirPath)
        {

            if (File.Exists(fileOrdirPath))
            {
                return File.GetAttributes(fileOrdirPath).ToString();
            }
            else if (Directory.Exists(fileOrdirPath))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(fileOrdirPath);
                return directoryInfo.Attributes.ToString();
            }
            else
            {
                return "文件或文件夹不存在";
            }
        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "获取文件或文件夹最后修改时间。Excel催化剂出品，必属精品！")]
        public static object GetFileOrDirModifyTime(
            [ExcelArgument(Description = "传入一个文件或文件夹路径")] string fileOrdirPath)
        {
            Common.ChangeNumberFormat("yyyy-mm-dd HH:MM:SS");
            if (File.Exists(fileOrdirPath))
            {
                return File.GetLastWriteTime(fileOrdirPath);
            }
            else if (Directory.Exists(fileOrdirPath))
            {
                return Directory.GetLastWriteTime(fileOrdirPath);
            }
            else
            {
                return "文件或文件夹不存在";
            }
        }


        [ExcelFunction(Category = "文件文件夹相关", Description = "在一个全路径下获取文件名，格式为：文件名+后缀名,如：C:\test\test.txt返回的是test.txt字符串。Excel催化剂出品，必属精品！")]
        public static string GetFileName(
             [ExcelArgument(Description = "传入一个含路径的文件全名字符串")] string srcFullpath)
        {
            return Path.GetFileName(srcFullpath);
        }

        [ExcelFunction(Category = "文件文件夹相关", Description = "在一个全路径下获取文件名的后缀名，如：C:\test\test.txt返回的是.txt)。Excel催化剂出品，必属精品！")]
        public static string GetFileExtension(
                     [ExcelArgument(Description = "传入一个含路径的文件全名字符串")] string srcFullpath)
        {
            return Path.GetExtension(srcFullpath);
        }


        [ExcelFunction(Category = "文件文件夹相关", Description = "在一个全路径下获取文件名不含文件后缀名，如：C:\test\test.txt返回的是test字符串)。Excel催化剂出品，必属精品！")]
        public static string GetFileNameWithoutExtension(
                             [ExcelArgument(Description = "传入一个含路径的文件全名字符串")] string srcFullpath)
        {
            return Path.GetFileNameWithoutExtension(srcFullpath);
        }


        [ExcelFunction(Category = "文件文件夹相关", Description = "获取指定目录下的子文件夹,srcFolder为传入的顶层目录，containsText可用作筛选包含containsText内容的文件夹，isSearchAllDirectory为是否查找顶层目录下的文件夹的所有子文件夹。Excel催化剂出品，必属精品！")]
        public static object GetSubFolders(
            [ExcelArgument(Description = "传入的顶层目录，最终返回的结果将是此目录下的文件夹或子文件夹")] string srcFolder,
            [ExcelArgument(Description = "查找的文件夹中是否需要包含指定字符串，不传参数默认为返回所有文件夹，可传入复杂的正则表达式匹配。")] string optContainsText,
            [ExcelArgument(Description = "是否查找顶层目录下的文件夹的所有子文件夹，TRUE和非0的字符或数字为搜索子文件夹，其他为否，不传参数时默认为否")] object optIsSearchAllDirectory,
            [ExcelArgument(Description = "返回的结果是按按列排列还是按行排列，传入L按列排列，传入H按行排列，不传参数或传入非L或H则默认按列排列")] string optAlignHorL)
        {

            string[] subfolders;
            if (Common.IsMissOrEmpty(optContainsText))
            {
                optContainsText = string.Empty;
            }
            //当isSearchAllDirectory为空或false，默认为只搜索顶层文件夹
            if (Common.IsMissOrEmpty(optIsSearchAllDirectory) || Common.TransBoolPara(optIsSearchAllDirectory) == false)
            {
                subfolders = Directory.EnumerateDirectories(srcFolder).Where(s => isContainsText(s, optContainsText)).ToArray();
            }
            else
            {

                subfolders = Directory.EnumerateDirectories(srcFolder, "*", SearchOption.AllDirectories).Where(s => isContainsText(s, optContainsText)).ToArray();
            }

            return Common.ReturnDataArray(subfolders, optAlignHorL);

        }


        [ExcelFunction(Category = "文件文件夹相关", Description = "获取指定目录下的文件清单,srcFolder为传入的顶层目录，containsText可用作筛选包含containsText内容的文件夹，isSearchAllDirectory为是否查找顶层目录下的文件夹的所有子文件夹。Excel催化剂出品，必属精品！")]
        public static object GetFiles(
                [ExcelArgument(Description = "传入的顶层目录，最终返回的结果将是此目录下的文件夹或子文件夹下的全路径文件名")] string srcFolder,
                [ExcelArgument(Description = "查找的文件名中是否需要包含指定字符串，不传参数默认为返回所有文件，可传入复杂的正则表达式匹配。")] string containsText,
                [ExcelArgument(Description = "是否查找顶层目录下的文件夹的所有子文件夹，TRUE和非0的字符或数字为搜索子文件夹，其他为否，不传参数时默认为否")] object isSearchAllDirectory,
                [ExcelArgument(Description = "返回的结果是按按列排列还是按行排列，传入L按列排列，传入H按行排列，不传参数或传入非L或H则默认按列排列")] string optAlignHorL)
        {
            string[] files;
            if (Common.IsMissOrEmpty(containsText))
            {
                containsText = string.Empty;
            }
            //当isSearchAllDirectory为空或false，默认为只搜索顶层文件夹
            if (Common.IsMissOrEmpty(isSearchAllDirectory) || Common.TransBoolPara(isSearchAllDirectory) == false)
            {
                files = Directory.EnumerateFiles(srcFolder).Where(s => isContainsText(Path.GetFileName(s), containsText)).ToArray();
            }
            else
            {

                files = Directory.EnumerateFiles(srcFolder, "*", SearchOption.AllDirectories).Where(s => isContainsText(Path.GetFileName(s), containsText)).ToArray();
            }

            return Common.ReturnDataArray(files, optAlignHorL);
        }



        [ExcelFunction(Category = "文件文件夹相关", Description = "获取指定目录下的不同层级的文件夹名称。Excel催化剂出品，必属精品！")]
        public static object GetFolderByDepth(
            [ExcelArgument(Description = "传入一个文件夹的详细路径")] string srcFolder,
            [ExcelArgument(Description = "文件夹深度，请传入整数，若是小数将截断，正数为从左到右查，负数为从右往左查,不传参数时默认为右边第1层文件夹")] int intFolderDepth
            )
        {
            List<string> listFolders = srcFolder.Split(new char[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries).ToList();

            //当有文件扩展名，路径为文件全路径，非文件夹全路径
            if (intFolderDepth == 0)
            {
                intFolderDepth = -1;
            }

            if (intFolderDepth < 0)
            {
                listFolders.Reverse();
            }

            if (Path.HasExtension(srcFolder))//当引用的文件夹索引过大时，因文件名占用一个元素，list从0索引开始，故要减2
            {
                if (Math.Abs(intFolderDepth) > listFolders.Count - 2)
                {
                    return ExcelError.ExcelErrorValue;
                }
                else
                {
                    return listFolders[Math.Abs(intFolderDepth)];
                }
            }
            else
            {
                if (Math.Abs(intFolderDepth) > listFolders.Count - 1)
                {
                    return ExcelError.ExcelErrorValue;
                }
                else if (intFolderDepth > 0)
                {
                    return listFolders[intFolderDepth];
                }
                else
                {
                    return listFolders[Math.Abs(intFolderDepth) - 1];
                }

            }

        }



        private static bool isContainsText(string s, string containstext)
        {
            if (string.IsNullOrEmpty(containstext))
            {
                return true;
            }
            else
            {
                return Regex.IsMatch(s, containstext);
            }

        }

    }
}
