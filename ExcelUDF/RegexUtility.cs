using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using System.Text.RegularExpressions;
using System.IO;

namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {
        //--input=输入
        //--pattern=匹配规则
        //--matchNum=确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0
        //--groupNum=确定第几组匹配，索引号从1开始，0为返回上层的match内容。
        //--isCompiled=是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好
        //--isECMAScript，用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突。
        //--RegexOptions.ECMAScript 选项只能与 RegexOptions.IgnoreCase 和 RegexOptions.Multiline 选项结合使用。在正则表达式中使用其他选项会导致 ArgumentOutOfRangeException。
        //--isRightToLeft，从右往左匹配。
        //--returnNum，反回split数组中的第几个元素，索引从0开始

        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配组，Pattern里传入（）来分组。Excel催化剂出品，必属精品！")]
        public static string RegexMatchGroup(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0，默认为0")] int matchNum = 0,
           [ExcelArgument(Description = "确定第几组匹配，索引号从1开始，0为返回上层的match内容，默认为1")] int groupNum = 1,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                MatchCollection matches = Regex.Matches(input, pattern, options);

                if (matchNum <= matches.Count - 1)
                {
                    Match match = matches[matchNum];
                    if (groupNum == 0)
                    {
                        return match.Value;
                    }
                    else
                    {
                        if (groupNum < match.Groups.Count)
                        {
                            return match.Groups[groupNum].Value;
                        }
                        else
                        {
                            return "";
                        }
                    }
                }
                else
                {
                    return "";
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return "";
            }
        }


        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配组，返回多个结果，Pattern里传入（）来分组。Excel催化剂出品，必属精品！")]
        public static object RegexMatchGroups(
   [ExcelArgument(Description = "输入的字符串")] string input,
   [ExcelArgument(Description = "匹配规则")] string pattern,
   [ExcelArgument(Description = "确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0，默认为0")] int matchNum = 0,
   [ExcelArgument(Description = "确定最终返回的数据是以行（H）排列还是以列(L)排列，传入非H字符或不传参数默认为L排列。")]  string optAlignHorL = "L",
   [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
   [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
   [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                MatchCollection matches = Regex.Matches(input, pattern, options);

                if (matchNum <= matches.Count - 1)
                {
                    Match match = matches[matchNum];
                    if (match.Groups.Count > 1)
                    {
                        string[] result = match.Groups.Cast<Group>().Select(m => m.Value).Skip(1).ToArray();
                        return Common.ReturnDataArray(result, optAlignHorL);
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return "";
            }
        }

        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配组，返回多个结果，Pattern里传入（）来分组。Excel催化剂出品，必属精品！")]
        public static object RegexMatchGroupsFromFile(
  [ExcelArgument(Description = "从文件中传入源字符串")] string fileFullPath,
  [ExcelArgument(Description = "匹配规则")] string pattern,
  [ExcelArgument(Description = "确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0，默认为0")] int matchNum = 0,
  [ExcelArgument(Description = "确定最终返回的数据是以行（H）排列还是以列(L)排列，传入非H字符或不传参数默认为L排列。")]  string optAlignHorL = "L",
  [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
  [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
  [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            using (StreamReader sr = new StreamReader(fileFullPath))
            {
                return RegexMatchGroups(
                    input: sr.ReadToEnd(),
                    pattern: pattern,
                    matchNum: matchNum,
                    optAlignHorL: optAlignHorL,
                    isCompiled: isCompiled,
                    isECMAScript: isECMAScript,
                    isRightToLeft: isRightToLeft
                    );

            };
        }


        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配，不含Group组匹配。Excel催化剂出品，必属精品！")]
        public static string RegexMatch(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "确定第几个匹配返回值，索引号从0开始，第1个匹配，传入0，默认为0")] int matchNum = 0,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                MatchCollection matches = Regex.Matches(input, pattern, options);

                if (matchNum < matches.Count )
                {
                    return matches[matchNum].Value;
                }
                else
                {
                    return "";
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return "";
            }
        }

        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配，返回多个结果，不含Group组匹配。Excel催化剂出品，必属精品！")]
        public static object RegexMatchs(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "确定最终返回的数据是以行（H）排列还是以列(L)排列，传入非H字符或不传参数默认为L排列。")]  string optAlignHorL = "L",
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                MatchCollection matches = Regex.Matches(input, pattern, options);

                if (matches.Count > 1)
                {
                    string[] result = matches.Cast<Match>().Select(m => m.Value).ToArray();
                    return Common.ReturnDataArray(result, optAlignHorL);
                }
                else
                {
                    return Regex.Match(input, pattern, options).Value;
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return "";
            }
        }

        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配，返回多个结果，不含Group组匹配。Excel催化剂出品，必属精品！")]
        public static object RegexMatchsFromFile(
   [ExcelArgument(Description = "从文件中传入源字符串")] string fileFullPath,
   [ExcelArgument(Description = "匹配规则")] string pattern,
   [ExcelArgument(Description = "确定最终返回的数据是以行（H）排列还是以列(L)排列，传入非H字符或不传参数默认为L排列。")]  string optAlignHorL = "L",
   [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
   [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
   [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            using (StreamReader sr = new StreamReader(fileFullPath))
            {
                return RegexMatchs(
                    input: sr.ReadToEnd(),
                    pattern: pattern,
                    optAlignHorL: optAlignHorL,
                    isCompiled: isCompiled,
                    isECMAScript: isECMAScript,
                    isRightToLeft: isRightToLeft
                    );
            };

        }

        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则替换。Excel催化剂出品，必属精品！")]
        public static string RegexReplace(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "匹配到的文件替换的字符串，默认为替换为空")] string replacement = "",
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                return Regex.Replace(input, pattern, replacement, options);
            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return input;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <param name="pattern"></param>
        /// <param name="returnNum">索引从0开始</param>
        /// <param name="isCompiled"></param>
        /// <param name="isECMAScript"></param>
        /// <param name="isRightToLeft"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则分割。Excel催化剂出品，必属精品！")]
        public static string RegexSplit(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "分割后返回第几个项目，索引号从0开始，第1个匹配，传入0，，默认为0")] int returnNum = 0,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                string[] splitResult = Regex.Split(input, pattern, options);
                if (returnNum <= splitResult.Length - 1)
                {
                    return splitResult[returnNum];
                }
                else
                {
                    return "";
                }

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return input;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <param name="pattern"></param>
        /// <param name="returnNum">索引从0开始</param>
        /// <param name="isCompiled"></param>
        /// <param name="isECMAScript"></param>
        /// <param name="isRightToLeft"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则分割。Excel催化剂出品，必属精品！")]
        public static object RegexSplits(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "确定最终返回的数据是以行（H）排列还是以列(L)排列，传入非H字符或不传参数默认为L排列。")]  string optAlignHorL = "L",
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {

            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                string[] splitResult = Regex.Split(input, pattern, options);

                return Common.ReturnDataArray(splitResult, optAlignHorL);

            }
            catch (ArgumentOutOfRangeException)
            {
                return "Options错误";
            }
            catch (ArgumentNullException)
            {
                return "";
            }

            catch (ArgumentException)
            {
                return "Pattern错误";
            }
            catch (Exception)
            {
                return input;
            }
        }

        [ExcelFunction(Category = "文本处理_正则相关", Description = "正则匹配判断。Excel催化剂出品，必属精品！")]
        public static bool RegexIsMatch(
           [ExcelArgument(Description = "输入的字符串")] string input,
           [ExcelArgument(Description = "匹配规则")] string pattern,
           [ExcelArgument(Description = "是否编译，是为1，否为0，暂时没有测试过哪个快在数据量大时，文档好像说数据量大用编译比较好，默认为false")] bool isCompiled = false,
           [ExcelArgument(Description = @"用来指定\w是否匹配一些特殊编码之类的例如中文，当false时会匹配中文,指定为true时，可能和其他的指定有些冲突，默认为false")] bool isECMAScript = false,
           [ExcelArgument(Description = "从右往左匹配，默认为false")] bool isRightToLeft = false)
        {
            try
            {
                RegexOptions options = GetRegexOptions(isCompiled, isECMAScript, isRightToLeft);
                return Regex.IsMatch(input, pattern, options);
            }

            catch (Exception)
            {
                return false;
            }
        }

        private static RegexOptions GetRegexOptions(bool isCompiled, bool isECMAScript, bool isRightToLeft)
        {
            List<RegexOptions> listOptions = new List<RegexOptions>();
            if (isCompiled == true)
            {
                listOptions.Add(RegexOptions.Compiled);
            }
            if (isRightToLeft == true)
            {
                listOptions.Add(RegexOptions.RightToLeft);
            }
            if (isECMAScript == true)
            {
                listOptions.Add(RegexOptions.ECMAScript);
            }

            RegexOptions options = new RegexOptions();
            foreach (var item in listOptions)
            {
                if (options == 0)
                {
                    options = item;
                }
                else
                {
                    options = options | item;
                }
            }

            return options;
        }

    }
}
