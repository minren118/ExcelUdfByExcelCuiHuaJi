using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace ExcelCuiHuaJi
{
    public partial class ExcelUDF
    {

        [ExcelFunction(Category = "随机函数",IsVolatile =true, Description = "根据传入的字符长度，随机返回传入个数的任意26个大写字母。Excel催化剂出品，必属精品！")]
        public static string RandEnglishCharsUpper(
            [ExcelArgument(Description = "传入需要返回的字符个数")] int charNum = 1)
        {
            List<char> listChar = new List<char>();
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < charNum; i++)
            {
                int randint = rnd.Next(65, 91);
                listChar.Add((char)randint);
            }
            return new string(listChar.ToArray());
        }

        [ExcelFunction(Category = "随机函数", IsVolatile = true, Description = "根据传入的字符长度，随机返回传入个数的任意26个小写字母。Excel催化剂出品，必属精品！")]
        public static string RandEnglishCharsLower(
            [ExcelArgument(Description = "传入需要返回的字符个数")] int charNum = 1)
        {
            List<char> listChar = new List<char>();
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < charNum; i++)
            {
                int randint = rnd.Next(97, 123);
                listChar.Add((char)randint);
            }
            return new string(listChar.ToArray());
        }

        [ExcelFunction(Category = "随机函数", IsVolatile = true, Description = "根据传入的字符长度，随机返回传入个数的任意26个字母，不分大小写。Excel催化剂出品，必属精品！")]
        public static string RandEnglishChars(
            [ExcelArgument(Description = "传入需要返回的字符个数")] int charNum = 1)
        {
            List<char> listChar = new List<char>();
            Random rnd = new Random(Guid.NewGuid().GetHashCode());

            for (int i = 0; i < charNum; i++)
            {
                int rndUorL = rnd.Next(0, 2);//不含上限的
                if (rndUorL == 0)
                {
                    int randint = rnd.Next(65, 91);
                    listChar.Add((char)randint);
                }
                else
                {
                    int randint = rnd.Next(97, 123);
                    listChar.Add((char)randint);
                }

            }
            return new string(listChar.ToArray());
        }

        [ExcelFunction(Category = "随机函数", IsVolatile = true, Description = "根据传入的字符长度，随机返回传入个数任意的0-9数字。Excel催化剂出品，必属精品！")]
        public static string RandNumberString(
            [ExcelArgument(Description = "传入需要返回的字符个数")] int charNum = 1)
        {
            List<char> listChar = new List<char>();
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < charNum; i++)
            {
                int randint = rnd.Next(48, 58);
                listChar.Add((char)randint);
            }
            return new string(listChar.ToArray());
        }


        [ExcelFunction(Category = "随机函数",  Description = "根据传入的字符长度，随机返回传入个数的英文字母或数字，即【0-9，a-z,A-Z】。Excel催化剂出品，必属精品！")]
        public static string RandNumberOrEnglishchars(
            [ExcelArgument(Description = "传入需要返回的字符个数")] int charNum = 1)
        {
            List<char> listChar = new List<char>();
            Random rnd = new Random(Guid.NewGuid().GetHashCode());

            for (int i = 0; i < charNum; i++)
            {
                int rndType = rnd.Next(0, 3);
                if (rndType == 0)
                {
                    int randint = rnd.Next(65, 91);
                    listChar.Add((char)randint);
                }
                else if(rndType==1)
                {
                    int randint = rnd.Next(97, 123);
                    listChar.Add((char)randint);
                }
                else
                {
                    int randint = rnd.Next(48, 58);
                    listChar.Add((char)randint);
                }

            }
            return new string(listChar.ToArray());
        }

        [ExcelFunction(Category = "随机函数",  Description = "根据传入的字符长度，随机返回传入个数的英文字母或数字，即【0-9，a-z,A-Z】。Excel催化剂出品，必属精品！")]
        public static object RandcharsByCustom(
             [ExcelArgument(Description = "传入指定范围内的字符，如0-3为0123，b-d为bcd，多个条件之间用逗号分开")] string customchars ,
            [ExcelArgument(Description = "传入需要返回的字符个数")] int charNum = 1)
        {
            //if (Regex.IsMatch(customchars,"[^a-zA-Z0-9-，,]"))
            //{
            //    return ExcelError.ExcelErrorNA;
            //}

            string[] paras = customchars.Split(new char[] { ',', '，' });
            List<byte> byteparas = new List<byte>();

            foreach (var para in paras)
            {
                //如果传入的是-，
                if (para.Contains("-") && para.Length>1)
                {
                    //且只有一个 -，且前后只有一个字符char
                    if (para.Length==3 && para.Trim(new char[] { '-'})==para)
                    {
                        //用-相连的两个字符不是同时大写或小写时返回错误
                        if (para.ToUpper()!=para && para.ToLower()!=para)
                        {
                            return ExcelError.ExcelErrorNA;
                        }
                        Byte[] encodedBytes = Encoding.ASCII.GetBytes(para);
                        byte lbyte = encodedBytes[0];
                        byte ubyte = encodedBytes[2];

                        //当用-分开的两个字符间的间隔大于26，或者一个是数字一个是字母
                        if ( Math.Max(lbyte,ubyte)>=65 && Math.Min(lbyte,ubyte)<=57)
                        {
                            return ExcelError.ExcelErrorNA;
                        }

                        for (byte i = Math.Min(lbyte,ubyte); i < Math.Max(lbyte, ubyte) + 1; i++)
                        {
                            byteparas.Add(i);
                        }
                    }
                    //有多个-，返回报错
                    else
                    {
                        return ExcelError.ExcelErrorNA;
                    }
                   
                }
                else
                {
                    byteparas.AddRange(Encoding.ASCII.GetBytes(para));
                }
            }

            List<char> listChar = new List<char>();
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < charNum; i++)
            {
                int randint = rnd.Next(byteparas.Count);
                listChar.Add((char)byteparas[randint]);
            }
            //return Encoding.ASCII.GetString(listChar.ToArray());
            return new string(listChar.ToArray());
        }


    }
}
