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

        [ExcelFunction(Category = "加密解密", Description = "MD5加密。Excel催化剂出品，必属精品！")]
        public static string Md5String(
             [ExcelArgument(Description = "传入要加密的字符串")] string input,
             [ExcelArgument(Description = "md5返回字符长度，16位或32位")] bool isCodeLength16
            )

        {
            if (!string.IsNullOrEmpty(input))
            {
                // 16位MD5加密（取32位加密的9~25字符）  
                if (isCodeLength16 == true)
                {
                    return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(input, "MD5").ToUpper().Substring(8, 16);
                }
                else
                {
                    return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(input, "MD5").ToUpper();
                }
            }
            return string.Empty;
        }

        //加密
        [ExcelFunction(Category = "加密解密", Description = "RSA加密函数。Excel催化剂出品，必属精品！")]
        public static string EncryptValue(
            [ExcelArgument(Description = "传入要加密的字符串")] string input,
            [ExcelArgument(Description = "密码因子")] string passwordChars
            )
        {
            CspParameters param = new CspParameters();
            param.KeyContainerName = passwordChars;//密匙容器的名称，保持加密解密一致才能解密成功
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(param))
            {
                byte[] plaindata = Encoding.Default.GetBytes(input);//将要加密的字符串转换为字节数组
                byte[] encryptdata = rsa.Encrypt(plaindata, false);//将加密后的字节数据转换为新的加密字节数组
                return Convert.ToBase64String(encryptdata);//将加密后的字节数组转换为字符串
            }
        }

        //解密
        [ExcelFunction(Category = "加密解密", Description = "RSA解密函数。Excel催化剂出品，必属精品！")]
        public static string DecryptValue(
            [ExcelArgument(Description = "传入要解密的字符串")] string input,
            [ExcelArgument(Description = "密码因子")] string passwordChars)
        {
            CspParameters param = new CspParameters();
            param.KeyContainerName = passwordChars;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(param))
            {
                byte[] encryptdata = Convert.FromBase64String(input);
                byte[] decryptdata = rsa.Decrypt(encryptdata, false);
                return Encoding.Default.GetString(decryptdata);
            }
        }



    }
}
