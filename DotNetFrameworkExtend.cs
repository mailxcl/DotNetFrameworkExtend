/**
 * .Net Framework 功能扩展
 * 创建人：谢成磊
 * 创建时间：2016/06/05
 * 最后修改时间：2017/07/10
 * 2016/11/16 添加数据库扩展 System.Data 命名空间
 * 2017/02/27 数据驱动方式操作数据库
 * 2017/07/10 添加整理FontHelper
**/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Win32;
using System.Management;
using System.Drawing;
using System.Drawing.Text;
using System.Drawing.Drawing2D;
using TinyJson;
using System.Text;

#region stdole
namespace stdole
{
    public static class ExtendStdole
    {
        /// <summary>
        /// 转为System.Drawing.Font
        /// </summary>
        /// <param name="fntDisp"></param>
        /// <returns></returns>
        public static System.Drawing.Font ToFont(this IFontDisp fntDisp)
        {
            return System.Windows.Forms.FontHelper.GetFontFrom(fntDisp);
        }
    }
}
#endregion

#region System
namespace System
{
    /// <summary>
    /// 扩展System命名空间下的类
    /// </summary>
    public static class ExtendSystem
    {
        /// <summary>
        /// 用已安装的程序打开系统识别扩展名的文件
        /// </summary>
        /// <param name="filePath"></param>
        public static void Open(this string filePath)
        {
            try
            {
                if (filePath == null || filePath == string.Empty)
                {
                    return;
                }
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            { throw ex; }
        }

        /// <summary>
        /// .Net Framework 版本高于或等于4.0.30319.42000返回True
        /// </summary>
        /// <param name="ver"></param>
        /// <returns></returns>
        public static bool GreaterOrEqual_4_0_30319_42000(this Version ver, bool showWarning)
        {
            bool versionValid = true;
            if (ver.Major < 4) versionValid = false;
            if (ver.Build < 30319) versionValid = false;
            if (ver.Revision < 42000) versionValid = false;

            if (versionValid == false && showWarning)
                System.Windows.Forms.MessageBox.Show(
                    ".Net Framework 版本低于4.0.30319.42000\r\n当前版本" + Environment.Version.ToString(),
                    "提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);

            return versionValid;
        }

        /// <summary>
        /// .Net Framework 版本高于或等于4.0.30319.1返回True
        /// </summary>
        /// <param name="ver"></param>
        /// <returns></returns>
        public static bool GreaterOrEqual_4_0_30319_1(this Version ver, bool showWarning)
        {
            bool versionValid = true;
            if (ver.Major < 4) versionValid = false;
            if (ver.Build < 30319) versionValid = false;
            if (ver.Revision < 1) versionValid = false;

            if (versionValid == false && showWarning)
                System.Windows.Forms.MessageBox.Show(
                    ".Net Framework 版本低于4.0.30319.1\r\n当前版本" + Environment.Version.ToString(),
                    "提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);

            return versionValid;
        }

        /// <summary>
        /// 将对象持久化为二进制
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="saveFile"></param>
        /// <param name="encryptKey"></param>
        /// <returns></returns>
        public static byte[] ToEncryptBytes(this object obj, string encryptKey)
        {
            MemoryStream ms = new MemoryStream();
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter b = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            b.Serialize(ms, obj);
            byte[] bytes = ms.GetBuffer();
            byte[] bytesEncrypt = bytes.Encrypt(encryptKey);
            ms.Close();

            return bytesEncrypt;
        }

        /// <summary>
        /// 从二进制中反序列化出对象
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="encryptKey"></param>
        /// <returns></returns>
        public static object GetObjectFromEncryptBytes(this byte[] bytes, string encryptKey)
        {
            object obj = null;
            try
            {
                byte[] bytesDecrypt = bytes.Decrypt(encryptKey);

                MemoryStream ms = new MemoryStream(bytesDecrypt);
                System.Runtime.Serialization.Formatters.Binary.BinaryFormatter b = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                obj = b.Deserialize(ms);
                ms.Close();
            }
            catch (Exception ex)
            {
            }

            return obj;
        }

        /// <summary>
        /// 将对象持久化为文本并保存
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="saveFile"></param>
        /// <param name="encryptKey"></param>
        /// <returns></returns>
        public static bool SaveObjectToEncryptFile(this object obj, string saveFile, string encryptKey)
        {
            try
            {
                byte[] bytesEncrypt = obj.ToEncryptBytes(encryptKey);

                FileStream fileStream = new FileStream(saveFile, FileMode.Create);
                fileStream.Write(bytesEncrypt, 0, bytesEncrypt.Length);
                fileStream.Flush(true);
                fileStream.Close();
            }
            catch (Exception ex)
            { 
            }
            return System.IO.File.Exists(saveFile);
        }

        /// <summary>
        /// 从加密文件中反序列化出对象
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="encryptKey"></param>
        /// <returns></returns>
        public static object GetObjectFromEncryptFile(this string filePath, string encryptKey)
        {
            if (!System.IO.File.Exists(filePath)) return null;

            byte[] bytes = System.IO.File.ReadAllBytes(filePath);
            return bytes.GetObjectFromEncryptBytes(encryptKey);
        }

        /// <summary>
        /// 合并文件夹,并创建新的文件夹
        /// </summary>
        /// <param name="path"></param>
        /// <param name="appendDir">要添加的文件夹</param>
        /// <returns></returns>
        public static string CombineDir(this string path, string appendDir)
        {
            string result = System.IO.Path.Combine(path, appendDir);
            if (!System.IO.Directory.Exists(result)) System.IO.Directory.CreateDirectory(result);
            return result;
        }

        /// <summary>
        /// 要定位的项的 System.String 键。键可以是 null。
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetAppConfigValue(this string name)
        {
            return ConfigurationManager.AppSettings[name];
        }

        /// <summary>
        /// 判断字符串是字母还是数字
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool IsLetterOrDigit(this string s)
        {
            bool result = true;
            char[] chars = s.ToCharArray();
            foreach (char c in chars)
            {
                if (!char.IsLetterOrDigit(c))
                    result = false;
            }
            return result;
        }

        /// <summary>
        /// 合并字符串中连续重复的字符
        /// </summary>
        /// <param name="s"></param>
        /// <param name="split"></param>
        /// <returns></returns>
        public static string UnionRepeatCharToOne(this string s, char split)
        {
            string[] str = s.Split(split);
            string result = string.Empty;
            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] != string.Empty)
                {
                    result += str[i].Trim() + split;
                }
            }
            return result;
        }

        /// <summary>
        /// 删除特定字符
        /// </summary>
        /// <param name="s"></param>
        /// <param name="character"></param>
        /// <returns></returns>
        public static string DeleteChar(this string s, char character)
        {
            string[] str = s.Split(character);
            string result = string.Empty;
            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] != string.Empty)
                {
                    result += str[i].Trim();
                }
            }
            return result;
        }

        public static string RemoveEnterChar(this string s)
        {
            if (s == null) return s;
            string r = Regex.Replace(s, "\r", "");
            r = Regex.Replace(r, "\n", "");
            return r;
        }

        /// <summary>
        /// 把字符串解析为double
        /// </summary>
        /// <param name="s"></param>
        /// <param name="value">输出的数值</param>
        public static double ParseToDouble(this string s)
        {
            double tmp = double.NaN;
            if (double.TryParse(s, out tmp)) return tmp;
            return tmp;
        }
        /// <summary>
        /// 把字符串解析为double
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static double? ParseToDouble2(this string s)
        {
            double? result = null;
            double tmp = double.NaN;
            if (double.TryParse(s, out tmp)) return tmp;
            return result;
        }

        /// <summary>
        /// 把字符串解析为int
        /// </summary>
        /// <param name="s"></param>
        /// <param name="value">输出的数值</param>
        public static int ParseToInt(this string s)
        {
            int tmp = int.MinValue;
            if (int.TryParse(s, out tmp)) return tmp;
            return tmp;
        }
        /// <summary>
        /// 把字符串解析为int
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int? ParseToInt2(this string s)
        {
            int? result = null;
            int tmp = int.MinValue;
            if (int.TryParse(s, out tmp)) return tmp;
            return result;
        }

        /// <summary>
        /// 把字符串解析为日期类型
        /// </summary>
        /// <param name="s"></param>
        /// <param name="value"></param>
        public static bool ParseToDate(this string s, ref DateTime value)
        {
            string dtStr = s.Trim();
            if (dtStr == string.Empty || dtStr == null) return false;

            dtStr = dtStr.Replace("年", "-");
            dtStr = dtStr.Replace("月", "-");
            dtStr = dtStr.Replace("日", "");
            dtStr = dtStr.Trim().Replace("——", "-");
            dtStr = dtStr.Replace("—", "-");
            dtStr = dtStr.Replace("--", "-");

            string split = GetDateTimeSplitChar().ToString();
            dtStr = dtStr.Replace("/", split);
            dtStr = dtStr.Replace(".", split);
            dtStr = dtStr.Replace("-", split);

            return DateTime.TryParse(dtStr, out value);
        }

        /// <summary>
        /// 文件名称合法化，不合法字符用‘_’代替
        /// </summary>
        /// <param name="s"></param>
        public static string FileNameLegalize(this string s)
        {
            char[] charArray = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };
            foreach (char c in charArray) s = s.Replace(c, '_');
            return s;
        }

        /// <summary>
        /// 获取系统时间间隔符
        /// </summary>
        /// <returns></returns>
        private static char? GetDateTimeSplitChar()
        {
            char? split = null;
            char[] dtChars = DateTime.Now.ToString().ToCharArray();
            for (int i = 0; i < dtChars.Length; i++)
            {
                split = dtChars[i];
                if (split.IsInteger())
                    continue;
                else
                    return split;
            }
            return split;
        }

        public static double DMS2Double(this string dmsString)
        {
            double result = double.NaN;
            if (dmsString.Contains('°'))
            {
                string[] dms = dmsString.Split("°′″".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                if (dms.Length < 3) return double.NaN;
                double d = dms[0].ParseToDouble();
                double m = dms[1].ParseToDouble();
                double s = dms[2].ParseToDouble();
                result = d + m / 60.0 + s / 3600.0;
            }
            else
            {
                string[] val = dmsString.Split(new char[1] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                if (val.Length == 2)
                {
                    double r = double.Parse(val[0]), du = 0.0, mi = 0.0;
                    string dus = "", mis = "";
                    if (val[1].Length >= 2)
                    {
                        dus = val[1].Substring(0, 2);
                        mis = val[1].Substring(2);
                        if (mis.Length == 1) mis = mis.Insert(1, ".");
                        if (mis.Length >= 2) mis = mis.Insert(2, ".");
                        du = double.Parse(dus) / 60.0;
                        mi = double.Parse(mis) / 3600.0;
                    }
                    result = (r + du + mi);
                }
                else if (val.Length == 1)
                    result = double.Parse(val[0]);
                else
                    result = double.Parse(dmsString);
            }
            return result;
        }

        /// <summary>
        /// 是否为数字
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool IsNumeric(this string s)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex(String.Format("({0})|({1})", strValidRealPattern, strValidIntegerPattern));

            return !objNotNumberPattern.IsMatch(s) &&
                   !objTwoDotPattern.IsMatch(s) &&
                   !objTwoMinusPattern.IsMatch(s) &&
                    objNumberPattern.IsMatch(s);
        }

        /// <summary>
        /// 是否为整型数字
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool IsInteger(this string s)
        {
            return new Regex("^[0-9]*$").IsMatch(s);
        }

        /// <summary>
        /// 判断字符是否为数字
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static bool IsInteger(this char? c)
        {
            if (c == null) return false;
            return (c >= 48 && c <= 57);
        }

        /// <summary>
        /// 字节转成兆字节
        /// </summary>
        /// <param name="Byte"></param>
        /// <param name="Decimal">小数点后位数</param>
        /// <returns></returns>
        public static double ByteToMB(this double Byte, int Decimal)
        {
            return Math.Round(Byte / 1024 / 1024, Decimal);
        }

        /// <summary>
        /// 汉字转换为Unicode编码
        /// </summary>
        /// <param name="s">要编码的汉字字符串</param>
        /// <returns>Unicode编码的的字符串</returns>
        public static string ToUnicode(this string s)
        {
            byte[] bts = System.Text.Encoding.Unicode.GetBytes(s);
            string r = "";
            for (int i = 0; i < bts.Length; i += 2) r += "\\u" + bts[i + 1].ToString("x").PadLeft(2, '0') + bts[i].ToString("x").PadLeft(2, '0');
            return r;
        }

        /// <summary>
        /// 将Unicode编码转换为汉字字符串
        /// </summary>
        /// <param name="s">Unicode编码字符串</param>
        /// <returns>汉字字符串</returns>
        public static string ToGB2312(this string s)
        {
            string r = "";
            MatchCollection mc = Regex.Matches(s, @"\\u([\w]{2})([\w]{2})", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            byte[] bts = new byte[2];
            foreach (Match m in mc)
            {
                bts[0] = (byte)int.Parse(m.Groups[2].Value, System.Globalization.NumberStyles.HexNumber);
                bts[1] = (byte)int.Parse(m.Groups[1].Value, System.Globalization.NumberStyles.HexNumber);
                r += System.Text.Encoding.Unicode.GetString(bts);
            }
            return r;
        }

        /// <summary>
        /// 十进制度转成度分秒格式
        /// </summary>
        /// <param name="degree"></param>
        /// <returns></returns>
        public static string ToDMS(this double degree, int decimalDigits)
        {
            int dec = decimalDigits;

            int d = (int)degree;
            int m = (int)((degree - d) * 60);
            double s = Math.Round(((degree - d) * 60 - m) * 60, dec);
            int sint = (int)s;
            double sdouble = s - sint;
            int sint2 = (int)(sdouble * Math.Pow(10, dec));

            return string.Format("{0}°{1}′{2}″",
                d.ToString().PadLeft(3, '0'),
                m.ToString().PadLeft(2, '0'),
                sint.ToString().PadLeft(2, '0') + "." + sint2.ToString().PadLeft(dec, '0'));
        }

        /// <summary>
        /// 十进制度转成度分秒格式 无单位字符串
        /// </summary>
        /// <param name="degree"></param>
        /// <returns></returns>
        public static string ToDMSNoUnit(this double degree)
        {
            int dec = 2;

            int d = (int)degree;
            int m = (int)((degree - d) * 60);
            double s = Math.Round(((degree - d) * 60 - m) * 60, dec);
            int sint = (int)s;
            double sdouble = s - sint;
            if (sdouble > 0.5) ++sint;

            return string.Format("{0}{1}{2}",
                d.ToString().PadLeft(3, '0'),
                m.ToString().PadLeft(2, '0'),
                sint.ToString().PadLeft(2, '0'));
        }

        /// <summary>
        /// 获取字符串的MD5码
        /// </summary>
        /// <param name="s">字符串</param>
        /// <returns>MD5值</returns>
        public static string ToMD5String(this string s)
        {
            try
            {
                byte[] data = System.Text.Encoding.Unicode.GetBytes(s.ToCharArray());
                //建立加密服务  
                System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                //加密Byte[]数组  
                byte[] retVal = md5.ComputeHash(data);
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                }
                md5.Clear();
                md5.Dispose();
                return sb.ToString();
            }
            catch (Exception e)
            {
                throw new Exception("MD5Helper.GetMD5FromString(string) Failed, ErrMsg:" + e.Message);
            }
        }

        /// <summary>
        /// 字符串加密
        /// </summary>
        /// <param name="s"></param>
        /// <param name="encryptKey">密钥</param>
        /// <returns></returns>
        public static string Encrypt(this string s, string encryptKey)
        {
            System.Security.Cryptography.DESCryptoServiceProvider descsp = new System.Security.Cryptography.DESCryptoServiceProvider();
            byte[] key = System.Text.Encoding.Unicode.GetBytes(encryptKey);

            byte[] rgbKey = new byte[8];
            byte[] rgbIV = new byte[8];
            for (int i = 0; i < 8; i++)
            {
                rgbKey[i] = key[i];
                rgbIV[i] = key[i];
            }

            byte[] data = System.Text.Encoding.Unicode.GetBytes(s);
            System.Security.Cryptography.ICryptoTransform ctrans = descsp.CreateEncryptor(rgbKey, rgbIV);
            MemoryStream MStream = new MemoryStream();
            System.Security.Cryptography.CryptoStream CStream = new System.Security.Cryptography.CryptoStream(
                MStream, ctrans, System.Security.Cryptography.CryptoStreamMode.Write);
            CStream.Write(data, 0, data.Length);
            CStream.FlushFinalBlock();
            return Convert.ToBase64String(MStream.ToArray());
        }

        /// <summary>
        /// 解密
        /// </summary>
        /// <param name="s"></param>
        /// <param name="encryptKey">密钥</param>
        /// <returns></returns>
        public static byte[] Decrypt(this byte[] data, string encryptKey)
        {
            System.Security.Cryptography.DESCryptoServiceProvider descsp = new System.Security.Cryptography.DESCryptoServiceProvider();
            byte[] key = System.Text.Encoding.Unicode.GetBytes(encryptKey);
            byte[] rgbKey = new byte[8];
            byte[] rgbIV = new byte[8];
            for (int i = 0; i < 8; i++)
            {
                rgbKey[i] = key[i];
                rgbIV[i] = key[i];
            }
            System.Security.Cryptography.ICryptoTransform ctrans = descsp.CreateDecryptor(rgbKey, rgbIV);

            MemoryStream MStream = new MemoryStream();
            System.Security.Cryptography.CryptoStream CStream = new System.Security.Cryptography.CryptoStream(
                MStream, ctrans,
                System.Security.Cryptography.CryptoStreamMode.Write);
            CStream.Write(data, 0, data.Length);
            CStream.FlushFinalBlock();
            return MStream.ToArray();
        }

        /// <summary>
        /// 加密
        /// </summary>
        /// <param name="s"></param>
        /// <param name="encryptKey">密钥</param>
        /// <returns></returns>
        public static byte[] Encrypt(this byte[] data, string encryptKey)
        {
            System.Security.Cryptography.DESCryptoServiceProvider descsp = new System.Security.Cryptography.DESCryptoServiceProvider();
            byte[] key = System.Text.Encoding.Unicode.GetBytes(encryptKey);

            byte[] rgbKey = new byte[8];
            byte[] rgbIV = new byte[8];
            for (int i = 0; i < 8; i++)
            {
                rgbKey[i] = key[i];
                rgbIV[i] = key[i];
            }

            System.Security.Cryptography.ICryptoTransform ctrans = descsp.CreateEncryptor(rgbKey, rgbIV);
            MemoryStream MStream = new MemoryStream();
            System.Security.Cryptography.CryptoStream CStream = new System.Security.Cryptography.CryptoStream(
                MStream, ctrans, System.Security.Cryptography.CryptoStreamMode.Write);
            CStream.Write(data, 0, data.Length);
            CStream.FlushFinalBlock();

            return MStream.ToArray();
        }

        /// <summary>
        /// 解密
        /// </summary>
        /// <param name="s"></param>
        /// <param name="encryptKey">密钥</param>
        /// <returns></returns>
        public static string Decrypt(this string s, string encryptKey)
        {
            byte[] data = Convert.FromBase64String(s);
            System.Security.Cryptography.DESCryptoServiceProvider descsp = new System.Security.Cryptography.DESCryptoServiceProvider();
            byte[] key = System.Text.Encoding.Unicode.GetBytes(encryptKey);
            byte[] rgbKey = new byte[8];
            byte[] rgbIV = new byte[8];
            for (int i = 0; i < 8; i++)
            {
                rgbKey[i] = key[i];
                rgbIV[i] = key[i];
            }
            System.Security.Cryptography.ICryptoTransform ctrans = descsp.CreateDecryptor(rgbKey, rgbIV);

            MemoryStream MStream = new MemoryStream();
            System.Security.Cryptography.CryptoStream CStream = new System.Security.Cryptography.CryptoStream(
                MStream, ctrans,
                System.Security.Cryptography.CryptoStreamMode.Write);
            CStream.Write(data, 0, data.Length);
            CStream.FlushFinalBlock();
            return System.Text.Encoding.Unicode.GetString(MStream.ToArray());
        }

        /// <summary>
        /// 任意字符串转全角字符串(SBC case)
        /// 全角空格为12288，半角空格为32
        /// 其他字符半角(33-126)与全角(65281-65374)的对应关系是：均相差65248
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToSBC(this string input)
        {
            char[] c = input.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 32)
                {
                    c[i] = (char)12288;
                    continue;
                }
                if (c[i] < 127)
                    c[i] = (char)(c[i] + 65248);
            }
            return new String(c);
        }

        /// <summary>
        /// 任意字符串转半角字符串(DBC case)
        /// 全角空格为12288，半角空格为32
        /// 其他字符半角(33-126)与全角(65281-65374)的对应关系是：均相差65248
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToDBC(this string input)
        {
            char[] c = input.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 12288)
                {
                    c[i] = (char)32;
                    continue;
                }
                if (c[i] > 65280 && c[i] < 65375)
                    c[i] = (char)(c[i] - 65248);
            }
            return new String(c);
        }

        public static PropertyInfo GetPropertyInfo(this object obj, string propName)
        {
            if (obj == null) return null;

            PropertyInfo[] props = obj.GetType().GetProperties();

            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                if (prop.Name == propName) return prop;
            }
            return null;
        }

        public static PropertyInfo GetFirstPropertyByType(this object obj, Type propType)
        {
            if (obj == null) return null;

            PropertyInfo[] props = obj.GetType().GetProperties();

            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                if (prop.PropertyType.Equals(propType)) return prop;
            }
            return null;
        }

        public static bool SetPropertyValue(this object obj, string propName, object val)
        {
            bool r = false;
            PropertyInfo prop = obj.GetPropertyInfo(propName);
            if (prop != null)
            {
                try
                {
                    Type propType = prop.PropertyType;
                    Type valType = val.GetType();

                    if (val != null && !propType.Equals(valType))
                        val = convert2TarType(propType, val);

                    prop.SetValue(obj, val, null);
                    r = true;
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.StackTrace, ex.Message);
                    //throw ex;
                }
            }
            else
            {
                r = false;
            }
            return r;
        }

        private static object convert2TarType(Type tarType, object val)
        {
            if (null == val) return null;
            if (DBNull.Value.Equals(val)) return null;

            if (val.GetType().Equals(typeof(byte[])))
            {
                if (tarType.Equals(typeof(byte[])))
                {
                    return (byte[])val;
                }
                else
                {
                    return null;
                }
            }

            if (tarType.Equals(typeof(Boolean?)) || tarType.Equals(typeof(Boolean)))
                val = Convert.ToBoolean(val);
            if (tarType.Equals(typeof(Byte?)) || tarType.Equals(typeof(Byte)))
                val = Convert.ToByte(val);
            if (tarType.Equals(typeof(Char?)) || tarType.Equals(typeof(Char)))
                val = Convert.ToChar(val);
            if (tarType.Equals(typeof(DateTime?)) || tarType.Equals(typeof(DateTime)))
                val = Convert.ToDateTime(val);
            if (tarType.Equals(typeof(Decimal?)) || tarType.Equals(typeof(Decimal)))
                val = Convert.ToDecimal(val);
            if (tarType.Equals(typeof(Double?)) || tarType.Equals(typeof(Double)))
                val = Convert.ToDouble(val);
            if (tarType.Equals(typeof(Int16?)) || tarType.Equals(typeof(Int16)))
                val = Convert.ToInt16(val);
            if (tarType.Equals(typeof(Int32?)) || tarType.Equals(typeof(Int32)))
                val = Convert.ToInt32(val);
            if (tarType.Equals(typeof(Int64?)) || tarType.Equals(typeof(Int64)))
                val = Convert.ToInt64(val);
            if (tarType.Equals(typeof(SByte?)) || tarType.Equals(typeof(SByte)))
                val = Convert.ToSByte(val);
            if (tarType.Equals(typeof(Single?)) || tarType.Equals(typeof(Single)))
                val = Convert.ToSingle(val);
            if (tarType.Equals(typeof(String)))
                val = Convert.ToString(val);
            if (tarType.Equals(typeof(UInt16?)) || tarType.Equals(typeof(UInt16)))
                val = Convert.ToUInt16(val);
            if (tarType.Equals(typeof(UInt32?)) || tarType.Equals(typeof(UInt32)))
                val = Convert.ToUInt32(val);
            if (tarType.Equals(typeof(UInt64?)) || tarType.Equals(typeof(UInt64)))
                val = Convert.ToUInt64(val);
      
            Type vaT = val.GetType();
            return val;
        }


        public static object GetPropertyValue(this object obj, string propName)
        {
            PropertyInfo prop = obj.GetPropertyInfo(propName);
            if (prop == null) return null;
            return prop.GetValue(obj, null);
        }

        public static string GetClassDisplayName(this object obj)
        {
            string result = string.Empty;
            if (obj == null) return result;

            Type argType = obj.GetType();
            DisplayNameAttribute displayNameAttr = (DisplayNameAttribute)Attribute.GetCustomAttribute(argType, typeof(DisplayNameAttribute));
            if (displayNameAttr != null)
                result = displayNameAttr.DisplayName;

            return result;
        }

        /// <summary>
        /// 属性赋值(只有属性名称一致才会赋值)
        /// </summary>
        /// <param name="tarObj">目标对象</param>
        /// <param name="srcObj">数据源对象</param>
        public static void GetAttributesValueFrom(this object tarObj, object srcObj)
        {
            if (srcObj == null) return;

            PropertyInfo[] src = srcObj.GetType().GetProperties();
            PropertyInfo[] tar = tarObj.GetType().GetProperties();

            for (int ii = 0; ii < tar.Length; ii++)
            {
                PropertyInfo tarProp = tar[ii];
                string tName = tarProp.Name;

                IEnumerable<PropertyInfo> srcProps = src.Where(e => { return e.Name.ToLower().Equals(tName.ToLower()); });
                if (srcProps.Count() > 0)
                {
                    PropertyInfo srcProp = srcProps.First();
                    object srcValue = srcProp.GetValue(srcObj, null);
                    tarProp.SetValue(tarObj, srcValue, null);
                }
            }
        }

        /// <summary>
        /// 属性赋值(只有属性名称一致才会赋值)
        /// </summary>
        /// <param name="tarObj">目标对象</param>
        /// <param name="srcObj">数据源对象</param>
        /// <param name="map">目标对象与数据源对象属性值映射表</param>
        public static void GetAttributesValueFrom2(this object tarObj, object srcObj, Dictionary<string, string> map)
        {
            if (srcObj == null) return;

            PropertyInfo[] src = srcObj.GetType().GetProperties();
            PropertyInfo[] tar = tarObj.GetType().GetProperties();

            for (int ii = 0; ii < tar.Length; ii++)
            {
                PropertyInfo tarProp = tar[ii];
                string tName = tarProp.Name;
                if (map.Keys.Contains(tName))
                    tName = map[tName];

                IEnumerable<PropertyInfo> srcProps = src.Where(e => { return e.Name.ToLower().Equals(tName.ToLower()); });
                if (srcProps.Count() > 0)
                {
                    PropertyInfo srcProp = srcProps.First();
                    object srcValue = srcProp.GetValue(srcObj, null);
                    tarProp.SetValue(tarObj, srcValue, null);
                }
            }
        }

        /// <summary>
        /// DateTime转为SQL(ORACLE)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string ToSQLStringORACLE(this DateTime dt)
        {
            return string.Format("to_date('{0}','yyyy-mm-dd hh24:mi:ss')", GetTimeStringFrom(dt));
        }

        public static string ToSQLStringORACLE_Day(this DateTime dt)
        {
            return string.Format("to_date('{0}','yyyy-mm-dd')", GetDayStringFrom(dt));
        }

        /// <summary>
        /// DateTime转为SQL(MSACCESS)
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ToSQLStringMSACCESS(this DateTime dt)
        {
            return "'" + GetTimeStringFrom(dt).Replace('-', '/') + "'";
        }

        /// <summary>
        /// DateTime转为SQL(MSSQL)
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ToSQLStringMSSQL(this DateTime dt)
        {
            return "'" + GetTimeStringFrom(dt).Replace('-', '/') + "'";

        }

        /// <summary>
        /// DateTime转为SQL(SQLITE)
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ToSQLStringSQLITE(this DateTime dt)
        {
            return string.Format("datetime('{0}')", GetTimeStringFrom(dt));
        }

        private static string GetDayStringFrom(DateTime dt)
        {
            return string.Format(
                "{0}-{1}-{2}",
                dt.Year.ToString().PadLeft(4, '0'),
                dt.Month.ToString().PadLeft(2, '0'),
                dt.Day.ToString().PadLeft(2, '0')
                );
        }

        private static string GetTimeStringFrom(DateTime dt)
        {
            return string.Format(
                "{0}-{1}-{2} {3}:{4}:{5}",
                dt.Year.ToString().PadLeft(4, '0'),
                dt.Month.ToString().PadLeft(2, '0'),
                dt.Day.ToString().PadLeft(2, '0'),
                dt.Hour.ToString().PadLeft(2, '0'),
                dt.Minute.ToString().PadLeft(2, '0'),
                dt.Second.ToString().PadLeft(2, '0')
                );
        }

        /// <summary>
        /// 设置对象属性为只读，类属性里要有[ReadOnlyAttribute(false)]标记
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <param name="readOnly"></param>
        public static void SetPropertyReadOnly(this object obj, string propertyName, bool readOnly)
        {
            Type type = typeof(System.ComponentModel.ReadOnlyAttribute);
            obj.SetPropertyAttr(propertyName, type, "isReadOnly", readOnly);
        }

        /// <summary>
        /// 设置对象属性的可见性
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propertyName"></param>
        /// <param name="visible"></param>
        public static void SetPropertyVisibility(this object obj, string propertyName, bool visible)
        {
            Type type = typeof(System.ComponentModel.BrowsableAttribute);
            obj.SetPropertyAttr(propertyName, type, "browsable", visible);
        }

        public static void SetPropertyCategory(this object obj, string propertyName, string category)
        {
            Type type = typeof(System.ComponentModel.CategoryAttribute);
            obj.SetPropertyAttr(propertyName, type, "category", category);
        }

        public static void SetPropertyDisplayName(this object obj, string propertyName, string category)
        {
            Type type = typeof(System.ComponentModel.DisplayNameAttribute);
            obj.SetPropertyAttr(propertyName, type, "displayName", category);
        }

        public static void SetPropertyDescription(this object obj, string propertyName, string category)
        {
            Type type = typeof(System.ComponentModel.DescriptionAttribute);
            obj.SetPropertyAttr(propertyName, type, "description", category);
        }

        private static void SetPropertyAttr(this object obj, string propertyName, Type attrType, string fieldName, object val)
        {
            System.ComponentModel.PropertyDescriptorCollection props = System.ComponentModel.TypeDescriptor.GetProperties(obj);
            System.ComponentModel.PropertyDescriptor descriptor = props[propertyName];
            if (descriptor == null) { return; }
            System.ComponentModel.AttributeCollection attrs = descriptor.Attributes;
            FieldInfo fld = attrType.GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
            fld.SetValue(attrs[attrType], val);
        }

        public static string GetEnumDescription(this Enum value)
        {
            Type enumType = value.GetType();

            string name = Enum.GetName(enumType, value);
            if (name != null)
            {
                FieldInfo fieldInfo = enumType.GetField(name);
                if (fieldInfo != null)
                {
                    DescriptionAttribute attr = Attribute.GetCustomAttribute(fieldInfo, typeof(DescriptionAttribute), false) as DescriptionAttribute;
                    if (attr != null && !string.IsNullOrEmpty(attr.Description))
                        return attr.Description;
                }
            }
            return string.Empty;
        }
    }

    [Serializable]
    public class ItemObject
    {
        public string Text { get; set; }
        public object Value { get; set; }
        public override string ToString()
        {
            return this.Text;
        }
        public override bool Equals(object obj)
        {
            ItemObject itm = obj as ItemObject;
            if (itm == null) return false;
            return (this.Text.Equals(itm.Text) && this.Value.Equals(itm.Value));
        }
    }
}
#endregion

#region System.ComponentModel
namespace System.ComponentModel
{
    /// <summary>  
    /// 枚举转换器  
    /// 用此类之前，必须保证在枚举项中定义了Description  
    /// </summary>  
    public class EnumConverterExt : ExpandableObjectConverter
    {
        /// <summary>  
        /// 枚举项集合  
        /// </summary>  
        Dictionary<object, string> dic;

        /// <summary>  
        /// 构造函数  
        /// </summary>  
        public EnumConverterExt()
        {
            dic = new Dictionary<object, string>();
        }

        /// <summary>  
        /// 加载枚举项集合  
        /// </summary>  
        /// <param name="context"></param>  
        private void LoadDic(ITypeDescriptorContext context)
        {
            dic = GetEnumValueDesDic(context.PropertyDescriptor.PropertyType);
        }

        /// <summary>  
        /// 是否可从来源转换  
        /// </summary>  
        /// <param name="context"></param>  
        /// <param name="sourceType"></param>  
        /// <returns></returns>  
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            if (sourceType == typeof(string))
                return true;

            return base.CanConvertFrom(context, sourceType);
        }

        /// <summary>  
        /// 从来源转换  
        /// </summary>  
        /// <param name="context"></param>  
        /// <param name="culture"></param>  
        /// <param name="value"></param>  
        /// <returns></returns>  
        public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value)
        {
            if (value is string)
            {
                //如果是枚举  
                if (context.PropertyDescriptor.PropertyType.IsEnum)
                {
                    if (dic.Count <= 0)
                        LoadDic(context);
                    if (dic.Values.Contains(value.ToString()))
                    {
                        foreach (object obj in dic.Keys)
                        {
                            if (dic[obj] == value.ToString())
                            {
                                return obj;
                            }
                        }
                    }
                }
            }

            return base.ConvertFrom(context, culture, value);
        }

        /// <summary>  
        /// 是否可转换  
        /// </summary>  
        /// <param name="context"></param>  
        /// <param name="destinationType"></param>  
        /// <returns></returns>  
        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return true;
        }

        /// <summary>  
        ///   
        /// </summary>  
        /// <param name="context"></param>  
        /// <returns></returns>  
        public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
        {
            return true;
        }

        /// <summary>  
        ///   
        /// </summary>  
        /// <param name="context"></param>  
        /// <returns></returns>  
        public override bool GetStandardValuesExclusive(ITypeDescriptorContext context)
        {
            return true;
        }

        /// <summary>  
        ///   
        /// </summary>  
        /// <param name="context"></param>  
        /// <returns></returns>  
        public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
        {
            //ListAttribute listAttribute = (ListAttribute)context.PropertyDescriptor.Attributes[typeof(ListAttribute)];  
            //StandardValuesCollection vals = new TypeConverter.StandardValuesCollection(listAttribute._lst);  

            //Dictionary<object, string> dic = GetEnumValueDesDic(typeof(PKGenerator));  

            //StandardValuesCollection vals = new TypeConverter.StandardValuesCollection(dic.Keys);  

            if (dic == null || dic.Count <= 0)
                LoadDic(context);

            StandardValuesCollection vals = new TypeConverter.StandardValuesCollection(dic.Keys);

            return vals;
        }

        /// <summary>  
        ///   
        /// </summary>  
        /// <param name="context"></param>  
        /// <param name="culture"></param>  
        /// <param name="value"></param>  
        /// <param name="destinationType"></param>  
        /// <returns></returns>  
        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            //DescriptionAttribute.GetCustomAttribute(  
            //EnumDescription  
            //List<KeyValuePair<Enum, string>> mList = UserCombox.ToListForBind(value.GetType());  
            //foreach (KeyValuePair<Enum, string> mItem in mList)  
            //{  
            //    if (mItem.Key.Equals(value))  
            //    {  
            //        return mItem.Value;  
            //    }  
            //}  
            //return "Error!";  

            //绑定控件  
            //            FieldInfo fieldinfo = value.GetType().GetField(value.ToString());  
            //Object[] objs = fieldinfo.GetCustomAttributes(typeof(System.ComponentModel.DescriptionAttribute), false);  
            //if (objs == null || objs.Length == 0)  
            //{  
            //    return value.ToString();  
            //}  
            //else  
            //{  
            //    System.ComponentModel.DescriptionAttribute da = (System.ComponentModel.DescriptionAttribute)objs[0];  
            //    return da.Description;  
            //}  

            if (dic.Count <= 0)
                LoadDic(context);

            foreach (object key in dic.Keys)
            {
                if (key.ToString() == value.ToString() || dic[key] == value.ToString())
                {
                    return dic[key].ToString();
                }
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }

        /// <summary>  
        /// 记载枚举的值+描述  
        /// </summary>  
        /// <param name="enumType"></param>  
        /// <returns></returns>  
        public Dictionary<object, string> GetEnumValueDesDic(Type enumType)
        {
            Dictionary<object, string> dic = new Dictionary<object, string>();
            FieldInfo[] fieldinfos = enumType.GetFields();
            foreach (FieldInfo field in fieldinfos)
            {
                if (field.FieldType.IsEnum)
                {
                    object[] objs = field.GetCustomAttributes(typeof(DescriptionAttribute), false);
                    if (objs.Length > 0)
                    {
                        dic.Add(Enum.Parse(enumType, field.Name), ((DescriptionAttribute)objs[0]).Description);
                    }
                }

            }
            return dic;
        }
    }

}

#endregion

#region System.Data
namespace System.Data
{
    /// <summary>
    /// 数据库操作扩展
    /// 创建人：谢成磊
    /// 创建时间：2016/11/16
    /// </summary>
    public static class SystemDataBase
    {
        #region 根据数据库类型DataBaseType创建数据库连接
        /// <summary>
        /// 数据库类型，默认为ORACLE
        /// </summary>
        [Serializable]
        [Description("数据库类型")]
        public enum DataBaseType
        {
            [Description("Oracle数据库")]
            DEFAULT = 1,
            [Description("Oracle数据库")]
            ORACLE = 1,
            [Description("Access数据库")]
            MSACCESS = 2,
            [Description("SqlServer数据库")]
            MSSQL = 3,
            [Description("Sqlite数据库")]
            SQLITE = 4
        }

        [Serializable]
        public abstract class ConnectArgs
        {
            public abstract DataBaseType Type { get; }
        }

        [Serializable]
        [DisplayName("Oracle连接参数")]
        public class OracleConnectArgs : ConnectArgs
        {
            private int port = 1521;
            //private bool needPort = true;

            [Browsable(false)]
            public override DataBaseType Type
            {
                get { return DataBaseType.ORACLE; }
            }

            [Category("服务器"), DisplayName("主机地址"), Description("数据库服务所在主机地址")]
            public string HOST { get; set; }

            [Category("服务器"), DisplayName("主机端口"), Description("数据库服务所在主机端口")]
            public int PORT { get { return port; } set { port = value; } }
            //[Category("服务器"), DisplayName("是否需要端口")]
            //public bool NeedPort { get { return needPort; } set { needPort = value; } }

            [Category("数据库"), DisplayName("实例名"), Description("数据库服务实例名，例如：orcl")]
            public string SID { get; set; }

            [Category("验证"), DisplayName("用户"), Description("数据库用户名")]
            public string UserId { get; set; }

            [Category("验证"), DisplayName("密码"), Description("数据库用户登录密码"), PasswordPropertyText(true)]
            public string Password { get; set; }

            public override string ToString()
            {
                string format = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS =(PROTOCOL=TCP)(HOST={0})(PORT={1})))"
                             + "(CONNECT_DATA=(SID={2})(SERVER=DEDICATED)));User Id={3};Password={4};";
                string result = string.Format(format, this.HOST, this.PORT, this.SID, this.UserId, this.Password);
                return result;
            }

            public string GetDataSource()
            {
                string r = string.Format("{0}/{1}", this.HOST, this.SID);
                return r;
            }
        }

        [Serializable]
        [DisplayName("Sqlite连接参数")]
        public class SqliteConnectArgs : ConnectArgs
        {
            [Browsable(false)]
            public override DataBaseType Type
            {
                get { return DataBaseType.SQLITE; }
            }

            [Category("数据库"), DisplayName("文件路径"), Description("数据库文件路径")]
            public string FilePath { get; set; }

            [Category("验证"), DisplayName("数据库密码"), Description("数据库如果有密码，请在此处填写密码。"), PasswordPropertyText(true)]
            public string Password { get; set; }

            public override string ToString()
            {
                string format = string.Empty;
                string result = string.Empty;
                if (!string.IsNullOrWhiteSpace(this.Password))
                {
                    format = "Data Source={0};Version=3;Password={1}";
                    result = string.Format(format, this.FilePath, this.Password);
                }
                else
                {
                    format = "Data Source={0};Version=3";
                    result = string.Format(format, this.FilePath);
                }
                return result;
            }
        }

        [Serializable]
        [DisplayName("Access连接参数")]
        public class AccessConnectArgs : ConnectArgs
        {
            [Browsable(false)]
            public override DataBaseType Type
            {
                get { return DataBaseType.MSACCESS; }
            }

            [Category("数据库"), DisplayName("文件路径"), Description("数据库文件路径")]
            public string FilePath { get; set; }

            public override string ToString()
            {
                string format = "Provider=Microsoft.Ace.OleDb.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";
                string result = string.Format(format, this.FilePath);
                return result;
            }
        }

        [Serializable]
        [DisplayName("SqlServer连接参数")]
        public class MSSQLConnectArgs : ConnectArgs
        {
            [Browsable(false)]
            public override DataBaseType Type
            {
                get { return DataBaseType.MSSQL; }
            }

            [Category("数据库"), DisplayName("服务器"), Description("数据库服务器地址")]
            public string Server { get; set; }

            [Category("数据库"), DisplayName("数据库"), Description("数据库名称")]
            public string Database { get; set; }

            [Category("验证"), DisplayName("用户名"), Description("数据库用户")]
            public string UserId { get; set; }

            [Category("验证"), DisplayName("密码"), Description("数据库用户密码"), PasswordPropertyText(true)]
            public string Password { get; set; }

            public override string ToString()
            {
                string format = "Data Source = {0};Initial Catalog = {1};User Id = {2};Password = {3};";
                string result = string.Format(format, this.Server, this.Database, this.UserId, this.Password);
                return result;
            }
        }

        /// <summary>
        /// 获取连接参数实例
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static ConnectArgs GetConnectArg(this DataBaseType type)
        {
            ConnectArgs arg = null;
            switch (type)
            {
                case DataBaseType.ORACLE:
                    arg = new OracleConnectArgs();
                    break;
                case DataBaseType.MSSQL:
                    arg = new MSSQLConnectArgs();
                    break;
                case DataBaseType.MSACCESS:
                    arg = new AccessConnectArgs();
                    break;
                case DataBaseType.SQLITE:
                    arg = new SqliteConnectArgs();
                    break;
                default:
                    throw new NotImplementedException();
            }
            return arg;
        }

        public static string CreateConnectString(this DataBaseType type, ConnectArgs arg)
        {
            if (arg == null) throw new ArgumentNullException("ConnectArg为空。");
            if (!arg.Type.Equals(type)) throw new Exception("连接参数的数据库类型与实际数据库类型不一致。");

            return arg.ToString();
        }

        /// <summary>
        /// 创建数据库连接
        /// </summary>
        /// <param name="connectStr"></param>
        /// <param name="type">MSSQL|Oracle|MSAccess</param>
        /// <returns></returns>
        public static IDbConnection CreateDB(this DataBaseType type, string connectStr)
        {
            IDbConnection result = null;
            string DllPath = string.Empty;

            switch (type)
            {
                case DataBaseType.SQLITE:
                    DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System.Data.SQLite.dll");
                    result = DllPath.GetInstance("System.Data.SQLite.SQLiteConnection", true) as IDbConnection;
                    break;

                case DataBaseType.ORACLE:
                    //DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Oracle.ManagedDataAccess.dll");
                    //result = DllPath.GetInstance("Oracle.ManagedDataAccess.Client.OracleConnection", true) as IDbConnection;
                    DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Oracle.DataAccess.dll");
                    result = DllPath.GetInstance("Oracle.DataAccess.Client.OracleConnection", true) as IDbConnection;
                    break;

                case DataBaseType.MSSQL:
                    return new System.Data.SqlClient.SqlConnection(connectStr);

                case DataBaseType.MSACCESS:
                    return new System.Data.OleDb.OleDbConnection(connectStr);
            }

            if (result != null) result.ConnectionString = connectStr;
            return result;
        }

        /// <summary>
        /// 创建数据库数据适配器
        /// </summary>
        /// <param name="type">MSSQL|Oracle|MSAccess</param>
        /// <returns></returns>
        public static IDbDataAdapter CreateDBDataAdapter(this DataBaseType type)
        {
            string DllPath = string.Empty;
            switch (type)
            {
                case DataBaseType.SQLITE:
                    DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System.Data.SQLite.dll");
                    return DllPath.GetInstance("System.Data.SQLite.SQLiteDataAdapter", true) as IDbDataAdapter;

                case DataBaseType.ORACLE:
                    //DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Oracle.ManagedDataAccess.dll");
                    //return DllPath.GetInstance("Oracle.ManagedDataAccess.Client.OracleDataAdapter", true) as IDbDataAdapter;
                    DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Oracle.DataAccess.dll");
                    return DllPath.GetInstance("Oracle.DataAccess.Client.OracleDataAdapter", true) as IDbDataAdapter;

                case DataBaseType.MSSQL:
                    return new System.Data.SqlClient.SqlDataAdapter();

                case DataBaseType.MSACCESS:
                    return new System.Data.OleDb.OleDbDataAdapter();

                default:
                    return null;
            }
        }

        /// <summary>
        /// 根据数据库类型获取数据参数实例
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="value"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IDataParameter CreateDBDataParameter(this DataBaseType type, string paramName, object value)
        {
            string DllPath = string.Empty;
            IDataParameter result = null;
            switch (type)
            {
                case DataBaseType.SQLITE:
                    DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System.Data.SQLite.dll");
                    result = DllPath.GetInstance("System.Data.SQLite.SQLiteParameter", true) as IDataParameter;
                    break;

                case DataBaseType.ORACLE:
                    //DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Oracle.ManagedDataAccess.dll");
                    //result = DllPath.GetInstance("Oracle.ManagedDataAccess.Client.OracleParameter", true) as IDataParameter;
                    DllPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Oracle.DataAccess.dll");
                    result = DllPath.GetInstance("Oracle.DataAccess.Client.OracleParameter", true) as IDataParameter;

                    break;

                case DataBaseType.MSSQL:
                    return new System.Data.SqlClient.SqlParameter(paramName, value);

                case DataBaseType.MSACCESS:
                    return new System.Data.OleDb.OleDbParameter(paramName, value);

                default:
                    return null;
            }
            result.ParameterName = paramName;
            result.Value = value;
            return result;
        }

        /// <summary>
        /// 根据数据库类型获取数据库参数的前缀
        /// </summary>
        /// <param name="type">ORACLE/MSACCESS/MSSQL/</param>
        /// <returns></returns>
        public static char GetParameterPrefix(this DataBaseType type)
        {
            switch (type)
            {
                case DataBaseType.ORACLE:
                    return ':';

                case DataBaseType.SQLITE:
                case DataBaseType.MSACCESS:
                case DataBaseType.MSSQL:
                    return '@';

                default:
                    return char.MinValue;
            }
        }

        public static string ExistTableSql(this DataBaseType type, string tableName)
        {
            string sqlFormat = string.Empty;
            switch (type)
            {
                case SystemDataBase.DataBaseType.ORACLE:
                    sqlFormat = "select * from user_tables where table_name='{0}'";
                    break;
                case SystemDataBase.DataBaseType.SQLITE:
                    sqlFormat = "select * from sqlite_master where type='table' and name='{0}'";
                    break;
                case SystemDataBase.DataBaseType.MSSQL:
                    sqlFormat = "select * from sysobjects where Name='{0}'";
                    break;
                case SystemDataBase.DataBaseType.MSACCESS:
                    sqlFormat = "select * from MSysObjects where Name='{0}'";
                    break;
            }
            string existSql = string.Format(sqlFormat, tableName);
            return existSql;
        }
        #endregion

        #region IDbConnection
        /// <summary>
        /// 验证连接是否有效
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="message">返回的消息</param>
        /// <returns></returns>
        public static bool Validate(this IDbConnection dbConn, out string message)
        {
            bool result = false;
            message = "连接成功！";
            try
            {
                dbConn.Open();
                if (dbConn.State == ConnectionState.Open)
                    result = true;
                else
                    message = "打开连接失败！";
            }
            catch (Exception ex)
            {
                message = "连接失败！原因:" + ex.Message;
            }
            finally
            {
                if (dbConn.State != ConnectionState.Closed) dbConn.Close();
            }
            return result;
        }

        /// <summary>
        /// 打开数据库连接
        /// </summary>
        /// <param name="dbConn"></param>
        public static void OpenDB(this IDbConnection dbConn)
        {
            switch (dbConn.State)
            {
                case ConnectionState.Closed:
                    dbConn.Open();
                    break;
                case ConnectionState.Open:
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 关闭数据库连接
        /// </summary>
        /// <param name="dbConn"></param>
        public static void CloseDB(this IDbConnection dbConn)
        {
            if (dbConn.State != ConnectionState.Closed) dbConn.Close();
        }

        /// <summary>
        /// 获得表字段及其属性 FIELDNAME DATATYPE DATALENGTH
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="tableName"></param>
        /// <param name="dbType"></param>
        /// <returns></returns>
        public static DataTable GetFieldsTable(this IDbConnection dbConn, string tableName, DataBaseType dbType = DataBaseType.ORACLE)
        {
            string sql = "";

            switch (dbType)
            {
                case DataBaseType.ORACLE:
                    sql = "select COLUMN_NAME as FIELDNAME,DATA_TYPE as DATATYPE,DATA_LENGTH as DATALENGTH from  user_tab_cols  where table_name=upper('{0}')";
                    break;
                case DataBaseType.MSACCESS:
                    DataTable schema = (dbConn as System.Data.OleDb.OleDbConnection).GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, new object[] { null, null, tableName, null });
                    DataTable result = CreateDbFieldsTable();

                    foreach (DataRow dr in schema.Rows)
                    {
                        DataRow drN = result.NewRow();
                        drN["FIELDNAME"] = dr["COLUMN_NAME"].ToString();
                        drN["DATATYPE"] = dr["DATA_TYPE"].ToString();
                        drN["DATALENGTH"] = int.Parse(dr["CHARACTER_MAXIMUM_LENGTH"].ToString());
                        result.Rows.Add(drN);
                    }

                    return result;
                case DataBaseType.MSSQL:
                    sql = "SELECT [sc].[name] AS FIELDNAME, "
                        + "[st].[name] AS DATATYPE, "
                        + "[sc].[length] AS DATALENGTH "
                        + "FROM [syscolumns] AS [sc] "
                        + "INNER JOIN [sysobjects] AS [so] ON [so].[id] = [sc].[id] "
                        + "INNER JOIN [systypes] AS [st] ON [sc].[xtype] = [st].[xtype] "
                        + "WHERE [so].[name] = '{0}'";
                    break;
                case DataBaseType.SQLITE:
                    sql = "select sql from sqlite_master where tbl_name = '" + tableName + "' and type='table'";
                    IDbCommand sqliteCmd = dbConn.CreateDbCommand(sql);
                    string tableSql = sqliteCmd.ExecuteScalar().ToString();
                    DataTable sqliteDt = CreateDbFieldsTable();
                    string[] fieldStrs = tableSql.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    if (fieldStrs.Length == 1)
                    {
                        fieldStrs = fieldStrs[0].Split(new string[] { ",", "(", ")" }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < fieldStrs.Length; i++)
                        {
                            string fldName = string.Empty;
                            string fldType = string.Empty;
                            int len = -1;
                            string fieldStr = fieldStrs[i];
                            if (fieldStr.ToUpper().StartsWith("CREATE TABLE")) continue;

                            fieldStr = fieldStr.Replace("NOT NULL", "");
                            fieldStr = fieldStr.Replace("PRIMARY KEY AUTOINCREMENT", "");
                            fieldStr = fieldStr.Trim();

                            string[] kv = fieldStr.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            if (!(kv.Length >= 2)) continue;
                            fldName = kv[0];
                            fldType = kv[1];
                            len = -1;
                            DataRow sqlDr = sqliteDt.NewRow();
                            sqlDr["FIELDNAME"] = fldName;
                            sqlDr["DATATYPE"] = fldType;
                            sqlDr["DATALENGTH"] = len;
                            sqliteDt.Rows.Add(sqlDr);
                        }
                    }
                    else
                        for (int i = 0; i < fieldStrs.Length; i++)
                        {
                            string fldName = string.Empty;
                            string fldType = string.Empty;
                            int len = -1;

                            #region 数据处理
                            string fieldStr = fieldStrs[i].Trim();
                            if (!fieldStr.StartsWith("[")) continue;
                            if (fieldStr.EndsWith(")"))
                            {
                                fieldStr = fieldStr.Substring(0, fieldStr.Length - 1);
                                fieldStr = fieldStr + ",";
                            }
                            fieldStr = fieldStr.Replace("NOT NULL", "");

                            string[] kv = fieldStr.Split(new char[] { '[', ']', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            if (!(kv.Length >= 2)) continue;
                            fldName = kv[0];
                            fldType = kv[1];
                            len = -1;
                            if (fldType.Contains("("))
                            {
                                string[] tl = fldType.Split(new char[] { '(', ')', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                if (tl.Length >= 2)
                                {
                                    fldType = tl[0];
                                    len = int.Parse(tl[1]);
                                }
                            }
                            #endregion

                            DataRow sqlDr = sqliteDt.NewRow();
                            sqlDr["FIELDNAME"] = fldName;
                            sqlDr["DATATYPE"] = fldType;
                            sqlDr["DATALENGTH"] = len;
                            sqliteDt.Rows.Add(sqlDr);
                        }

                    return sqliteDt;
                default:
                    sql = ""; break;
            }
            if (sql == "") return null;
            sql = string.Format(sql, tableName);

            IDbDataAdapter adpt = dbType.CreateDBDataAdapter();
            DataSet ds = adpt.GetDataSet(dbConn, sql);
            if (ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                dt.PrimaryKey = new DataColumn[] { dt.Columns["FIELDNAME"] };
                return dt;
            }
            return null;
        }
        private static DataTable CreateDbFieldsTable()
        {
            DataTable result = new DataTable();
            result.Columns.Add(new DataColumn("FIELDNAME", typeof(string)));
            result.Columns.Add(new DataColumn("DATATYPE", typeof(string)));
            result.Columns.Add(new DataColumn("DATALENGTH", typeof(int)));

            result.PrimaryKey = new DataColumn[] { result.Columns["FIELDNAME"] };
            return result;
        }

        /// <summary>
        /// 创建数据库命令接口
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL">查询语句</param>
        /// <returns></returns>
        public static IDbCommand CreateDbCommand(this IDbConnection dbConn, string SQL)
        {
            return dbConn.CreateDbCommand(SQL, -1);
        }

        /// <summary>
        /// 创建数据库命令接口
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL"></param>
        /// <param name="timeOut">查询语句</param>
        /// <returns></returns>
        public static IDbCommand CreateDbCommand(this IDbConnection dbConn, string SQL, int timeOut)
        {
            IDbCommand cmd = dbConn.CreateCommand();
            cmd.CommandText = SQL;
            if (timeOut >= 0) cmd.CommandTimeout = timeOut;
            return cmd;
        }

        /// <summary>
        /// 获得DataReader
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL"></param>
        /// <returns></returns>
        public static IDataReader ExecuteReader(this IDbConnection dbConn, string SQL)
        {
            dbConn.OpenDB();
            IDataReader result = dbConn.CreateDbCommand(SQL).ExecuteReader();
            return result;
        }

        /// <summary>
        /// 获取下一个ID
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="tablename"></param>
        /// <param name="fieldname"></param>
        /// <param name="startvalue"></param>
        /// <param name="stepvalue"></param>
        /// <returns></returns>
        public static int NextID(
            this IDbConnection dbConn,
            string tablename,
            string fieldname,
            int startvalue = 10000,
            int stepvalue = 1)
        {
            IDbCommand cmd = dbConn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "XCL_GETNEXTID";

            IDbDataParameter para1 = cmd.CreateParameter();//--表名称
            para1.Direction = ParameterDirection.Input;
            para1.DbType = DbType.String;
            para1.Value = tablename;
            para1.ParameterName = "tablename";

            IDbDataParameter para2 = cmd.CreateParameter();//--字段名称
            para2.Direction = ParameterDirection.Input;
            para2.DbType = DbType.String;
            para2.Value = fieldname;
            para2.ParameterName = "fieldname";

            IDbDataParameter para3 = cmd.CreateParameter();//--初始值
            para3.Direction = ParameterDirection.Input;
            para3.DbType = DbType.Int32;
            para3.Value = startvalue;
            para3.ParameterName = "startvalue";

            IDbDataParameter para4 = cmd.CreateParameter();//--步长
            para4.Direction = ParameterDirection.Input;
            para4.DbType = DbType.Int32;
            para4.Value = stepvalue;
            para4.ParameterName = "stepvalue";

            int result = -1;
            IDbDataParameter paraOut = cmd.CreateParameter();//--步长
            paraOut.Direction = ParameterDirection.Output;
            paraOut.DbType = DbType.Int32;
            paraOut.Value = result;
            paraOut.ParameterName = "result";

            cmd.Parameters.Add(para1);
            cmd.Parameters.Add(para2);
            cmd.Parameters.Add(para3);
            cmd.Parameters.Add(para4);
            cmd.Parameters.Add(paraOut);

            dbConn.OpenDB();
            cmd.ExecuteNonQuery();
            dbConn.CloseDB();
            result = Convert.ToInt32(paraOut.Value);

            return result;
        }

        /// <summary>
        /// 判断表是否存在
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="tableName">表名称</param>
        /// <returns></returns>
        public static bool HasTable(this IDbConnection dbConn, string tableName)
        {
            string sql = string.Format("SELECT * FROM TAB WHERE TABTYPE='TABLE' and TNAME='{0}'", tableName.ToUpper());
            return dbConn.HasRecord(sql);
        }

        /// <summary>
        /// 判断视图是否存在
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="viewName">视图名</param>
        /// <returns></returns>
        public static bool HasView(this IDbConnection dbConn, string viewName)
        {
            string sql = string.Format("SELECT * FROM TAB WHERE TABTYPE='VIEW' and TNAME='{0}'", viewName);
            return dbConn.HasRecord(sql);
        }

        /// <summary>
        /// 判断字段是否存在
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="tableName">表名称</param>
        /// <param name="fieldName">字段名称</param>
        /// <returns></returns>
        public static bool HasField(this IDbConnection dbConn, string tableName, string fieldName)
        {
            string fieldCountSql = string.Format(
                "select count(*) from(select COLUMN_NAME as FIELDNAME from user_tab_cols where table_name = upper('{0}')) where FIELDNAME = '{1}'",
                tableName, fieldName.ToUpper());

            string msg = "";
            Decimal fldNum = (Decimal)dbConn.ExecuteScalar(fieldCountSql, out msg);
            bool HasField = fldNum > 0;

            return HasField;
        }

        private static string getSidxName(string tableName) { return "SIDX_" + tableName; }

        /// <summary>
        /// 创建Oracle空间数据库表
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="tableName">表名称</param>
        /// <param name="fields">字段列表</param>
        /// <param name="srid">SRID</param>
        /// <returns></returns>
        public static bool CreateOracleSpatialTable(this IDbConnection dbConn, string tableName, Dictionary<string, string> fields, int srid)
        {
            string _tableName = tableName.ToUpper().Trim();//表名称一定要大写，否则在创建空间索引的时候报表或视图不存在。
            string _shape = "";

            if (fields == null) fields = new Dictionary<string, string>();

            bool hasShapeField = false;
            List<string> delList = new List<string>();
            foreach (string key in fields.Keys)
            {
                string val = fields[key];
                if (val != null)
                {
                    if (val.ToUpper().Contains("SDO_GEOMETRY"))
                    {
                        delList.Add(key);
                        _shape = key;
                        hasShapeField = true;
                    }
                }
            }
            foreach (string key in delList)
            {
                fields.Remove(key);
            }

            if (hasShapeField)
                fields.Add(_shape, "MDSYS.SDO_GEOMETRY");
            else
                throw new MissingFieldException("创建空间数据表时没有找到空间字段");

            // 创建空间数据表
            string fieldString = "";
            if (fields != null)
                foreach (KeyValuePair<string, string> kv in fields)
                    fieldString += string.Format("{0} {1},", kv.Key, kv.Value);
            fieldString = fieldString.TrimEnd(',');

            string createTableSql = string.Format("CREATE TABLE {0} ({1})",
                _tableName, fieldString);

            bool tableCreated = dbConn.ExecuteNonQuery(createTableSql) > 0 ? true : false;

            // 插入用户SDO元数据表记录
            string metaSql = string.Format("SELECT * FROM USER_SDO_GEOM_METADATA WHERE TABLE_NAME='{0}'",
                _tableName);
            bool hasMeta = dbConn.HasRecord(metaSql);
            if (hasMeta)
            {
                string delMeta = string.Format("DELETE FROM USER_SDO_GEOM_METADATA WHERE TABLE_NAME='{0}'", tableName.ToUpper());
                int DRPMETA = dbConn.ExecuteNonQuery(delMeta);
            }

            string sridStr = srid.ToString();// NULL
            string metaInsert = string.Format("INSERT INTO USER_SDO_GEOM_METADATA (\"TABLE_NAME\", \"COLUMN_NAME\", \"DIMINFO\", \"SRID\") VALUES ('{0}','{1}',{2},{3})",
                _tableName,
                _shape,
                "MDSYS.SDO_DIM_ARRAY(MDSYS.SDO_DIM_ELEMENT('X',-180,180,0.0005),MDSYS.SDO_DIM_ELEMENT('Y',-90,90,0.0005))",
                sridStr);
            hasMeta = dbConn.ExecuteNonQuery(metaInsert) > 0 ? true : false;

            // 创建空间索引
            if (hasMeta)
            {
                string idxName = getSidxName(_tableName);
                string hasIdx = string.Format("SELECT * FROM SYS.USER_INDEXES WHERE TABLE_NAME='{0}' AND INDEX_NAME='{1}'", tableName, idxName);
                if (dbConn.HasRecord(hasIdx))
                {
                    string drpIdx = string.Format("DROP INDEX \"{0}\"", idxName);
                    int DRP_SIDX = dbConn.ExecuteNonQuery(drpIdx);
                }
                string createIdx = string.Format("CREATE INDEX  \"{0}\" ON \"{1}\" (\"{2}\") INDEXTYPE IS \"MDSYS\".\"SPATIAL_INDEX\"",
                    idxName, _tableName, _shape.ToUpper());
                try
                {
                    dbConn.ExecuteNonQuery(createIdx);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// 删除Oracle空间数据库表
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="tableName">表名称</param>
        /// <returns></returns>
        public static bool DeleteOracleSpatialTable(this IDbConnection dbConn, string tableName)
        {
            try
            {
                string _tableName = tableName.ToUpper().Trim();
                // 第一步 删除空间数据表
                string dropTable = string.Format("DROP TABLE {0} CASCADE CONSTRAINTS", _tableName);
                int DRPTABLE = dbConn.ExecuteNonQuery(dropTable);

                // 第二步 删除用户SDO元数据表记录
                string delMeta = string.Format("DELETE FROM USER_SDO_GEOM_METADATA WHERE TABLE_NAME='{0}'", _tableName);
                int DELMETA = dbConn.ExecuteNonQuery(delMeta);

                // 删除空间索引 空间索引已经被第一步删除
                //string idxName = getSidxName(tableName);
                //string drpIdx = string.Format("DROP INDEX \"{0}\"", idxName);
                //int DRP_SIDX = dbConn.ExecuteNonQuery(drpIdx);

                return true;
            }
            catch
            {
            }
            return false;
        }

        /// <summary>
        /// 判断数据库中是否含有查询语句描述的记录,使用后关闭连接
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="selectSQL">选择语句</param>
        /// <returns></returns>
        public static bool HasRecord(this IDbConnection dbConn, string selectSQL)
        {
            return dbConn.HasRecord(selectSQL, false);
        }

        /// <summary>
        /// 判断数据库中是否含有查询语句描述的【唯一记录】,使用后关闭连接
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="selectSQL">选择语句</param>
        /// <returns></returns>
        public static bool HasRecordOnlyOne(this IDbConnection dbConn, string selectSQL)
        {
            return dbConn.HasRecordOnlyOne(selectSQL, false);
        }

        /// <summary>
        /// 判断数据库中是否含有查询语句描述的记录
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="selectSQL">选择语句</param>
        /// <param name="notClose">true:不关闭数据库连接，false:关闭数据库连接</param>
        /// <returns></returns>
        public static bool HasRecord(this IDbConnection dbConn, string selectSQL, bool notClose)
        {
            string head = selectSQL.Trim().Substring(0, 6);
            if (!(head.ToLower() == "select")) throw new SyntaxErrorException("HasRecord方法只接受SELECT语句");

            dbConn.OpenDB();
            IDbCommand cmd = dbConn.CreateCommand();
            cmd.CommandText = selectSQL;

            IDataReader reader = cmd.ExecuteReader();
            bool result = reader.Read();
            reader.Close();

            if (!notClose) dbConn.Close();

            return result;
        }

        /// <summary>
        /// 判断数据库中是否含有查询语句描述的【唯一记录】
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="selectSQL">选择语句</param>
        /// <param name="notClose">true:不关闭数据库连接，false:关闭数据库连接</param>
        /// <returns></returns>
        public static bool HasRecordOnlyOne(this IDbConnection dbConn, string selectSQL, bool notClose)
        {
            string head = selectSQL.Trim().Substring(0, 6);
            if (!(head.ToLower() == "select")) throw new SyntaxErrorException("HasRecord方法只接受SELECT语句");

            dbConn.OpenDB();
            IDbCommand cmd = dbConn.CreateCommand();
            cmd.CommandText = selectSQL;

            IDataReader reader = cmd.ExecuteReader();
            bool result1 = reader.Read();
            bool result2 = true;
            if (result1) result2 = reader.Read();
            reader.Close();

            if (!notClose) dbConn.Close();
            bool result = false;
            if (result1 == true && result2 == false)
                result = true;
            else
                result = false;
            return result;
        }

        /// <summary>
        /// 针对 .NET Framework 数据提供程序的 Connection 对象执行 SQL 语句，并返回受影响的行数。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL"></param>
        /// <param name="dic"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(this IDbConnection dbConn, string SQL, IDictionary<object, object> dic)
        {
            return dbConn.ExecuteNonQuery(SQL, dic, false);
        }

        /// <summary>
        /// 针对 .NET Framework 数据提供程序的 Connection 对象执行 SQL 语句，并返回受影响的行数。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL"></param>
        /// <param name="dic"></param>
        /// <param name="notClose"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(this IDbConnection dbConn, string SQL, IDictionary<object, object> dic, bool notClose)
        {
            dbConn.OpenDB();
            IDbCommand cmd = dbConn.CreateDbCommand(SQL);
            cmd.SetParameters(dic);
            int result = cmd.ExecuteNonQuery();
            if (!notClose) dbConn.Close();
            return result;
        }

        /// <summary>
        /// 针对 .NET Framework 数据提供程序的 Connection 对象执行 SQL 语句，并返回受影响的行数。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(this IDbConnection dbConn, string SQL)
        {
            return dbConn.ExecuteNonQuery(SQL, false);
        }

        /// <summary>
        /// 针对 .NET Framework 数据提供程序的 Connection 对象执行 SQL 语句，并返回受影响的行数。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="SQL"></param>
        /// <param name="notClose"></param>
        /// <returns></returns>
        public static int ExecuteNonQuery(this IDbConnection dbConn, string SQL, bool notClose)
        {
            dbConn.OpenDB();
            IDbCommand cmd = dbConn.CreateCommand();
            cmd.CommandText = SQL;
            int result = cmd.ExecuteNonQuery();
            if (!notClose) dbConn.Close();
            return result;
        }

        /// <summary>
        /// 执行查询，并返回查询所返回的结果集中第一行的第一列。忽略额外的列或行。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="Sql"></param>
        /// <param name="Msg"></param>
        /// <returns></returns>
        public static object ExecuteScalar(this IDbConnection dbConn, string Sql, out string Msg)
        {
            dbConn.OpenDB();

            IDbCommand cmd = dbConn.CreateDbCommand(Sql);
            IDbTransaction trans = dbConn.BeginTransaction();

            object obj = null;
            try
            {
                cmd.Transaction = trans;
                obj = cmd.ExecuteScalar();
                trans.Commit();
                Msg = string.Format("操作成功");
            }
            catch (Exception ex)
            {
                Msg = "操作失败：" + ex.Message;
                trans.Rollback();
            }
            finally
            {
                dbConn.CloseDB();
            }
            return obj;
        }

        /// <summary>
        /// 针对 .NET Framework 数据提供程序的 Connection 对象执行 SQL 语句，并返回受影响的行数。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="Sql"></param>
        /// <param name="Msg"></param>
        /// <returns></returns>
        public static int Execute(this IDbConnection dbConn, string Sql)
        {
            dbConn.OpenDB();

            IDbCommand cmd = dbConn.CreateDbCommand(Sql);
            IDbTransaction trans = dbConn.BeginTransaction();

            int success = -1;
            try
            {
                cmd.Transaction = trans;
                int count = cmd.ExecuteNonQuery();
                success = count;
                trans.Commit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                trans.Rollback();
            }
            finally
            {
                dbConn.CloseDB();
            }
            return success;
        }

        /// <summary>
        /// 针对 .NET Framework 数据提供程序的 Connection 对象执行 SQL 语句，受影响的行数大于0返回true。
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="Sql"></param>
        /// <param name="Msg"></param>
        /// <returns></returns>
        public static bool Execute(this IDbConnection dbConn, string Sql, out string Msg)
        {
            dbConn.OpenDB();

            IDbCommand cmd = dbConn.CreateDbCommand(Sql);
            IDbTransaction trans = dbConn.BeginTransaction();

            bool success = false;
            try
            {
                cmd.Transaction = trans;
                int count = cmd.ExecuteNonQuery();
                success = count > 0 ? true : false;
                trans.Commit();
                Msg = string.Format("操作成功：影响了{0}行。", count);
            }
            catch (Exception ex)
            {
                Msg = "操作失败：" + ex.Message;
                trans.Rollback();
            }
            finally
            {
                dbConn.CloseDB();
            }
            return success;
        }

        /// <summary>
        /// 获取记录个数
        /// </summary>
        /// <param name="dbConn"></param>
        /// <param name="dbType"></param>
        /// <param name="tableName">表名称/视图名称</param>
        /// <returns></returns>
        public static int RecordCount(this IDbConnection dbConn, SystemDataBase.DataBaseType dbType, string tableName)
        {
            int totalRowNum = -1;

            string sql = dbType.GetSelectSQL(tableName, null, "count(*)");
            dbConn.OpenDB();
            IDbTransaction trans = dbConn.BeginTransaction();
            try
            {
                IDbCommand cmd = dbConn.CreateDbCommand(sql);
                cmd.Transaction = trans;
                object obj = cmd.ExecuteScalar();
                totalRowNum = Convert.ToInt32(obj);
            }
            catch
            {
                totalRowNum = -1;
            }
            finally
            {
                dbConn.CloseDB();
            }
            return totalRowNum;
        }

        /// <summary>
        /// 获取唯一值
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="tableName"></param>
        /// <param name="fieldName"></param>
        /// <param name="whereClause"></param>
        /// <returns></returns>
        public static List<object> UniqueValue(this IDbConnection conn, string tableName, string fieldName, string whereClause = "")
        {
            string sql = string.Format("select distinct({0}) as uniquevalue from {1}", fieldName, tableName);
            if (!string.IsNullOrWhiteSpace(whereClause))
                sql = string.Format("{0} where {1}", sql, whereClause);

            List<object> list = new List<object>();

            if (!conn.HasTable(tableName))
                return list;

            conn.OpenDB();

            IDbCommand cmd = conn.CreateDbCommand(sql);
            IDataReader reader = cmd.ExecuteReader();
            try
            {
                while (reader.Read())
                {
                    object val = reader["uniquevalue"];
                    list.Add(val);
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                reader.Close();
                conn.CloseDB();
            }

            return list;
        }

        #endregion

        #region IDbCommand
        /// <summary>
        /// 给IDbCommand的Parameters赋值
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="KVs"></param>
        public static void SetParameters(this IDbCommand cmd, IDictionary<object, object> KVs)
        {
            if (KVs == null) throw new ArgumentNullException("KVs");
            foreach (object key in KVs.Keys)
            {
                int index = cmd.Parameters.Add(key);
                cmd.Parameters[index] = KVs[key];
            }
        }
        /// <summary>
        /// 匹配对象和数据库表字段，并返回字段及其对应的值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="fieldTable"></param>
        /// <param name="dbType"></param>
        /// <returns></returns>
        public static IList<IDataParameter> ParseDbParameterCollection(this object obj, DataTable fieldTable, DataBaseType dbType)
        {
            IList<IDataParameter> result = new List<IDataParameter>();

            char prefix = dbType.GetParameterPrefix();

            PropertyInfo[] props = obj.GetType().GetProperties();
            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                string paramName = prefix + prop.Name + '_';
                DataRow[] drs = fieldTable.Select(string.Format("FIELDNAME = '{0}'", prop.Name));
                if (drs.Length > 0)
                {
                    object val = prop.GetValue(obj, null);
                    if (val is Enum) val = (int)val;
                    if (val is bool) val = (bool)val ? 1 : 0;

                    IDataParameter param = dbType.CreateDBDataParameter(paramName, val);
                    if (prop.PropertyType.Name.ToUpper() == "SDOGEOMETRY")
                    {
                        param.DbType = DbType.Object;
                        param.SetPropertyValue("UdtTypeName", "MDSYS.SDO_GEOMETRY");
                    }
                    result.Add(param);
                }
            }
            return result;
        }
        public static IList<IDataParameter> ParseDbParameterCollection2(this object obj, DataTable fieldTable, DataBaseType dbType, Dictionary<string, string> map)
        {
            IList<IDataParameter> result = new List<IDataParameter>();

            char prefix = dbType.GetParameterPrefix();

            PropertyInfo[] props = obj.GetType().GetProperties();
            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                string propName = prop.Name;
                if (map.Keys.Contains(propName)) propName = map[propName];

                string paramName = prefix + propName + '_';
                DataRow[] drs = fieldTable.Select(string.Format("FIELDNAME = '{0}'", propName));
                if (drs.Length > 0)
                {
                    object val = prop.GetValue(obj, null);

                    IDataParameter param = dbType.CreateDBDataParameter(paramName, val);
                    if (prop.PropertyType.Name.ToUpper() == "SDOGEOMETRY")
                    {
                        param.DbType = DbType.Object;
                        param.SetPropertyValue("UdtTypeName", "MDSYS.SDO_GEOMETRY");
                    }

                    result.Add(param);
                }
            }
            return result;
        }

        /// <summary>
        /// 给对象属性赋值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="dataRow"></param>
        public static void GetValueFrom(this object obj, DataRow dataRow)
        {
            PropertyInfo[] props = obj.GetType().GetProperties();
            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                int idx = dataRow.Table.Columns.IndexOf(prop.Name);
                if (idx >= 0)
                {
                    object val = dataRow[idx];
                    if (val.GetType().Equals(typeof(DBNull))) continue;
                    obj.SetPropertyValue(prop.Name, val);
                }
            }
        }
        /// <summary>
        /// 给对象属性赋值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="dataRow"></param>
        public static void GetValueFrom2(this object obj, DataRow dataRow, Dictionary<string, string> map)
        {
            PropertyInfo[] props = obj.GetType().GetProperties();
            if (map == null) map = new Dictionary<string, string>();
            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                if (!map.Keys.Contains(prop.Name)) continue;
                string fieldName = map[prop.Name];

                int idx = dataRow.Table.Columns.IndexOf(fieldName);
                if (idx >= 0)
                {
                    object val = dataRow[idx];
                    if (val.GetType().Equals(typeof(DBNull))) continue;

                    Type valType = val.GetType();
                    if (valType.FullName == "System.__ComObject")
                    {
                        continue;
                    }
                    obj.SetPropertyValue(prop.Name, val);
                }
            }

        }

        public static void SetValueTo(this object obj,ref DataRow dataRow)
        {
            System.Reflection.PropertyInfo[] props = obj.GetType().GetProperties();

            for (int i = 0; i < props.Length; i++)
            {
                System.Reflection.PropertyInfo prop = props[i];
                int idx = dataRow.Table.Columns.IndexOf(prop.Name);
                dataRow[idx] = prop.GetValue(obj, null);
            }
        }
        #endregion

        #region IDbDataAdapter
        /// <summary>
        /// 获得dataset
        /// </summary>
        /// <param name="dbAdapter"></param>
        /// <param name="dbConnection"></param>
        /// <param name="SQL"></param>
        /// <returns></returns>
        public static DataSet GetDataSet(this IDbDataAdapter dbAdapter, IDbConnection dbConnection, string SQL)
        {
            string sql = SQL.Trim();
            string head = SQL.TrimStart().ToUpper().Substring(0, 6);

            DataSet result = new DataSet();
            dbConnection.OpenDB();
            IDbCommand cmd = dbConnection.CreateDbCommand(sql);
            switch (head)
            {
                case "INSERT": dbAdapter.InsertCommand = cmd; break;
                case "DELETE": dbAdapter.DeleteCommand = cmd; break;
                case "UPDATE": dbAdapter.UpdateCommand = cmd; break;
                case "SELECT": dbAdapter.SelectCommand = cmd; break;
                default: throw new Exception("只支持增删改查操作");
            }
            dbAdapter.Fill(result);

            dbConnection.CloseDB();
            return result;
        }
        /// <summary>
        /// 获得DataTable
        /// </summary>
        /// <param name="dbAdapter"></param>
        /// <param name="dbConnection"></param>
        /// <param name="SQL"></param>
        /// <returns></returns>
        public static DataTable GetTable(this IDbDataAdapter dbAdapter, IDbConnection dbConnection, string SQL)
        {
            DataTable tb = null;
            DataSet cityDs = dbAdapter.GetDataSet(dbConnection, SQL);
            if (cityDs.Tables.Count > 0) tb = cityDs.Tables[0];
            if (tb == null) tb = new DataTable();
            return tb;
        }
        #endregion

        #region IDataParameterCollection

        public static string GetPrefixName(this char prefix, string fieldName)
        {
            return prefix + fieldName + "_";
        }

        public static void AddCollect(this IDataParameterCollection collect, IList<IDataParameter> other)
        {
            foreach (IDataParameter para in other)
            {
                collect.Add(para);
            }
        }

        public static void AddParameters(this IDataParameterCollection collect,
            DataTable fields, HttpRequest request,
            DataBaseType dbType)
        {
            string fieldName = "FIELDNAME", dataType = "DATATYPE";
            char prefix = dbType.GetParameterPrefix();
            foreach (DataRow dr in fields.Rows)
            {
                string name = dr[fieldName].ToString().Trim(),
                       type = dr[dataType].ToString().Trim(),
                       prefix_name = prefix.GetPrefixName(name),
                       valueStr = HttpUtility.UrlDecode(request[name]);

                object obj = DBNull.Value;
                switch (type.ToUpper())
                {
                    case "NUMBER":
                    case "INT":
                    case "TINYINT":
                        if (valueStr != null) obj = Convert.ToInt32(valueStr);
                        break;

                    case "DATE":
                    case "SMALLDATETIME":
                    case "DATETIME":
                        if (valueStr != null) obj = Convert.ToDateTime(valueStr);
                        break;
                    default:
                        if (valueStr != null) obj = valueStr;
                        break;
                }
                if (obj == DBNull.Value) continue;
                IDataParameter para = dbType.CreateDBDataParameter(prefix_name.ToUpper(), obj);
                collect.Add(para);
            }
        }
        #endregion

        #region DataTable
        public static DataTable NewTableFiltPkBy(this DataTable table, IList<string> filter)
        {
            DataTable dt = table.Clone();
            foreach (string s in filter)
            {
                DataRow dr_ = table.Rows.Find(s);
                if (dr_ == null) continue;
                DataRow dr = dt.NewRow();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    dr[i] = dr_[i];
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        #endregion

        #region SQLBuilder

        /// <summary>
        /// 获取下一个记录的SQL语句
        /// </summary>
        /// <param name="dbType"></param>
        /// <param name="tableName"></param>
        /// <param name="curRowIndex"></param>
        /// <param name="whereClause"></param>
        /// <returns></returns>
        public static string GetNextItemSQL(this DataBaseType dbType, string tableName, int curRowIndex, string whereClause, string fieldsFormat = "*")
        {
            string sqlTable = string.Empty;

            string tablePage = string.Empty;
            switch (dbType)
            {
                case DataBaseType.SQLITE:
                    sqlTable = string.Format("({0})", dbType.GetSelectSQL(tableName, whereClause));
                    tablePage = string.Format("select * from {0} limit 1 offset {1}", sqlTable, curRowIndex);
                    break;
                case DataBaseType.DEFAULT:
                    sqlTable = string.Format("({0}) bb", dbType.GetSelectSQL(tableName + " tt", whereClause, string.Format(fieldsFormat, "tt.")));
                    tablePage = string.Format("(SELECT aa.*, ROWNUM rowindex FROM (SELECT * FROM {0}) aa)", sqlTable);
                    tablePage = dbType.GetSelectSQL(tablePage, string.Format("rowindex ={0}", curRowIndex));
                    break;
                case DataBaseType.MSACCESS:
                case DataBaseType.MSSQL:
                    //sqlTable = string.Format("({0}) _b", dbType.GetSelectSQL(tableName, whereClause, "*,%%lockres%% RID"));
                    sqlTable = string.Format("({0}) _b", dbType.GetSelectSQL(tableName, whereClause, "*,ROW_NUMBER() OVER(ORDER BY id) RID"));
                    tablePage = string.Format("select top 1 * from ( select * from {0}) _a where RID not in ( select top {1} RID from {0})",
                        sqlTable, curRowIndex);
                    break;
            }
            return tablePage;
        }

        public static string GetSelectSQL(this DataBaseType dbType, string tableName, string whereClause = "", string selectedFields = "*")
        {
            string _selectedFields = selectedFields;
            if (string.IsNullOrWhiteSpace(_selectedFields))
            {
                _selectedFields = "*";
            }
            string result = "select {0} from {1} where {2}";
            if (string.IsNullOrWhiteSpace(whereClause))
            {
                return string.Format("select {0} from {1}", _selectedFields, tableName);
            }
            return string.Format(result, _selectedFields, tableName, whereClause);
        }

        public static string GetDeleteSQL(this DataBaseType dbType, string tableName, string whereClause)
        {
            string result = "delete from {0} where {1}";
            return string.Format(result, tableName, whereClause);
        }

        public static string GetUpdateSQL(this DataBaseType dbType, DataTable fields, string tableName, string whereClause)
        {
            string result = "";
            string fieldsStr = "";

            string fieldName = "FIELDNAME";
            string dataType = "DATATYPE";
            string dataLength = "DATALENGTH";

            if (fields == null || fields.Rows.Count == 0) return result;
            if (!fields.Columns.Contains(fieldName)
             || !fields.Columns.Contains(dataType)
             || !fields.Columns.Contains(dataLength))
                return result;

            result = "update {0} set {1} where {2}";
            char prefix = dbType.GetParameterPrefix();
            string format = "\"{0}\"=" + prefix + "{0}_,";
            if (dbType == DataBaseType.MSSQL || dbType == DataBaseType.MSACCESS) format = "[{0}]=" + prefix + "{0}_,";

            foreach (DataRow field in fields.Rows)
            {
                string name = field[fieldName].ToString();
                fieldsStr += string.Format(format, name.ToUpper());
            }

            result = string.Format(result, tableName, fieldsStr.TrimEnd(','), whereClause);

            return result;
        }

        public static string GetInsertSQL(this DataBaseType dbType, DataTable fields, string tableName)
        {
            string result = "";

            string fieldsStr = "";
            string fieldsPara = "";

            string fieldName = "FIELDNAME";
            string dataType = "DATATYPE";
            string dataLength = "DATALENGTH";

            if (fields == null || fields.Rows.Count == 0) return result;
            if (!fields.Columns.Contains(fieldName)
             || !fields.Columns.Contains(dataType)
             || !fields.Columns.Contains(dataLength))
                return result;

            result = "insert into {0}({1}) values({2})";
            char prefix = dbType.GetParameterPrefix();
            string format1 = "\"{0}\",",
                   format2 = prefix + "{0}_,";
            if (dbType == DataBaseType.MSSQL || dbType == DataBaseType.MSACCESS) format1 = "[{0}],";

            foreach (DataRow field in fields.Rows)
            {
                string name = field[fieldName].ToString();
                fieldsStr += string.Format(format1, name.ToUpper());
                fieldsPara += string.Format(format2, name.ToUpper());
            }

            result = string.Format(result, tableName, fieldsStr.TrimEnd(','), fieldsPara.TrimEnd(','));

            return result;
        }

        /// <summary>
        /// DateTime转为SQL
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string ToSQL(this DataBaseType type, DateTime dateTime)
        {
            switch (type)
            {
                case DataBaseType.ORACLE: return dateTime.ToSQLStringORACLE();
                case DataBaseType.MSACCESS: return dateTime.ToSQLStringMSACCESS();
                case DataBaseType.MSSQL: return dateTime.ToSQLStringMSSQL();
                case DataBaseType.SQLITE: return dateTime.ToSQLStringSQLITE();
                default: return string.Empty;
            }
        }
        #endregion
    }

    #region 数据驱动方式接入数据库

    public interface IBaseDAL
    {
        bool Exist();

        bool Insert();

        bool Update();

        bool UpdateAttribute(string attrName, object newVal);

        bool Delete();

        void SetIdentifiedFields(string fields);

        void SetTableName(string name);

        object QueryOne();

        int CurrentIndex();

        int QueryCount();

        void ReSet();

        object Next();

        List<object> Query();

        IBaseDAL CreateNew();

    }

    public class BaseDALClass
    {
        private int curRowIndex = 0;
        private int totalRowNum = 0;
        private int pageSize = 10;

        // 数据库连接
        private IDbConnection conn = null;
        // 数据库类型
        private SystemDataBase.DataBaseType dbType = SystemDataBase.DataBaseType.DEFAULT;
        // 表名称
        private string T_TableName = string.Empty;
        // 识别字段
        private string[] identifyFields = null;
        // 类属性和表字段的映射表
        private DataTable ctMapping = null;

        private bool checkAndCreateNonExistTable = false;

        private string orclTableSpace = string.Empty;

        public IDbConnection DbConnect { set { conn = value; } }
        public SystemDataBase.DataBaseType DbType { set { dbType = value; } }
        public string[] IdentifyFields { set { identifyFields = value; } }
        public string TableName { set { T_TableName = value; } }
        public DataTable CTMapping { set { ctMapping = value; } }
        public bool CheckAndCreateNonExistTable { set { checkAndCreateNonExistTable = value; } }
        public string OracleTableSpace { set { orclTableSpace = value; } }

        public BaseDALClass(DataDrivenInitArg arg)
        {
            DbConnect = arg.Connect;
            DbType = arg.DbType;
            TableName = arg.TableName;
            IdentifyFields = arg.IdentifyFields;
            CTMapping = arg.CtMapping;
            CheckAndCreateNonExistTable = arg.CheckAndCreateNonExistTable;
            orclTableSpace = arg.OracleTableSpace;
        }

        public static DataTable CreateCTMappingTableSchema()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ClassProperty", typeof(string));
            dt.Columns.Add("PropertyType", typeof(string));
            dt.Columns.Add("PropertyCategory", typeof(string));
            dt.Columns.Add("PropertyDisplayName", typeof(string));
            dt.Columns.Add("PropertyDescription", typeof(string));
            dt.Columns.Add("PropertyReadOnly", typeof(bool));
            dt.Columns.Add("PropertyVisibility", typeof(bool));

            dt.Columns.Add("TableField", typeof(string));
            dt.Columns.Add("FieldType", typeof(string));
            dt.Columns.Add("FieldDescription", typeof(string));

            return dt;
        }

        private string getAllTabelFields(string tableAlias)
        {
            string r = string.Empty;
            if (ctMapping != null)
            {
                foreach (DataRow dr in ctMapping.Rows)
                {
                    string fldType = dr["FieldType"].ToString();
                    string fldName = dr["TableField"].ToString();
                    if (fldType.Trim().ToUpper() == "SDO_GEOMETRY")
                        fldName = string.Format("{0}.{1}.get_wkt() as {1}", tableAlias, fldName);
                    r += fldName + ",";
                }
                r = r.TrimEnd(',');
            }
            return r;
        }

        private string getAllTabelFieldsFormat()
        {
            string r = string.Empty;
            if (ctMapping != null)
            {
                foreach (DataRow dr in ctMapping.Rows)
                {
                    string fldType = dr["FieldType"].ToString();
                    string fldName = dr["TableField"].ToString();
                    if (fldType.Trim().ToUpper() == "SDO_GEOMETRY")
                        fldName = string.Format("{0}{1}.get_wkt() as {1}", "{0}", fldName);
                    else
                        fldName = "{0}" + fldName;

                    r += fldName + ",";
                }
                r = r.TrimEnd(',');
            }
            return r;
        }


        /// <summary>
        /// 判断是否存在表，若没有则创建，返回true说明有表或者成功创建表
        /// </summary>
        /// <param name="tableName">表名称</param>
        /// <param name="orclTableSpace">oracle表空间名,只有oracle才填写</param>
        /// <returns></returns>
        private bool CheckThenCreateNonExitTable()
        {
            if (checkAndCreateNonExistTable)
            {
                if (!IsExistTable(T_TableName))
                    return CreateTable(T_TableName, orclTableSpace);
            }
            return false;
        }

        /// <summary>
        /// 创建数据库
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="orclTableSpace"></param>
        /// <returns></returns>
        public bool CreateTable(string tableName, string orclTableSpace = "")
        {
            bool result = false;
            if (!IsExistTable(tableName))
            {
                string createTableSql = string.Empty;
                string createTableFormat = string.Empty;
                string fieldFormat = string.Empty;

                switch (dbType)
                {
                    case SystemDataBase.DataBaseType.ORACLE:
                        createTableFormat = "CREATE TABLE {0} ({1}) TABLESPACE " + orclTableSpace;
                        fieldFormat = "{0} {1} {2},";
                        break;
                    case SystemDataBase.DataBaseType.SQLITE:
                        createTableFormat = "CREATE TABLE {0} ({1})";
                        fieldFormat = "{0} {1} {2},";
                        break;
                    case SystemDataBase.DataBaseType.MSSQL:
                    case SystemDataBase.DataBaseType.MSACCESS:
                        createTableFormat = "CREATE TABLE [{0}] ({1})";
                        fieldFormat = "[{0}] {1} {2},";
                        break;
                }

                string fields = string.Empty;
                if (ctMapping != null)
                    foreach (DataRow dr in ctMapping.Rows)
                    {
                        string field = dr["TableField"].ToString();
                        string type = dr["FieldType"].ToString();
                        string desc = dr["FieldDescription"].ToString();
                        fields += string.Format(fieldFormat, field, type, desc);
                    }
                fields = fields.Substring(0, fields.Length - 1);
                createTableSql = string.Format(createTableFormat, tableName, fields);

                string msg = string.Empty;
                result = conn.Execute(createTableSql, out msg);
            }
            else
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// 判断数据库表是否存在
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public bool IsExistTable(string tableName)
        {
            bool result = false;
            string existSql = dbType.ExistTableSql(tableName);
            if (conn == null) return result;

            return conn.HasRecord(existSql);
        }

        private Dictionary<string, string> GetCTMap()
        {
            Dictionary<string, string> map = new Dictionary<string, string>();
            if (ctMapping != null)
                foreach (DataRow dr in ctMapping.Rows)
                {
                    string classProperty = dr["ClassProperty"].ToString();
                    string tableField = dr["TableField"].ToString();
                    map.Add(classProperty, tableField);
                }
            return map;
        }

        public static DataTable AdjustFieldsOrder(object obj, DataTable fieldTable, Dictionary<string, string> map)
        {
            DataTable result = fieldTable.Clone();

            try
            {
                PropertyInfo[] props = obj.GetType().GetProperties();
                for (int i = 0; i < props.Length; i++)
                {
                    PropertyInfo prop = props[i];
                    string propName = prop.Name;
                    if (map.Keys.Contains(propName)) propName = map[propName];
                    DataRow[] drs = fieldTable.Select(string.Format("FIELDNAME = '{0}'", propName));
                    if (drs.Length > 0)
                    {
                        DataRow dr = result.NewRow();
                        dr.ItemArray = drs[0].ItemArray;
                        result.Rows.Add(dr);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return result;
        }

        public static DataTable AdjustFieldsOrder(object obj, DataTable fieldTable)
        {
            DataTable result = fieldTable.Clone();

            try
            {
                PropertyInfo[] props = obj.GetType().GetProperties();
                for (int i = 0; i < props.Length; i++)
                {
                    PropertyInfo prop = props[i];
                    string propName = prop.Name;
                    DataRow[] drs = fieldTable.Select(string.Format("FIELDNAME = '{0}'", propName));
                    if (drs.Length > 0)
                    {
                        DataRow dr = result.NewRow();
                        dr.ItemArray = drs[0].ItemArray;
                        result.Rows.Add(dr);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return result;
        }

        private string isExistWhereClause(object obj)
        {
            string r = "1=1";
            if (identifyFields == null || identifyFields.Length == 0
                || ctMapping == null || ctMapping.Rows.Count == 0
                || obj == null)
                r = "1=1";
            else
            {
                r = string.Empty;
                foreach (string idFld in identifyFields)
                {
                    DataRow[] drs = ctMapping.Select(string.Format("TableField='{0}'", idFld));
                    if (drs.Length > 0)
                    {
                        DataRow dr = drs[0];
                        string fld = dr["TableField"].ToString();
                        string fldType = dr["FieldType"].ToString().ToLower();
                        string propName = dr["ClassProperty"].ToString();

                        Type objType = obj.GetType();
                        PropertyInfo propertyInfo = objType.GetProperty(propName); //获取指定名称的属性

                        object val = propertyInfo.GetValue(obj, null); //获取属性值
                        string valStr = string.Empty;

                        string format = "{0}='{1}',";
                        if (!fldType.Contains("char")) format = "{0}={1},";

                        if (val == null || val == DBNull.Value)
                        {
                            valStr = "NULL";
                            format = "{0}={1},";
                        }
                        else
                        {
                            valStr = val.ToString();
                            if (val.GetType().Equals(typeof(DateTime))) valStr = dbType.ToSQL((DateTime)val);
                        }
                        r += string.Format(format, fld, valStr);
                    }
                    r = r.TrimEnd(',');
                }
            }
            return r;
        }

        public bool Exist(object obj)
        {
            string sql = dbType.GetSelectSQL(T_TableName, isExistWhereClause(obj));
            return conn.HasRecord(sql, true);
        }

        public bool ExistOnlyOne(object obj)
        {
            string sql = dbType.GetSelectSQL(T_TableName, isExistWhereClause(obj));
            return conn.HasRecordOnlyOne(sql, true);
        }

        public bool Insert(object obj)
        {
            bool result = false;

            CheckThenCreateNonExitTable();

            if (Exist(obj)) return result;

            Dictionary<string, string> map = GetCTMap();

            DataTable fldDt = conn.GetFieldsTable(T_TableName, dbType);

            // 调整fieldTable顺序
            DataTable newtable = AdjustFieldsOrder(obj, fldDt, map);

            IList<IDataParameter> collect = obj.ParseDbParameterCollection2(newtable, dbType, map);
            string sql = dbType.GetInsertSQL(newtable, T_TableName);

            conn.OpenDB();
            IDbTransaction trans = conn.BeginTransaction();

            try
            {
                IDbCommand cmd = conn.CreateDbCommand(sql);
                cmd.Parameters.AddCollect(collect);
                cmd.Transaction = trans;
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch (Exception ex)
            {
                result = false;
                trans.Rollback();

                Console.WriteLine(ex.Message);
                LogException(ex);
            }
            finally
            {
                conn.CloseDB();
            }
            return result;
        }

        public bool Update(object obj)
        {
            bool result = false;
            CheckThenCreateNonExitTable();

            if (!ExistOnlyOne(obj)) return result;

            Dictionary<string, string> map = GetCTMap();

            DataTable fldDt = conn.GetFieldsTable(T_TableName, dbType);

            // 调整fieldTable顺序
            DataTable newtable = AdjustFieldsOrder(obj, fldDt, map);

            IList<IDataParameter> collect = obj.ParseDbParameterCollection2(newtable, dbType, map);
            string sql = dbType.GetUpdateSQL(newtable, T_TableName, isExistWhereClause(obj));

            conn.OpenDB();
            IDbTransaction trans = conn.BeginTransaction();
            IDbCommand cmd = conn.CreateDbCommand(sql);
            cmd.Parameters.AddCollect(collect);
            cmd.Transaction = trans;

            try
            {
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                result = false;
                trans.Rollback();
                LogException(ex);
            }
            finally
            {
                conn.CloseDB();
            }

            return result;
        }

        public bool UpdateAttribute(object obj, string attrName, object newVal)
        {
            bool result = false;
            CheckThenCreateNonExitTable();

            if (!ExistOnlyOne(obj)) return result;

            Dictionary<string, string> map = GetCTMap();

            DataTable fldDt = conn.GetFieldsTable(T_TableName, dbType);

            foreach (DataRow dr in fldDt.Rows)
            {
                string fldName = dr["FIELDNAME"].ToString();
                string fldName2 = map[attrName].ToString();
                if (fldName != fldName2) fldDt.Rows.Remove(dr);
            }

            IList<IDataParameter> collect = obj.ParseDbParameterCollection2(fldDt, dbType, map);
            string sql = dbType.GetUpdateSQL(fldDt, T_TableName, isExistWhereClause(obj));

            conn.OpenDB();
            IDbTransaction trans = conn.BeginTransaction();
            IDbCommand cmd = conn.CreateDbCommand(sql);
            cmd.Parameters.AddCollect(collect);
            cmd.Transaction = trans;

            try
            {
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch (Exception ex)
            {
                result = false;
                trans.Rollback();
                LogException(ex);
            }
            finally
            {
                conn.CloseDB();
            }

            return result;
        }

        public bool Delete(object obj)
        {
            bool result = false;
            CheckThenCreateNonExitTable();

            if (!Exist(obj)) return result;

            string sql = dbType.GetDeleteSQL(T_TableName, isExistWhereClause(obj));

            conn.OpenDB();

            IDbTransaction trans = conn.BeginTransaction();
            IDbCommand cmd = conn.CreateDbCommand(sql);
            cmd.Transaction = trans;

            try
            {
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch (Exception ex)
            {
                trans.Rollback();
                LogException(ex);
            }
            finally
            {
                conn.CloseDB();
            }
            return result;
        }

        public void SetIdentifiedFields(string[] identifyFields)
        {
            IdentifyFields = identifyFields;
        }

        public void SetTableName(string name)
        {
            T_TableName = name;
        }

        public object QueryOne(IBaseDAL dal)
        {
            // 加载Oracle.DataAccess.dll，使oracle spatial可用
            if (File.Exists("Oracle.DataAccess.dll"))
            {
                Assembly.LoadFrom("Oracle.DataAccess.dll");
            }

            CheckThenCreateNonExitTable();

            object obj = dal.CreateNew();
            //string tableAlias = "s";

            string sql = dbType.GetSelectSQL(
                T_TableName,// + " " + tableAlias,
                isExistWhereClause((object)dal));//,getAllTabelFields(tableAlias)

            Dictionary<string, string> map = GetCTMap();

            conn.OpenDB();
            IDbCommand cmd = conn.CreateDbCommand(sql);

            IDbDataAdapter adpt = dbType.CreateDBDataAdapter();
            DataSet ds = adpt.GetDataSet(conn, sql);
            if (ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.Rows[0];
                    obj.GetValueFrom2(dr, map);
                }
                else
                {
                    obj = null;
                }
            }
            else
            {
                obj = null;
            }
            return obj;
        }

        public int CurrentIndex()
        {
            return curRowIndex;
        }

        public int QueryCount(IBaseDAL dal)
        {
            CheckThenCreateNonExitTable();
            object dalObj = ReComputeDalObj(dal);

            string sql = dbType.GetSelectSQL(
                T_TableName,
                isExistWhereClause(dalObj),
                "count(*)");
            conn.OpenDB();

            totalRowNum = -1;

            IDbTransaction trans = conn.BeginTransaction();
            try
            {
                IDbCommand cmd = conn.CreateDbCommand(sql);
                cmd.Transaction = trans;
                object obj = cmd.ExecuteScalar();
                totalRowNum = Convert.ToInt32(obj);
            }
            catch (Exception ex)
            {
                LogException(ex);
            }
            finally
            {
                conn.CloseDB();
            }
            return totalRowNum;
        }

        public void ReSet()
        {
            curRowIndex = 0;
        }

        public object Next(IBaseDAL dal)
        {
            //if (!IsExistTable(T_TableName)) return null;

            if (curRowIndex >= totalRowNum || curRowIndex < 0) return null;

            object dalObj = ReComputeDalObj(dal);

            string isExistClause = isExistWhereClause(dalObj);

            string sql = dbType.GetNextItemSQL(T_TableName, curRowIndex + 1, isExistClause, getAllTabelFieldsFormat());

            object obj = dal.CreateNew();

            Dictionary<string, string> map = GetCTMap();

            conn.OpenDB();
            IDbCommand cmd = conn.CreateDbCommand(sql);

            IDbDataAdapter adpt = dbType.CreateDBDataAdapter();
            DataSet ds = adpt.GetDataSet(conn, sql);
            if (ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.Rows[0];
                    obj.GetValueFrom2(dr, map);
                    curRowIndex++;
                }
                else
                {
                    obj = null;
                }
            }
            else
            {
                obj = null;
            }
            return obj;
        }

        public List<object> Query(IBaseDAL dal)
        {
            CheckThenCreateNonExitTable();

            List<object> objs = new List<object>();

            object dalObj = ReComputeDalObj(dal);

            string sql = dbType.GetSelectSQL(T_TableName, isExistWhereClause(dalObj));

            Dictionary<string, string> map = GetCTMap();

            conn.OpenDB();
            IDbCommand cmd = conn.CreateDbCommand(sql);

            IDbDataAdapter adpt = dbType.CreateDBDataAdapter();
            DataSet ds = adpt.GetDataSet(conn, sql);
            if (ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                foreach (DataRow dr in dt.Rows)
                {
                    object obj = dal.CreateNew();
                    obj.GetValueFrom2(dr, map);
                    objs.Add(obj);
                }
            }
            return objs;
        }

        private object ReComputeDalObj(IBaseDAL dal)
        {
            object dalObj = (object)dal;
            if (identifyFields != null || identifyFields.Length > 0)
                foreach (string ifld in identifyFields)
                {
                    if (ctMapping != null)
                    {
                        DataRow[] drs = ctMapping.Select(string.Format("TableField='{0}'", ifld));
                        if (drs.Length > 0)
                        {
                            DataRow dr = drs[0];
                            string clsProp = dr["ClassProperty"].ToString();
                            object val = dalObj.GetPropertyValue(clsProp);
                            if (val == null) { dalObj = null; break; }
                        }
                    }
                }
            else
                dalObj = null;

            return dalObj;
        }

        private void LogException(Exception e)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
            if (!System.IO.Directory.Exists(path)) System.IO.Directory.CreateDirectory(path);
            DateTime dt = DateTime.Now;
            string file = Path.Combine(path,
               string.Format(
            "数据驱动日志{0}{1}{2}{3}{4}{5}{6}.txt",
            dt.Year.ToString().PadLeft(4, '0'),
            dt.Month.ToString().PadLeft(2, '0'),
            dt.Day.ToString().PadLeft(2, '0'),
            dt.Hour.ToString().PadLeft(2, '0'),
            dt.Minute.ToString().PadLeft(2, '0'),
            dt.Second.ToString().PadLeft(2, '0'),
            dt.Millisecond
            ));

            File.WriteAllText(file, "数据驱动日志：" + e.Message + e.StackTrace);
        }

    }

    public class DataDrivenInitArg
    {
        public IDbConnection Connect { get; set; }
        public SystemDataBase.DataBaseType DbType { get; set; }
        public string NameSpace { set; get; }
        public string ClassName { set; get; }
        public string ClassDisplayName { get; set; }
        public string TableName { set; get; }
        public string[] IdentifyFields { set; get; }
        public DataTable CtMapping { set; get; }
        public bool CheckAndCreateNonExistTable { set; get; }
        public string OracleTableSpace { set; get; }
    }

    /// <summary>
    /// 应用于属性和表字段名一致的情况
    /// </summary>
    [Serializable]
    public class SimpleBaseDALClass
    {
        private IDbConnection conn = null;
        private SystemDataBase.DataBaseType dbType = SystemDataBase.DataBaseType.DEFAULT;
        private string T_TableName = "";
        private string identifyWhereClause = "";

        public string IdentifyWhereClause { set { identifyWhereClause = value; } }

        public SimpleBaseDALClass(IDbConnection connect, SystemDataBase.DataBaseType type, string tableName, string whereClauseIdentify)
        {
            conn = connect;
            dbType = type;
            T_TableName = tableName;
            identifyWhereClause = whereClauseIdentify;
        }

        public bool ExistOnlyOne(object obj)
        {
            string sql = dbType.GetSelectSQL(T_TableName, identifyWhereClause);
            return conn.HasRecordOnlyOne(sql, true);
        }

        public bool Exist(object obj)
        {
            string sql = dbType.GetSelectSQL(T_TableName, identifyWhereClause);
            return conn.HasRecord(sql, true);
        }

        public bool Insert(object obj)
        {
            bool result = false;

            if (Exist(obj)) return result;

            DataTable fldDt = conn.GetFieldsTable(T_TableName, dbType);

            // 调整fieldTable顺序
            DataTable newtable = BaseDALClass.AdjustFieldsOrder(obj, fldDt);

            IList<IDataParameter> collect = obj.ParseDbParameterCollection(newtable, dbType);
            string sql = dbType.GetInsertSQL(newtable, T_TableName);

            conn.OpenDB();
            IDbTransaction trans = conn.BeginTransaction();

            try
            {
                IDbCommand cmd = conn.CreateDbCommand(sql);
                cmd.Parameters.AddCollect(collect);
                cmd.Transaction = trans;
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                result = false;
                trans.Rollback();
            }
            finally
            {
                conn.CloseDB();
            }
            return result;
        }

        public bool Update(object obj)
        {
            bool result = false;
            if (!ExistOnlyOne(obj)) return result;

            DataTable fldDt = conn.GetFieldsTable(T_TableName, dbType);

            // 调整fieldTable顺序
            DataTable newtable = BaseDALClass.AdjustFieldsOrder(obj, fldDt);

            IList<IDataParameter> collect = obj.ParseDbParameterCollection(newtable, dbType);
            string sql = dbType.GetUpdateSQL(newtable, T_TableName, identifyWhereClause);

            conn.OpenDB();
            IDbTransaction trans = conn.BeginTransaction();
            IDbCommand cmd = conn.CreateDbCommand(sql);
            cmd.Parameters.AddCollect(collect);
            cmd.Transaction = trans;

            try
            {
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                result = false;
                trans.Rollback();
            }
            finally
            {
                conn.CloseDB();
            }

            return result;
        }

        public bool Delete(object obj)
        {
            bool result = false;
            if (!Exist(obj)) return result;

            string sql = dbType.GetDeleteSQL(T_TableName, identifyWhereClause);

            conn.OpenDB();

            IDbTransaction trans = conn.BeginTransaction();
            IDbCommand cmd = conn.CreateDbCommand(sql);
            cmd.Transaction = trans;

            try
            {
                result = cmd.ExecuteNonQuery() > 0 ? true : false;
                trans.Commit();
            }
            catch
            {
                trans.Rollback();
            }
            finally
            {
                conn.CloseDB();
            }
            return result;
        }

    }

    /// <summary>
    /// 季度枚举
    /// </summary>
    public enum EnumQuarter : int
    {
        [Description("第一季度")]
        Quarter1th = 1,

        [Description("第二季度")]
        Quarter2nd = 2,

        [Description("第三季度")]
        Quarter3rd = 3,

        [Description("第四季度")]
        Quarter4th = 4
    }
    #endregion
}
#endregion

#region System.Drawing
namespace System.Drawing
{
    public static class ExtendSystemDrawing
    {
        /// <summary>
        /// 将图像转换为其用Base64数字编码的等效字符串表示形式。
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static string ToBase64String(this Image image)
        {
            if (image == null) throw new NullReferenceException();

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            image.Save(ms, image.RawFormat);
            byte[] buffer = ms.ToArray();
            ms.Close();

            return Convert.ToBase64String(buffer);
        }

        /// <summary>
        /// Base64数字编码的字符串转换为图像。
        /// </summary>
        /// <param name="base64String"></param>
        /// <returns></returns>
        public static Image ToImage(this string base64String)
        {
            if (string.IsNullOrWhiteSpace(base64String)) return null;
            byte[] buffer = Convert.FromBase64String(base64String);

            IO.MemoryStream stream = new IO.MemoryStream();
            stream.Write(buffer, 0, buffer.Length);
            Image img = Image.FromStream(stream);
            stream.Close();

            return img;
        }

        /// <summary>
        /// Icon转换为Image
        /// </summary>
        /// <param name="icon"></param>
        /// <returns></returns>
        public static Image ToImage(this Icon icon)
        {
            return Image.FromHbitmap(icon.ToBitmap().GetHbitmap());
        }

        /// <summary>
        /// Image转换为Icon
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static Icon ToIcon(this Image image)
        {
            return Icon.FromHandle(((Bitmap)image).GetHicon());
        }

        public static byte[] ToByte2(this Image image)
        {
            MemoryStream ms = new MemoryStream();
            image.Save(ms, ImageFormat.Bmp);
            byte[] bs = ms.ToArray();
            ms.Close();
            return bs;
        }

        /// <summary>
        /// 转换为字节数组
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this Image image)
        {
            try
            {
                if (image == null) return null;
                using (Bitmap bitmap = new Bitmap(image))
                {
                    using (MemoryStream stream = new MemoryStream())
                    {
                        bitmap.Save(stream, ImageFormat.Jpeg);
                        return stream.GetBuffer();
                    }
                }
            }
            finally
            {
                if (image != null)
                {
                    image.Dispose();
                    image = null;
                }
            }
        }

        /// <summary>
        /// 转换为Image类型
        /// </summary>
        /// <param name="bs"></param>
        /// <returns></returns>
        public static Image ToImage(this byte[] bs)
        {
            if (bs == null) return null;
            if (bs.Length == 0) return null;
            Bitmap bmp = null;
            using (MemoryStream ms = new MemoryStream(bs))
            {
                bmp = new Bitmap(ms);
                ms.Close();
            }
            return bmp;
        }

        /// <summary>
        /// 转为stdole.IFontDisp
        /// </summary>
        /// <param name="fnt"></param>
        /// <returns></returns>
        public static stdole.IFontDisp ToIFontDisp(this Font fnt)
        {
            return System.Windows.Forms.FontHelper.GetIFontDispFrom(fnt);
        }

    }

    /// <summary>
    /// 系统图标获取类
    /// 创建人：谢成磊
    /// 创建时间：2015.08.16
    /// </summary>
    public static class IconHelper
    {
        /// <summary> 
        /// 通过扩展名得到图标和描述 
        /// </summary> 
        /// <param name="ext">扩展名</param> 
        /// <param name="LargeIcon">得到大图标</param> 
        /// <param name="smallIcon">得到小图标</param> 
        public static void GetExtsIconAndDescription(string Extension, out Icon largeIcon, out Icon smallIcon, out string description)
        {
            largeIcon = smallIcon = null;
            description = "";
            var extsubkey = Registry.ClassesRoot.OpenSubKey(Extension); //从注册表中读取扩展名相应的子键 
            if (extsubkey != null)
            {
                var extdefaultvalue = (string)extsubkey.GetValue(null); //取出扩展名对应的文件类型名称 

                //未取到值，返回预设图标 
                if (extdefaultvalue == null)
                {
                    GetDefaultIcon(out largeIcon, out smallIcon);
                    return;
                }

                var typesubkey = Registry.ClassesRoot.OpenSubKey(extdefaultvalue); //从注册表中读取文件类型名称的相应子键 
                if (typesubkey != null)
                {
                    description = (string)typesubkey.GetValue(null); //得到类型描述字符串 
                    var defaulticonsubkey = typesubkey.OpenSubKey("DefaultIcon"); //取默认图标子键 
                    if (defaulticonsubkey != null)
                    {
                        //得到图标来源字符串 
                        var defaulticon = (string)defaulticonsubkey.GetValue(null); //取出默认图标来源字符串 
                        var iconstringArray = defaulticon.Split(',');
                        int nIconIndex = 0;
                        if (iconstringArray.Length > 1) int.TryParse(iconstringArray[1], out nIconIndex);
                        //得到图标 
                        System.IntPtr phiconLarge = new System.IntPtr();
                        System.IntPtr phiconSmall = new System.IntPtr();
                        ExtractIconExW(iconstringArray[0].Trim('"'), nIconIndex, ref phiconLarge, ref phiconSmall, 1);
                        if (phiconLarge.ToInt32() > 0) largeIcon = Icon.FromHandle(phiconLarge);
                        if (phiconSmall.ToInt32() > 0) smallIcon = Icon.FromHandle(phiconSmall);
                    }
                }
            }
        }

        /// <summary> 
        /// 获取缺省图标 
        /// </summary> 
        /// <param name="largeIcon"></param> 
        /// <param name="smallIcon"></param> 
        public static void GetDefaultIcon(out Icon largeIcon, out Icon smallIcon)
        {
            largeIcon = smallIcon = null;
            System.IntPtr phiconLarge = new System.IntPtr();
            System.IntPtr phiconSmall = new System.IntPtr();
            ExtractIconExW(Path.Combine(System.Environment.SystemDirectory, "shell32.dll"), 0, ref phiconLarge, ref phiconSmall, 1);
            if (phiconLarge.ToInt32() > 0) largeIcon = Icon.FromHandle(phiconLarge);
            if (phiconSmall.ToInt32() > 0) smallIcon = Icon.FromHandle(phiconSmall);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lpszFile">WCHAR*</param>
        /// <param name="nIconIndex"></param>
        /// <param name="phiconLarge">HICON*</param>
        /// <param name="phiconSmall">HICON*</param>
        /// <param name="nIcons">unsigned int</param>
        /// <returns>unsigned int</returns>
        [DllImportAttribute("shell32.dll", EntryPoint = "ExtractIconExW", CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall)]
        public static extern uint ExtractIconExW([System.Runtime.InteropServices.InAttribute()] [System.Runtime.InteropServices.MarshalAsAttribute(UnmanagedType.LPWStr)] string lpszFile, int nIconIndex, ref System.IntPtr phiconLarge, ref System.IntPtr phiconSmall, uint nIcons);

        /// <summary>
        /// 根据扩展名创建图标
        /// </summary>
        /// <param name="fileType"></param>
        /// <param name="icon"></param>
        /// <param name="isLarge">是否大图标</param>
        /// <returns>类型描述字符串</returns>
        public static string GetFileIcon(string Extension, out Icon icon, bool isLarge)
        {
            string des = string.Empty;

            Icon small;
            Icon large;

            if (Extension.Trim() == "" || Extension == null) //预设图标
            {
                GetDefaultIcon(out large, out small);
            }
            else if (Extension.ToUpper() == ".EXE") //应用程序图标单独获取
            {
                System.IntPtr l = System.IntPtr.Zero;
                System.IntPtr s = System.IntPtr.Zero;

                ExtractIconExW(Path.Combine(System.Environment.SystemDirectory, "shell32.dll"), 2, ref l, ref s, 1);

                large = Icon.FromHandle(l);
                small = Icon.FromHandle(s);
            }
            else //其它类型的图标
            {
                GetExtsIconAndDescription(Extension, out large, out small, out des);
            }

            if ((large == null) || (small == null)) //无法获取图标,预设图标
                GetDefaultIcon(out large, out small);

            if (isLarge)
                icon = large;
            else
                icon = small;

            return des;
        }

        [StructLayoutAttribute(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct HICON__
        {
            public int unused;
        }
    }
}
#endregion

#region System.Drawing.Imaging
namespace System.Drawing.Imaging
{
    public static class ImageFormatExtend
    {
        public static string ToFileExtension(this ImageFormat imgFormat)
        {
            string ext = "";
            if (imgFormat.Equals(ImageFormat.Jpeg)) ext = ".jpg";
            else if (imgFormat.Equals(ImageFormat.Bmp)) ext = ".bmp";
            else if (imgFormat.Equals(ImageFormat.Emf)) ext = ".emf";
            else if (imgFormat.Equals(ImageFormat.Exif)) ext = ".exif";
            else if (imgFormat.Equals(ImageFormat.Gif)) ext = ".gif";
            else if (imgFormat.Equals(ImageFormat.Icon)) ext = ".ico";
            else if (imgFormat.Equals(ImageFormat.MemoryBmp)) ext = ".bmp";
            else if (imgFormat.Equals(ImageFormat.Png)) ext = ".png";
            else if (imgFormat.Equals(ImageFormat.Tiff)) ext = ".tif";
            else if (imgFormat.Equals(ImageFormat.Wmf)) ext = ".wmf";

            return ext;
        }
    }
}
#endregion

#region System.GeographyCalculate
namespace System.GeographyCalculate
{
    /// <summary>
    /// 大椭圆计算类
    /// 创建人：谢成磊
    /// 创建时间：2015.01.04
    /// </summary>
    public class GreatEllipse
    {
        private static double
            a = 6378137.0,//cgcs2000
            b = 6356752.3141403561,
            e,
            r,
            l,
            m;
        private static long
            rate = 1000000000;

        /// <summary>
        /// 计算纬线长
        /// </summary>
        /// <param name="latHigh"></param>
        /// <param name="latLow"></param>
        /// <returns></returns>
        public static double LengthSN(double latHigh, double latLow)
        {
            double x1 = ElArcLen(latHigh * Math.PI / 180, a, b);
            double x2 = ElArcLen(latLow * Math.PI / 180, a, b);
            return (x2 - x1);
        }

        /// <summary>
        /// 计算经线长
        /// </summary>
        /// <param name="detLon">经度差</param>
        /// <param name="lat">纬度</param>
        /// <returns></returns>
        public static double LengthEW(double detLon, double lat)
        {
            detLon = detLon * Math.PI / 180;
            lat = lat * Math.PI / 180;
            double 离心角 = Math.Atan(b * Math.Tan(lat) / a);
            double 极角 = Math.Atan(b * b * Math.Tan(lat) / a / a);
            double x = a * Math.Cos(离心角);
            double y = b * Math.Sin(离心角);
            double r = Math.Sqrt(x * x + y * y);
            r = r * Math.Cos(极角);
            return (r * detLon);
        }

        //a、b分别是椭圆的长半轴和短半轴///////||使用说明：ElArcLen()计算的是
        //e2椭圆的扁率的平方,n的值从1到9///////||         从极点到纬度detA的长度
        //                                     ||         而不是从赤道

        /// <summary>
        /// 从极点到纬度detA的长度
        /// </summary>
        /// <param name="detA"></param>
        /// <param name="a">长半轴</param>
        /// <param name="b">短半轴</param>
        /// <returns></returns>
        public static double ElArcLen(double detA, double a, double b)//detA是纬度|弧度制
        {
            int n = 9;
            double e2 = 1 - b * b / (a * a);
            //角度转换开始//ee为离心角的余角
            double ee = Math.PI / 2 - Math.Atan(b * Math.Tan(detA) / a);
            //角度转换结束
            double result = ee;

            for (int i = 1; i <= n; i++)
            {
                result -= power(0.5 * e2, i) * factorial2(i - 1) * sin2N(ee, 2 * i) / factorial(i);
            }
            return (result * a);
        }
        //iAngle是弧度制的角度；_2N是sin函数的偶数幂指数
        //_2N值不能大于18
        private static double sin2N(double iAngle, int _2N)
        {
            if (_2N <= 18 && _2N % 2 == 0 && _2N != 0)
            {
                double sum = 0;
                double xs = 1;
                int _n = _2N / 2;
                for (int i = 0; i <= _n; i++)
                {
                    sum += f(iAngle, i) * power(-1, i) * factorial(_n) / (factorial(i) * factorial(_n - i));
                }
                xs = power(0.5, _n);
                return (sum * xs);
            }
            else
            {
                return -1;
            }
        }
        //计算2*iAngle余弦的i次幂从0到小于PI/2区间的定积分
        //目前i值支持到9
        private static double f(double iAngle, int i)
        {
            double iA = 2 * iAngle;
            double result = 0;
            switch (i)
            {
                case 0:
                    result = iAngle;
                    break;
                case 1:
                    result = 0.5 * Math.Sin(iA);
                    break;
                case 2:
                    result = 0.5 * iAngle + Math.Sin(2 * iA) / 8;
                    break;
                case 3:
                    result = 0.5 * Math.Sin(iA) - power(Math.Sin(iA), 3) / 6;
                    break;
                case 4:
                    result = (3 * iAngle + Math.Sin(2 * iA)) / 8 + Math.Sin(4 * iA) / 64;
                    break;
                case 5:
                    result = 0.5 * Math.Sin(iA) - power(Math.Sin(iA), 3) / 3 + power(Math.Sin(iA), 5) * 0.1;
                    break;
                case 6:
                    result = (5 * iAngle + 2 * Math.Sin(2 * iA) + Math.Sin(4 * iA)) / 16 - power(Math.Sin(2 * iA), 3) / 96;
                    break;
                case 7:
                    result = Math.Sin(iA) / 2 - power(Math.Sin(iA), 3) / 2 + 0.3 * power(Math.Sin(iA), 5) - power(Math.Sin(iA), 7) / 14;
                    break;
                case 8:
                    result = 35 * iAngle / 128 + Math.Sin(2 * iA) / 8 + 7 * Math.Sin(4 * iA) / 256 + Math.Sin(8 * iA) / 2048 - power(Math.Sin(2 * iA), 3) / 48;
                    break;
                case 9:
                    result = Math.Sin(iA) / 2 - 2 * power(Math.Sin(iA), 3) / 3 + 3 * power(Math.Sin(iA), 5) / 5 - 2 * power(Math.Sin(iA), 7) / 7 + power(Math.Sin(iA), 9) / 18;
                    break;
                default:
                    result = -1;
                    break;
            }
            return result;
        }
        private static double factorial(int n)//计算n的阶乘，步长为1
        {
            if (n > 0)
            {
                double s = 1;
                for (int i = 1; i <= n; i++)
                {
                    s *= i;
                }
                return s;
            }
            else if (n == 0)
            { return 1; }
            else
            { return -1; }
        }
        private static double factorial2(int n)//阶乘，步长为2
        {
            if ((2 * n - 1) > 0)
            {
                double s = 1;
                for (int i = 1; i <= 2 * n - 1; i += 2)
                {
                    s *= i;
                }
                return s;
            }
            else if (2 * n - 1 == -1 || 2 * n - 1 == 0)
            { return 1; }
            else
            { return -1; }
        }
        private static double power(double n, int i)//n的i次幂
        {
            double result = 1;
            if (i > 0)
            {
                for (int j = 0; j < i; j++)
                {
                    result *= n;
                }
            }
            return result;
        }

        /// <summary>
        /// 将大地坐标转换为空间直角坐标系(以地心与北极的连线为Z轴,
        /// 地心与零度经线与赤道的交点连线为X轴,地心与90度经线与赤道的交点连线为Y轴)。//北半球专用，东经专用
        /// </summary>
        /// <param name="srcPt">坐标，结构为(经度,纬度[,大地高]),大地高可以忽略</param>
        /// <returns></returns>
        public static double[] BLH_XYZ(double[] srcPt)
        {
            int d_num = srcPt.Length;
            if (d_num < 2 || d_num > 3) throw new FormatException("double[]对象结构不合理，结构：(经度,纬度[,大地高])");

            double 经度 = srcPt[0], 纬度 = srcPt[1], 大地高 = 0;
            if (d_num == 3) 大地高 = srcPt[2];

            经度 = 经度 * Math.PI / 180;
            纬度 = 纬度 * Math.PI / 180;
            double 离心角 = Math.Atan(b * Math.Tan(纬度) / a);
            double 极角 = Math.Atan(b * b * Math.Tan(纬度) / a / a);
            double x = a * Math.Cos(离心角);
            double y = b * Math.Sin(离心角);
            double r = Math.Sqrt(x * x + y * y);
            double R = Math.Sqrt(r * r + 大地高 * 大地高 - 2 * r * 大地高 * Math.Cos(Math.PI - 纬度 + 极角));
            double 极角_ = Math.Asin(Math.Sin(Math.PI - 纬度 + 极角) * 大地高 / R);
            极角 += 极角_;
            double[] xyz = new double[3];
            xyz[0] = R * Math.Cos(极角) * Math.Cos(经度);
            xyz[1] = R * Math.Cos(极角) * Math.Sin(经度);
            xyz[2] = R * Math.Sin(极角);
            return xyz;
        }
        public static double[] BLH2XYZ(double[] srcPt)
        {
            int d_num = srcPt.Length;
            if (d_num != 3) throw new FormatException("double[]对象结构不合理，结构：(经度,纬度,大地高)");

            double k = Math.PI / 180;
            double L = srcPt[0] * k, B = srcPt[1] * k, H = srcPt[2];

            double W = Math.Sqrt(1 - Math.Pow(e, 2) * Math.Pow(Math.Sin(B), 2));
            double N = a / W;

            double t1 = Math.Cos(B);
            double t2 = Math.Sin(B);
            double t3 = Math.Cos(L);
            double t4 = Math.Sin(L);

            double X = (N + H) * Math.Cos(B) * Math.Cos(L);
            double Y = (N + H) * Math.Cos(B) * Math.Sin(L);
            double Z = (N * (1 - Math.Pow(e, 2)) + H) * Math.Sin(B);//Math.Truncate(
            double[] xyz = new double[3] { X, Y, Z };
            return xyz;
        }
        public static double[] XYZ2BLH(double[] srcPt)
        {
            int d_num = srcPt.Length;
            if (d_num != 3) throw new FormatException("double[]对象结构不合理,应为(x,y,z)");

            double X = srcPt[0], Y = srcPt[1], Z = srcPt[2];
            double R = Math.Sqrt(Math.Pow(X, 2) + Math.Pow(Y, 2) + Math.Pow(Z, 2));
            double φ = Math.Atan(Z / Math.Sqrt(Math.Pow(X, 2) + Math.Pow(Y, 2)));
            double W = double.NaN;
            double tolerance = 0.00000000000000001;
            double k = 180 / Math.PI;
            double L = Math.Atan(Y / X);

            if (X < 0 && Y > 0) { L = Math.PI + L; }
            if (X < 0 && Y < 0) { L = Math.PI - L; }

            double B = getB(-Math.PI, Math.PI, ref W, tolerance, φ, Z);
            double N = a / W;
            double H = R * Math.Cos(φ) / Math.Cos(B) - N;
            B *= k;
            L *= k;
            double[] result = new double[3] { L, B, H };
            return result;
        }

        private static double getB(double MinB, double MaxB, ref double W, double tolerance, double φ, double Z)
        {
            double MidB = (MaxB + MinB) / 2;
            W = Math.Sqrt(1 - Math.Pow(e, 2) * Math.Pow(Math.Sin(MidB), 2));
            double B = Math.Atan(Math.Tan(φ) * (a * Math.Pow(e, 2) * Math.Sin(MidB) / Z / W + 1));

            double disMax = Math.Abs(MaxB - B);
            double disMin = Math.Abs(MinB - B);
            double tmpMaxB = double.NaN, tmpMinB = double.NaN;
            if (disMax > disMin)
            {
                tmpMinB = MinB;
                tmpMaxB = MidB;
            }
            else
            {
                tmpMinB = MidB;
                tmpMaxB = MaxB;
            }

            double compare = Math.Abs(MidB - B);
            if (compare <= tolerance) return B;

            return getB(tmpMinB, tmpMaxB, ref W, tolerance, φ, Z);
        }

        public static void SetSemiAxis(double A, double B)
        {
            if (a < b) throw new InvalidDataException("长半轴小于短半轴");
            a = A;
            b = B;
            e = Math.Sqrt((Math.Pow(a, 2) - Math.Pow(b, 2)) / Math.Pow(a, 2));
        }

        /// <summary>
        /// 距离计算（地理坐标，ps:积分计算，耗时多，精度高）
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public static double ComputeDistanceBearing(GeoPoint start, GeoPoint end)
        {
            GeoPoint p1, p2;
            if (start.Longitude < end.Longitude)
            {
                p1 = start;
                p2 = end;
            }
            else
            {
                p2 = start;
                p1 = end;
            }

            return Length(
                Deg2Rad(p1.Longitude),
                Deg2Rad(p1.Latitude),
                Deg2Rad(p2.Longitude),
                Deg2Rad(p2.Latitude)
                );
        }

        private static double Deg2Rad(double deg)
        {
            return (deg * Math.PI / 180);
        }
        private static double Length(double LonA, double LatA, double LonB, double LatB)
        {
            e = Math.Sqrt(Math.Pow(a, 2) - Math.Pow(b, 2)) / a;
            r = 1 - e * e;

            double molecular_l = Math.Cos(LatA) * Math.Sin(LonA) * r * Math.Sin(LatB) - Math.Cos(LatB) * Math.Sin(LonB) * r * Math.Sin(LatA);
            double molecular_m = Math.Cos(LatB) * Math.Cos(LonB) * r * Math.Sin(LatA) - Math.Cos(LatA) * Math.Cos(LonA) * r * Math.Sin(LatB);
            double Denominator = Math.Cos(LatA) * Math.Cos(LonA) * Math.Cos(LatB) * Math.Sin(LonB) - Math.Cos(LatA) * Math.Sin(LonA) * Math.Cos(LatB) * Math.Cos(LonB);

            l = molecular_l / Denominator;
            m = molecular_m / Denominator;

            double step = 1.0 / rate;
            long start = (long)(LonA * rate);
            long end = (long)(LonB * rate);

            double result = 0;
            for (long i = start; i < end; i++)
            {
                result += Integral(i * 1.0 / rate) * step;
            }
            return result;
        }
        private static double Integral(double longitude)
        {
            double t_ = Tan_Lat_(longitude);
            double t_r = Tan_Lat_r(longitude);

            double tmp = l * Math.Sin(longitude) - m * Math.Cos(longitude);
            tmp = Math.Pow(tmp, 2);

            return (a * Math.Sqrt((t_ * tmp / (t_r * t_r) + 1) / t_r));
        }
        private static double Tan_Lat_r(double longitude)
        {
            double temp = Tan_Lat(longitude);
            return (1 + r * temp * temp);
        }
        private static double Tan_Lat_(double longitude)
        {
            double temp = Tan_Lat(longitude);
            return (1 + temp * temp);
        }
        private static double Tan_Lat(double longitude)
        {
            return (-(l * Math.Cos(longitude) + m * Math.Sin(longitude)) / r);
        }

        /// <summary>
        /// 距离计算（地理坐标）
        /// </summary>
        /// <param name="p1"></param>
        /// <param name="p2"></param>
        /// <param name="course1"></param>
        /// <param name="course2"></param>
        /// <returns></returns>
        public static double ComputeDistanceBearing(GeoPoint p1, GeoPoint p2, out double course1, out double course2)
        {
            course1 = -1;
            course2 = course1;

            if (p1.x() == p2.x() && p1.y() == p2.y())
                return 0;

            double M_PI = Math.PI;
            double f = (a - b) / a;

            double p1_lat = Deg2Rad(p1.y()), p1_lon = Deg2Rad(p1.x());
            double p2_lat = Deg2Rad(p2.y()), p2_lon = Deg2Rad(p2.x());

            double L = p2_lon - p1_lon;
            double U1 = Math.Atan((1 - f) * Math.Tan(p1_lat));
            double U2 = Math.Atan((1 - f) * Math.Tan(p2_lat));
            double sinU1 = Math.Sin(U1), cosU1 = Math.Cos(U1);
            double sinU2 = Math.Sin(U2), cosU2 = Math.Cos(U2);
            double lambda = L;
            double lambdaP = 2 * M_PI;

            double sinLambda = 0;
            double cosLambda = 0;
            double sinSigma = 0;
            double cosSigma = 0;
            double sigma = 0;
            double alpha = 0;
            double cosSqAlpha = 0;
            double cos2SigmaM = 0;
            double C = 0;
            double tu1 = 0;
            double tu2 = 0;

            int iterLimit = 20;
            while (Math.Abs(lambda - lambdaP) > 1e-12 && --iterLimit > 0)
            {
                sinLambda = Math.Sin(lambda);
                cosLambda = Math.Cos(lambda);
                tu1 = (cosU2 * sinLambda);
                tu2 = (cosU1 * sinU2 - sinU1 * cosU2 * cosLambda);
                sinSigma = Math.Sqrt(tu1 * tu1 + tu2 * tu2);
                cosSigma = sinU1 * sinU2 + cosU1 * cosU2 * cosLambda;
                sigma = Math.Atan2(sinSigma, cosSigma);
                alpha = Math.Asin(cosU1 * cosU2 * sinLambda / sinSigma);
                cosSqAlpha = Math.Cos(alpha) * Math.Cos(alpha);
                cos2SigmaM = cosSigma - 2 * sinU1 * sinU2 / cosSqAlpha;
                C = f / 16 * cosSqAlpha * (4 + f * (4 - 3 * cosSqAlpha));
                lambdaP = lambda;
                lambda = L + (1 - C) * f * Math.Sin(alpha) *
                         (sigma + C * sinSigma * (cos2SigmaM + C * cosSigma * (-1 + 2 * cos2SigmaM * cos2SigmaM)));
            }

            if (iterLimit == 0)
                return -1;  // formula failed to converge

            double uSq = cosSqAlpha * (a * a - b * b) / (b * b);
            double A = 1 + uSq / 16384 * (4096 + uSq * (-768 + uSq * (320 - 175 * uSq)));
            double B = uSq / 1024 * (256 + uSq * (-128 + uSq * (74 - 47 * uSq)));
            double deltaSigma = B * sinSigma * (cos2SigmaM + B / 4 * (cosSigma * (-1 + 2 * cos2SigmaM * cos2SigmaM) -
                                                 B / 6 * cos2SigmaM * (-3 + 4 * sinSigma * sinSigma) * (-3 + 4 * cos2SigmaM * cos2SigmaM)));
            double s = b * A * (sigma - deltaSigma);
            course1 = Math.Atan2(tu1, tu2);
            course2 = Math.Atan2(cosU1 * sinLambda, -sinU1 * cosU2 + cosU1 * sinU2 * cosLambda) + M_PI;
            return s;
        }

        /// <summary>
        /// 面积计算（地理坐标） 没有测试
        /// </summary>
        /// <param name="points"></param>
        /// <returns></returns>
        public static double ComputePolygonArea(IList<GeoPoint> points)
        {
            computeAreaInit();

            double x1, y1, x2, y2, dx, dy;
            double Qbar1, Qbar2;
            double area, M_PI = Math.PI;

            int n = points.Count;
            x2 = Deg2Rad(points[n - 1].x());
            y2 = Deg2Rad(points[n - 1].y());
            Qbar2 = getQbar(y2);

            area = 0.0;

            for (int i = 0; i < n; i++)
            {
                x1 = x2;
                y1 = y2;
                Qbar1 = Qbar2;

                x2 = Deg2Rad(points[i].x());
                y2 = Deg2Rad(points[i].y());
                Qbar2 = getQbar(y2);

                if (x1 > x2)
                    while (x1 - x2 > M_PI)
                        x2 += m_TwoPI;
                else if (x2 > x1)
                    while (x2 - x1 > M_PI)
                        x1 += m_TwoPI;

                dx = x2 - x1;
                area += dx * (m_Qp - getQ(y2));

                if ((dy = y2 - y1) != 0.0)
                    area += dx * getQ(y2) - (dx / dy) * (Qbar2 - Qbar1);
            }
            if ((area *= m_AE) < 0.0)
                area = -area;

            /* kludge - if polygon circles the south pole the area will be
            * computed as if it cirlced the north pole. The correction is
            * the difference between total surface area of the earth and
            * the "north pole" area.
            */
            if (area > m_E)
                area = m_E;
            if (area > m_E / 2)
                area = m_E - area;

            return area;
        }
        static void computeAreaInit()
        {
            double M_PI = Math.PI;
            double a2 = (a * a);
            double e2 = 1 - (a2 / (b * b));
            double e4, e6;

            m_TwoPI = M_PI + M_PI;

            e4 = e2 * e2;
            e6 = e4 * e2;

            m_AE = a2 * (1 - e2);

            m_QA = (2.0 / 3.0) * e2;
            m_QB = (3.0 / 5.0) * e4;
            m_QC = (4.0 / 7.0) * e6;

            m_QbarA = -1.0 - (2.0 / 3.0) * e2 - (3.0 / 5.0) * e4 - (4.0 / 7.0) * e6;
            m_QbarB = (2.0 / 9.0) * e2 + (2.0 / 5.0) * e4 + (4.0 / 7.0) * e6;
            m_QbarC = -(3.0 / 25.0) * e4 - (12.0 / 35.0) * e6;
            m_QbarD = (4.0 / 49.0) * e6;

            m_Qp = getQ(M_PI / 2);
            m_E = 4 * M_PI * m_Qp * m_AE;
            if (m_E < 0.0)
                m_E = -m_E;
        }
        static double getQ(double x)
        {
            double sinx, sinx2;

            sinx = Math.Sin(x);
            sinx2 = sinx * sinx;

            return sinx * (1 + sinx2 * (m_QA + sinx2 * (m_QB + sinx2 * m_QC)));
        }
        static double getQbar(double x)
        {
            double cosx, cosx2;

            cosx = Math.Cos(x);
            cosx2 = cosx * cosx;

            return cosx * (m_QbarA + cosx2 * (m_QbarB + cosx2 * (m_QbarC + cosx2 * m_QbarD)));
        }
        static double m_TwoPI, m_AE, m_QA, m_QB, m_QC, m_QbarA, m_QbarB, m_QbarC, m_QbarD, m_Qp, m_E;

        /// <summary>
        /// 根据3度带带号计算中央子午线
        /// </summary>
        /// <param name="zoneNo"></param>
        /// <returns></returns>
        public static double GetMiddleMeridian_3Deg(int zoneNo)
        {
            double mid = 3 * zoneNo;
            return mid;
        }

        /// <summary>
        /// 根据6度带带号计算中央子午线
        /// </summary>
        /// <param name="zoneNo"></param>
        /// <returns></returns>
        public static double GetMiddleMeridian_6Deg(int zoneNo)
        {
            double mid = (6 * zoneNo) - 3;
            return mid;
        }

        /// <summary>
        /// 计算3度带投影带号
        /// </summary>
        /// <param name="longitude"></param>
        /// <returns></returns>
        public static int GetZoneNo_3Deg(double longitude)
        {
            double zoneNo_ = (longitude + 1.5) / 3;
            int zoneNo = (int)zoneNo_;
            return zoneNo;
        }

        /// <summary>
        /// 计算6度带投影带号
        /// </summary>
        /// <param name="longitude"></param>
        /// <returns></returns>
        public static int GetZoneNo_6Deg(double longitude)
        {
            double zoneNo_ = (longitude + 6) / 6;
            int zoneNo = (int)zoneNo_;
            return zoneNo;
        }
    }

    /// <summary>
    /// 地理坐标点类型
    /// </summary>
    public class GeoPoint
    {
        /// <summary>
        /// 经度，横坐标
        /// </summary>
        public double Longitude { get; set; }

        /// <summary>
        /// 纬度，纵坐标
        /// </summary>
        public double Latitude { get; set; }

        /// <summary>
        /// 大地高
        /// </summary>
        public double Height { get; set; }

        public double x()
        {
            return Longitude;
        }
        public double y()
        {
            return Latitude;
        }
    }

    public class CartesianPoint
    {
        public CartesianPoint() { }
        public CartesianPoint(double X, double Y, double Z)
        {
            this.X = X;
            this.Y = Y;
            this.Z = Z;
        }
        public double X { set; get; }
        public double Y { set; get; }
        public double Z { set; get; }
        public string ToDMSString()
        {
            int decimalDigits = 5;
            return this.X.ToDMS(decimalDigits) + ',' + this.Y.ToDMS(decimalDigits) + ',' + this.Z.ToString();
        }
        public override string ToString()
        {
            return this.X.ToString() + ',' + this.Y.ToString() + ',' + this.Z.ToString();
        }
    }

    /// <summary>
    /// This code is a direct derivation from: 
    /// GeoHash Routines for Javascript 2008 (c) David Troy.
    /// The source of which can be found at: 
    /// https://github.com/davetroy/geohash-js
    /// </summary>
    public static class Geohash
    {
        #region Direction enum

        public enum Direction
        {
            Top = 0,
            Right = 1,
            Bottom = 2,
            Left = 3
        }

        #endregion

        private const string Base32 = "0123456789bcdefghjkmnpqrstuvwxyz";
        private static readonly int[] Bits = new[] { 16, 8, 4, 2, 1 };

        private static readonly string[][] Neighbors = 
        {
            new[]
            {
                "p0r21436x8zb9dcf5h7kjnmqesgutwvy", // Top
                "bc01fg45238967deuvhjyznpkmstqrwx", // Right
                "14365h7k9dcfesgujnmqp0r2twvyx8zb", // Bottom
                "238967debc01fg45kmstqrwxuvhjyznp", // Left
            }, 
            new[]
            {
                "bc01fg45238967deuvhjyznpkmstqrwx", // Top
                "p0r21436x8zb9dcf5h7kjnmqesgutwvy", // Right
                "238967debc01fg45kmstqrwxuvhjyznp", // Bottom
                "14365h7k9dcfesgujnmqp0r2twvyx8zb", // Left
            }
        };

        private static readonly string[][] Borders = 
        {
            new[] {"prxz", "bcfguvyz", "028b", "0145hjnp"},
            new[] {"bcfguvyz", "prxz", "0145hjnp", "028b"}
        };

        public static string CalculateAdjacent(String hashCode, Direction direction)
        {
            string hash = hashCode.ToLower();

            char lastChr = hash[hash.Length - 1];
            int type = hash.Length % 2;
            var dir = (int)direction;
            string nHash = hash.Substring(0, hash.Length - 1);

            if (Borders[type][dir].IndexOf(lastChr) != -1)
            {
                nHash = CalculateAdjacent(nHash, (Direction)dir);
            }
            return nHash + Base32[Neighbors[type][dir].IndexOf(lastChr)];
        }

        /// <summary>
        /// 计算周围九个hashcode
        /// </summary>
        /// <param name="hash"></param>
        /// <returns></returns>
        public static string[] NineAdjacentHashCodes(string hash)
        {
            string[] hashCodes = new string[9];
            hashCodes[0] = hash;
            hashCodes[1] = CalculateAdjacent(hashCodes[0], Geohash.Direction.Top);//CenterTop
            hashCodes[2] = CalculateAdjacent(hashCodes[0], Geohash.Direction.Bottom);//CenterBottom
            hashCodes[3] = CalculateAdjacent(hashCodes[1], Geohash.Direction.Left);//TopLeft
            hashCodes[4] = CalculateAdjacent(hashCodes[1], Geohash.Direction.Right);//TopRight
            hashCodes[5] = CalculateAdjacent(hashCodes[0], Geohash.Direction.Left);//CenterLeft
            hashCodes[6] = CalculateAdjacent(hashCodes[0], Geohash.Direction.Right);//CenterRight
            hashCodes[7] = CalculateAdjacent(hashCodes[2], Geohash.Direction.Left);//BottomLeft
            hashCodes[8] = CalculateAdjacent(hashCodes[2], Geohash.Direction.Right);//BottomRight
            return hashCodes;
        }

        public static void RefineInterval(ref double[] interval, int cd, int mask)
        {
            if ((cd & mask) != 0)
            {
                interval[0] = (interval[0] + interval[1]) / 2;
            }
            else
            {
                interval[1] = (interval[0] + interval[1]) / 2;
            }
        }

        /// <summary>
        /// 返回四至范围，Top,Bottom,Left,Right
        /// </summary>
        /// <param name="geohash"></param>
        /// <returns></returns>
        public static double[] Decode(String geohash)
        {
            bool even = true;
            double[] lat = { -90.0, 90.0 };
            double[] lon = { -180.0, 180.0 };

            foreach (char c in geohash)
            {
                int cd = Base32.IndexOf(c);
                for (int j = 0; j < 5; j++)
                {
                    int mask = Bits[j];
                    if (even)
                    {
                        RefineInterval(ref lon, cd, mask);
                    }
                    else
                    {
                        RefineInterval(ref lat, cd, mask);
                    }
                    even = !even;
                }
            }
            return new[] { lat[1], lat[0], lon[0], lon[1] }; //return new[] { (lat[0] + lat[1]) / 2, (lon[0] + lon[1]) / 2 };
        }

        /// <summary>
        /// 根据经纬度生成GeoHash
        /// </summary>
        /// <param name="latitude"></param>
        /// <param name="longitude"></param>
        /// <param name="precision">返回的位数</param>
        /// <returns></returns>
        public static string Encode(double latitude, double longitude, int precision = 12)
        {
            bool even = true;
            int bit = 0;
            int ch = 0;
            string geohash = "";

            double[] lat = { -90.0, 90.0 };
            double[] lon = { -180.0, 180.0 };

            if (precision < 1 || precision > 20) precision = 12;

            while (geohash.Length < precision)
            {
                double mid;

                if (even)
                {
                    mid = (lon[0] + lon[1]) / 2;
                    if (longitude > mid)
                    {
                        ch |= Bits[bit];
                        lon[0] = mid;
                    }
                    else
                        lon[1] = mid;
                }
                else
                {
                    mid = (lat[0] + lat[1]) / 2;
                    if (latitude > mid)
                    {
                        ch |= Bits[bit];
                        lat[0] = mid;
                    }
                    else
                        lat[1] = mid;
                }

                even = !even;
                if (bit < 4)
                    bit++;
                else
                {
                    geohash += Base32[ch];
                    bit = 0;
                    ch = 0;
                }
            }
            return geohash;
        }

        /// <summary>
        /// 根据四至生成GeoHash
        /// </summary>
        /// <param name="latMin"></param>
        /// <param name="latMax"></param>
        /// <param name="lonMin"></param>
        /// <param name="lonMax"></param>
        /// <returns></returns>
        public static string Encode(double latMin, double latMax, double lonMin, double lonMax)
        {
            double detLat = latMax - latMin;
            double detLon = lonMax - lonMin;

            if (detLat < 0 || detLon < 0) throw new ArgumentException("参数错误");

            bool flag = false;
            if (detLat * 2 > detLon) flag = true;

            double lat = (latMax + latMin) / 2;
            double lon = (lonMax + lonMin) / 2;

            double length_ = 0.0;
            if (flag)
                length_ = GreatEllipse.LengthSN(latMax, latMin);
            else
                length_ = GreatEllipse.LengthEW(detLon, latMin);

            string hashCode = string.Empty;
            for (int i = 2; i <= 20; i++)
            {
                hashCode = Encode(lat, lon, i);// Center
                double[] extent = Geohash.Decode(hashCode);
                double length = 0.0;
                if (flag)
                    length = GreatEllipse.LengthSN(extent[0], extent[1]);
                else
                    length = GreatEllipse.LengthEW(extent[3] - extent[2], extent[1]);

                if (length <= length_)
                {
                    hashCode = Encode(lat, lon, i - 1);
                    break;
                }
            }
            return hashCode;
        }

        public static bool IsOutOfBorder(string geoHash, double latMin, double latMax, double lonMin, double lonMax)
        {
            double[] extent = Decode(geoHash);
            bool outBorder =
                             lonMin >= extent[3] //right
                          || lonMax <= extent[2] //left
                          || latMax <= extent[1] //bottom
                          || latMin >= extent[0];//top 
            return outBorder;
        }

        public static List<string> EncodeList(
            double centerLatitude, double centerLongitude,
            double latMin, double latMax, double lonMin, double lonMax,
            int precision = 12)
        {
            List<string> result = new List<string>();

            string center = Encode(centerLatitude, centerLongitude, precision);

            result.Add(center);

            List<string> Ts = new List<string>();
            Ts.Add(center);
            List<string> Bs = new List<string>();
            Bs.Add(center);
            List<string> Ls = new List<string>();
            Ls.Add(center);
            List<string> Rs = new List<string>();
            Rs.Add(center);

            recurseBorders(ref result, Ts, Bs, Ls, Rs, latMin, latMax, lonMin, lonMax);

            return result;
        }

        private static void singleBorder(ref List<string> result,
            ref List<string> newBorder, List<string> existBorder,
            Direction direction,
            double latMin, double latMax, double lonMin, double lonMax)
        {
            if (null == existBorder || existBorder.Count <= 0) return;

            Direction firstDirect = Direction.Top;
            Direction lastDirect = Direction.Bottom;

            if (direction == Direction.Left
                || direction == Direction.Right)
            {
                firstDirect = Direction.Top;
                lastDirect = Direction.Bottom;
            }

            if (direction == Direction.Top
              || direction == Direction.Bottom)
            {
                firstDirect = Direction.Right;
                lastDirect = Direction.Left;
            }

            string tmp = CalculateAdjacent(CalculateAdjacent(existBorder.First(), direction), firstDirect);
            bool outBorder = IsOutOfBorder(tmp, latMin, latMax, lonMin, lonMax);
            if (!outBorder && !newBorder.Contains(tmp)) newBorder.Add(tmp);
            if (!outBorder && !result.Contains(tmp)) result.Add(tmp);
            foreach (string hash in existBorder)
            {
                tmp = CalculateAdjacent(hash, direction);
                outBorder = IsOutOfBorder(tmp, latMin, latMax, lonMin, lonMax);
                if (!outBorder && !newBorder.Contains(tmp)) newBorder.Add(tmp);
                if (!outBorder && !result.Contains(tmp)) result.Add(tmp);
            }
            tmp = CalculateAdjacent(CalculateAdjacent(existBorder.Last(), direction), lastDirect);
            outBorder = IsOutOfBorder(tmp, latMin, latMax, lonMin, lonMax);
            if (!outBorder && !newBorder.Contains(tmp)) newBorder.Add(tmp);
            if (!outBorder && !result.Contains(tmp)) result.Add(tmp);
        }


        private static void recurseBorder(ref List<string> result,
            ref List<string> newBorder, List<string> existBorder,
            Direction direction,
            double latMin, double latMax, double lonMin, double lonMax)
        {
            if (null == existBorder || existBorder.Count <= 0) return;

            foreach (string hash in existBorder)
            {
                string tmp = CalculateAdjacent(hash, Direction.Left);
                bool outBorder = IsOutOfBorder(tmp, latMin, latMax, lonMin, lonMax);
                if (!outBorder && !newBorder.Contains(tmp)) newBorder.Add(tmp);
                if (!outBorder && !result.Contains(tmp)) result.Add(tmp);
            }
        }

        private static void recurseBorders(ref List<string> result,
            List<string> Ts, List<string> Bs, List<string> Ls, List<string> Rs,
              double latMin, double latMax, double lonMin, double lonMax)
        {
            if (null == result) result = new List<string>();

            List<string> _Ts = new List<string>();
            List<string> _Bs = new List<string>();
            List<string> _Ls = new List<string>();
            List<string> _Rs = new List<string>();

            //List<string> _TLs = new List<string>();
            //List<string> _TRs = new List<string>();
            //List<string> _BLs = new List<string>();
            //List<string> _BRs = new List<string>();



            singleBorder(ref result, ref _Ls, Ls, Direction.Left, latMin, latMax, lonMin, lonMax);
            singleBorder(ref result, ref _Rs, Rs, Direction.Right, latMin, latMax, lonMin, lonMax);
            singleBorder(ref result, ref _Ts, Ts, Direction.Top, latMin, latMax, lonMin, lonMax);
            singleBorder(ref result, ref _Bs, Bs, Direction.Bottom, latMin, latMax, lonMin, lonMax);

            if (_Ts.Count == 0
                && _Bs.Count == 0
                && _Ls.Count == 0
                && _Rs.Count == 0)
                return;

            recurseBorders(ref result, _Ts, _Bs, _Ls, _Rs, latMin, latMax, lonMin, lonMax);
        }

        /// <summary>
        /// 计算哈希串代表地理范围的底边长度（单位：米）
        /// </summary>
        /// <param name="geohash"></param>
        /// <returns></returns>
        public static double CalculateWidth(string geohash)
        {
            double[] extent = Geohash.Decode(geohash);
            double ew = GreatEllipse.LengthEW(extent[3] - extent[2], extent[1]);
            return ew;
        }

        /// <summary>
        /// 计算哈希串代表地理范围的高度（单位：米）
        /// </summary>
        /// <param name="geohash"></param>
        /// <returns></returns>
        public static double CalculateHeight(string geohash)
        {
            double[] extent = Geohash.Decode(geohash);
            double sn = GreatEllipse.LengthSN(extent[0], extent[1]);
            return sn;
        }

    }
}
#endregion

#region System.Hardware
namespace System.Hardware
{
    public class MemoryHelper
    {
        /// <summary>
        /// 内存回收
        /// </summary>
        /// <param name="process"></param>
        /// <param name="minSize"></param>
        /// <param name="maxSize"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize")]
        public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);

        /// <summary>
        /// 释放内存
        /// </summary>
        public static void Clear()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }
        }
    }


    [SerializableAttribute]
    public class Machine
    {
        /// <summary>
        /// 获取机器序列编码
        /// </summary>
        public static string GetSerialString()
        {
            Device d1 = new Device(EnumDeviceType.Processor);
            Device d2 = new Device(EnumDeviceType.HardDrive);

            object obj1 = d1.Properties["ProcessorId"].Value;
            object obj2 = d2.Properties["SerialNumber"].Value;

            return (obj1.ToString() + obj2.ToString()).ToMD5String();
        }
    }

    internal enum EnumDeviceType
    {
        Processor = 0,
        HardDrive = 1,
        Motherboard = 2
    }

    internal class Device
    {
        private PropertyDataCollection properties = null;
        internal PropertyDataCollection Properties { get { return properties; } }

        internal Device(EnumDeviceType type)
        {
            string win32Class = string.Empty;
            switch (type)
            {
                case EnumDeviceType.Processor:
                    win32Class = "Win32_Processor";
                    break;
                case EnumDeviceType.HardDrive:
                    win32Class = "Win32_DiskDrive";
                    break;
                case EnumDeviceType.Motherboard:
                    win32Class = "Win32_MotherboardDevice";
                    break;
            }
            ManagementClass management = new ManagementClass(win32Class);
            ManagementObjectCollection objCollect = management.GetInstances();

            foreach (ManagementObject obj in objCollect)
            {
                properties = obj.Properties;
                break;
            }
        }

        internal void Print()
        {
            PropertyDataCollection.PropertyDataEnumerator enumor = properties.GetEnumerator();
            enumor.Reset();
            while (enumor.MoveNext())
            {
                PropertyData property = enumor.Current;
                Console.WriteLine(string.Format("Property:[{0},{1}]", property.Name, property.Value));
            }
        }
    }
}
#endregion

#region System.IO
namespace System.IO
{
    /// <summary>
    /// 扩展System.IO命名空间下的类
    /// </summary>
    public static class ExtendSystemIO
    {

        /// <summary>
        /// 获取路径下的文件
        /// </summary>
        /// <param name="path"></param>
        /// <param name="searchPatterns"></param>
        /// <returns></returns>
        public static string[] GetFiles(this string path, params string[] searchPatterns)
        {
            if (path == null || searchPatterns.Length <= 0 || !System.IO.Directory.Exists(path))
            {
                return new string[0];
            }
            else
            {
                DirectoryInfo di = new DirectoryInfo(path);
                List<FileInfo[]> lst = new List<FileInfo[]>();
                for (int i = 0; i < searchPatterns.Length; i++)
                {
                    FileInfo[] fileInfos = di.GetFiles(searchPatterns[i]);
                    lst.Add(fileInfos);
                }
                List<string> files = new List<string>();
                foreach (FileInfo[] fis in lst)
                {
                    foreach (FileInfo fi in fis)
                    {
                        if (!files.Contains(fi.FullName)) files.Add(fi.FullName);
                    }
                }
                return files.ToArray();
            }
        }

        /// <summary>
        /// 文件夹字符串后面添加文件名，返回全路径
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string AppendFileName(this string folder, string fileName)
        {
            if (!System.IO.Directory.Exists(folder)) throw new DirectoryNotFoundException("不存在路径：" + folder);
            return folder.TrimEnd('\\') + '\\' + fileName;
        }
    }

    /// <summary>
    /// 文件异步复制类，可以用于复制大文件
    /// 创建人：谢成磊
    /// 创建时间：2015.11.13
    /// </summary>
    public static class FileHelper
    {
        /// <summary>
        /// 文件异步复制,可复制大文件。用于共享目录内的文件拷贝
        /// </summary>
        /// <param name="SourceFile">源文件</param>
        /// <param name="TargetFile">目标文件</param>
        /// <param name="BufferMB">缓冲大小（单位：MB）</param>
        /// <param name="Speed">复制平均速度(Byte Per Sec)</param>
        /// <param name="Progress">复制进度</param>
        /// <returns>文件校验结果</returns>
        public static bool AsyncCopy(string SourceFile, string TargetFile, int BufferMB, ref double Speed, ref int Progress)
        {
            int bufferSize = BufferMB * 1048576;//缓冲大小，单位兆

            System.IO.FileStream fsr = new System.IO.FileStream(SourceFile, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None, bufferSize, true);
            System.IO.FileStream fsw = new System.IO.FileStream(TargetFile, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.None, bufferSize, true);

            byte[] buffer = new byte[bufferSize];
            System.Int64 len = fsr.Length;
            System.DateTime start = System.DateTime.Now;
            System.TimeSpan ts;
            while (fsr.Position < len)
            {
                int readCount = fsr.Read(buffer, 0, bufferSize);
                fsw.Write(buffer, 0, readCount);

                ts = System.DateTime.Now.Subtract(start);
                Speed = fsr.Position / ts.TotalMilliseconds * 1000;
                Progress = (int)(fsr.Position * 100 / len);
            }
            fsr.Close();
            fsw.Close();

            return GetFileMD5(SourceFile) == GetFileMD5(TargetFile);
        }

        /// <summary>
        /// 获取文件的MD5码  
        /// </summary>  
        /// <param name="fileName">传入的文件名(全路径)</param>  
        /// <returns>MD5值</returns>  
        public static string GetFileMD5(string fileName)
        {
            try
            {
                FileStream file = new FileStream(fileName, FileMode.Open);
                System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                byte[] retVal = md5.ComputeHash(file);
                file.Close();
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                }
                return sb.ToString();
            }
            catch (Exception e)
            {
                throw new Exception("MD5Helper.GetMD5FromFile(string) Failed, ErrMsg:" + e.Message);
            }
        }

        private const int OF_READWRITE = 2;
        private const int OF_SHARE_DENY_NONE = 0x40;
        private static readonly IntPtr HFILE_ERROR = new IntPtr(-1);

        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        /// <summary>
        /// 判断文件是否被占用
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool IsOccupied(string fileName)
        {
            IntPtr vHandle = _lopen(fileName, OF_READWRITE | OF_SHARE_DENY_NONE);
            return vHandle == HFILE_ERROR;
        }


    }

    /// <summary>
    /// 对象IO接口
    /// </summary>
    public interface IObjectIO
    {
        object Load(string fileName);
        bool Save(string fileName);
    }
}
#endregion

#region System.MapServer
namespace System.MapServer
{
    public class MapServerCapabilities
    {
        public string Title { get; set; }

        public ResourceURL ResourceURL { get; set; }

        public string LowerCorner { get; set; }

        public string UpperCorner { get; set; }

        public Dictionary<int, double> Resolutions { get; set; }

        public double GetScale(int Level)
        {
            if (this.Resolutions == null) return double.NaN;
            if (this.Resolutions.ContainsKey(Level))
                return this.Resolutions[Level];
            return double.NaN;
        }
    }

    public class ResourceURL
    {
        public string Format;
        public string ResourceType;
        public string Template;
    }

    public class MapServerHelper
    {
        public static MapServerCapabilities ParseXML(System.Xml.XmlDocument xml)
        {
            MapServerCapabilities msc = new MapServerCapabilities();
            if (null == xml) return msc;
            try
            {
                System.Xml.XmlElement root = xml.DocumentElement;

                foreach (System.Xml.XmlNode node2 in root.ChildNodes)
                {
                    if (node2.Name == "Contents")
                    {
                        foreach (System.Xml.XmlNode node3 in node2.ChildNodes)
                        {
                            if (node3.Name == "TileMatrixSet")
                            {
                                if (null == msc.Resolutions) msc.Resolutions = new Dictionary<int, double>();

                                foreach (System.Xml.XmlNode node4 in node3.ChildNodes)
                                {
                                    if (node4.Name == "TileMatrix")
                                    {
                                        Xml.XmlNode level = null;
                                        Xml.XmlNode scale = null;

                                        foreach (System.Xml.XmlNode tmp in node4.ChildNodes)
                                        {
                                            if (tmp.Name == "ows:Identifier") { level = tmp; continue; }
                                            if (tmp.Name == "ScaleDenominator") { scale = tmp; continue; }
                                        }
                                        string LevelStr = level.FirstChild.Value;
                                        string ScaleDenominatorStr = scale.FirstChild.Value;

                                        double ScaleDenominator = double.Parse(ScaleDenominatorStr);
                                        int Level = int.Parse(LevelStr);

                                        msc.Resolutions.Add(Level, ScaleDenominator);
                                    }
                                }
                            }

                            if (node3.Name == "Layer")
                            {
                                System.Xml.XmlNode layer = node3;
                                System.Xml.XmlNode title = null;
                                foreach (System.Xml.XmlNode tmp in layer.ChildNodes)
                                {
                                    if (tmp.Name == "ows:Title") { title = tmp; break; }
                                }

                                System.Xml.XmlNode ResourceURL = null;
                                foreach (System.Xml.XmlNode tmp in layer.ChildNodes)
                                {
                                    if (tmp.Name == "ResourceURL") { ResourceURL = tmp; break; }
                                }

                                msc.Title = title.FirstChild.Value;

                                ResourceURL url = new ResourceURL();
                                if (null != ResourceURL)
                                {
                                    url.Format = ResourceURL.Attributes["format"].Value;
                                    url.ResourceType = ResourceURL.Attributes["resourceType"].Value;
                                    url.Template = ResourceURL.Attributes["template"].Value;
                                }
                                msc.ResourceURL = url;

                                System.Xml.XmlNode LowerCorner = null;
                                System.Xml.XmlNode UpperCorner = null;

                                foreach (System.Xml.XmlNode tmp in layer.ChildNodes)
                                {
                                    if (tmp.Name == "ows:BoundingBox")
                                    {
                                        foreach (System.Xml.XmlNode tmp2 in tmp.ChildNodes)
                                        {
                                            if (tmp2.Name == "ows:LowerCorner")
                                            {
                                                LowerCorner = tmp2;
                                            }
                                            if (tmp2.Name == "ows:UpperCorner")
                                            {
                                                UpperCorner = tmp2;
                                            }
                                        }
                                    }
                                }

                                msc.LowerCorner = LowerCorner.FirstChild.Value;
                                msc.UpperCorner = UpperCorner.FirstChild.Value;

                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            { }
            return msc;
        }

        /// <summary>
        /// 获取一般OGC WMTS服务的Capabilities
        /// http://t0.tianditu.com/cva_c/wmts
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static MapServerCapabilities GetCapabilities_WMTS_OGC(string url)
        {
            MapServerCapabilities r = null;
            try
            {
                System.Xml.XmlDocument xml = GetCapabilities(url, "?service=WMTS&request=GetCapabilities");
                r = ParseXML(xml);
            }
            catch
            { }
            return r;
        }

        /// <summary>
        /// 获取ArcServer WMTS 服务的 Capabilities
        /// http://192.168.1.207:6080/arcgis/rest/services//58ji1/MapServer
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static MapServerCapabilities GetCapabilities_WMTS_ArcServer(string url)
        {
            MapServerCapabilities r = null;
            try
            {
                System.Xml.XmlDocument xml = GetCapabilities(url, "/WMTS/1.0.0/WMTSCapabilities.xml");
                r = ParseXML(xml);
            }
            catch
            { }
            return r;
        }

        public static System.Xml.XmlDocument GetCapabilities(
            string url,
            string parameters = "?service=WMTS&request=GetCapabilities")
        {
            System.Net.CookieCollection cookies = new System.Net.CookieCollection();
            System.Net.HttpWebResponse response = System.Net.HttpWebResponseHelper.CreateGetHttpResponse(
                string.Format("{0}{1}", url, parameters),
                null,
                null,
                cookies);

            Stream responseStream = null;
            StreamReader sReader = null;

            try
            {
                // 获取响应流
                responseStream = response.GetResponseStream();
                // 对接响应流(以"utf-8"字符集)
                sReader = new StreamReader(responseStream, System.Text.Encoding.UTF8);
                // 开始读取数据
                string value = sReader.ReadToEnd();

                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                doc.LoadXml(value);

                return doc;
            }
            catch (Exception)
            {
                //日志异常
            }
            finally
            {
                if (null != sReader)
                {
                    sReader.Close();
                }
                if (null != responseStream)
                {
                    responseStream.Close();
                }
                if (null != response)
                {
                    response.Close();
                }
            }
            return null;
        }

        /// <summary>
        /// 天地图POI查询
        /// </summary>
        /// <param name="keyword">查询关键字</param>
        /// <param name="xmin"></param>
        /// <param name="ymin"></param>
        /// <param name="xmax"></param>
        /// <param name="ymax"></param>
        /// <param name="start"></param>
        /// <param name="count"></param>
        /// <param name="queryTerminal"></param>
        /// <param name="level"></param>
        /// <returns></returns>
        public static List<POI> QueryPOI_TDT(
            string keyword,
            double xmin, double ymin,
            double xmax, double ymax,
            int start = 0, int count = 1000,
            int queryTerminal = 1000,
            int level = 18)
        {
            Stream responseStream = null;
            StreamReader sReader = null;
            System.Net.HttpWebResponse response = null;
            List<POI> pois = new List<POI>();

            try
            {
                string mapBound = string.Format("{0},{1},{2},{3}", xmin, ymin, xmax, ymax);
                string url = "http://map.tianditu.com/query.shtml";
                string param = string.Format("postStr={{\"keyWord\":\"{0}\",\"level\":\"{1}\",\"mapBound\":\"{2}\",\"queryType\":\"1\",\"count\":\"{3}\",\"start\":\"{4}\",\"queryTerminal\":\"{5}\"}}&type=query",
                    keyword, level, mapBound, count, start, queryTerminal);

                // HttpUtility.UrlEncode(tmp);
                string fullURL = string.Format("{0}?{1}", url, param);

                System.Net.CookieCollection cookies = new System.Net.CookieCollection();

                response = System.Net.HttpWebResponseHelper.CreateGetHttpResponse(
                    fullURL,
                    null,
                    null,
                    cookies);

                // 获取响应流
                responseStream = response.GetResponseStream();
                // 对接响应流(以"utf-8"字符集)
                sReader = new StreamReader(responseStream, System.Text.Encoding.UTF8);
                // 开始读取数据
                string value = sReader.ReadToEnd();

                Dictionary<string, object> objects = value.FromJson<object>() as Dictionary<string, object>;

                List<object> poiObjs = new List<object>();

                if (objects.ContainsKey("pois")) poiObjs = objects["pois"] as List<object>;

                foreach (object poiObj in poiObjs)
                {
                    Dictionary<string, object> feaDic = poiObj as Dictionary<string, object>;
                    if (feaDic == null) continue;

                    POI poi = new POI();
                    poi.Name = feaDic["name"].ToString();
                    poi.Address = feaDic["address"].ToString();
                    string lonlat = feaDic["lonlat"].ToString();
                    string[] xy = lonlat.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    poi.X = double.Parse(xy[0]);
                    poi.Y = double.Parse(xy[1]);

                    pois.Add(poi);
                }
                return pois;
            }
            catch (Exception)
            {
                //日志异常
            }
            finally
            {
                if (null != sReader)
                {
                    sReader.Close();
                }
                if (null != responseStream)
                {
                    responseStream.Close();
                }
                if (null != response)
                {
                    response.Close();
                }
            }
            return pois;
        }
    }

    public struct BackGroundLayerConfig
    {
        public string Name;
        public string URL;
        public string Type;
        public int LevelMin;
        public int LevelMax;

        public static BackGroundLayerConfig Parse(string name, string cfgTxt)
        {
            BackGroundLayerConfig cfg = new BackGroundLayerConfig();
            cfg.Name = name;

            string[] parameters = cfgTxt.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            try
            {
                foreach (string param in parameters)
                {
                    string[] kv = param.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);
                    if (kv.Length < 2) continue;

                    string key = kv[0].ToUpper();
                    switch (key)
                    {
                        case "URL": cfg.URL = kv[1]; break;
                        case "TYPE": cfg.Type = kv[1]; break;
                        case "LEVELMIN": cfg.LevelMin = int.Parse(kv[1]); break;
                        case "LEVELMAX": cfg.LevelMax = int.Parse(kv[1]); break;
                    }

                }
            }
            catch (Exception)
            {
            }
            return cfg;
        }

        public override string ToString()
        {
            string result = "";
            System.Reflection.PropertyInfo[] props = this.GetType().GetProperties();

            for (int i = 0; i < props.Length; i++)
            {
                System.Reflection.PropertyInfo prop = props[i];
                object val = this.GetPropertyValue(prop.Name);
                if (null != val)
                {
                    if (string.IsNullOrWhiteSpace(result))
                        result += prop.Name.ToUpper() + "=" + val.ToString();
                    else
                        result += ";" + prop.Name.ToUpper() + "=" + val.ToString();

                }
            }
            return result;
        }
    }

    public struct POI
    {
        public string Name;
        public string Address;
        public double X;
        public double Y;

        public override string ToString()
        {
            return Name;
        }
    }



}
#endregion

#region System.Net
namespace System.Net
{
    /// <summary>  
    /// 有关HTTP请求的辅助类  
    /// </summary>  
    public class HttpWebResponseHelper
    {
        private static readonly string DefaultUserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";

        /// <summary>  
        /// 创建GET方式的HTTP请求  
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="timeout">请求的超时时间</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>  
        public static HttpWebResponse CreateGetHttpResponse(string url, int? timeout, string userAgent, CookieCollection cookies)
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Method = "GET";
            request.UserAgent = DefaultUserAgent;
            request.ContentType = "application/x-www-form-urlencoded;charset=utf8"; // 必须指明字符集！

            if (!string.IsNullOrEmpty(userAgent))
            {
                request.UserAgent = userAgent;
            }
            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }
            return request.GetResponse() as HttpWebResponse;
        }

        /// <summary>  
        /// 创建POST方式的HTTP请求  
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="parameters">随同请求POST的参数名称及参数值字典</param>  
        /// <param name="timeout">请求的超时时间</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="requestEncoding">发送HTTP请求时所用的编码</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>  
        public static HttpWebResponse CreatePostHttpResponse(string url, IDictionary<string, string> parameters, int? timeout, string userAgent, System.Text.Encoding requestEncoding, CookieCollection cookies)
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            if (requestEncoding == null)
            {
                throw new ArgumentNullException("requestEncoding");
            }
            HttpWebRequest request = null;
            //如果是发送HTTPS请求  
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback
                    = new System.Net.Security.RemoteCertificateValidationCallback(CheckValidationResult);
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                request = WebRequest.Create(url) as HttpWebRequest;
            }
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            if (!string.IsNullOrEmpty(userAgent))
            {
                request.UserAgent = userAgent;
            }
            else
            {
                request.UserAgent = DefaultUserAgent;
            }

            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }
            //如果需要POST数据  
            if (!(parameters == null || parameters.Count == 0))
            {
                System.Text.StringBuilder buffer = new System.Text.StringBuilder();
                int i = 0;
                foreach (string key in parameters.Keys)
                {
                    if (i > 0)
                    {
                        buffer.AppendFormat("&{0}={1}", key, parameters[key]);
                    }
                    else
                    {
                        buffer.AppendFormat("{0}={1}", key, parameters[key]);
                    }
                    i++;
                }
                byte[] data = requestEncoding.GetBytes(buffer.ToString());
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            return request.GetResponse() as HttpWebResponse;
        }

        /// <summary>  
        /// 创建POST方式的HTTP请求
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="parameters">随同请求POST的参数名称及参数值字典</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>
        public static HttpWebResponse CreatePostHttpResponse(
            string url, string parameters, CookieCollection cookies)
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }

            HttpWebRequest request = null;
            Stream stream = null;//用于传参数的流

            try
            {
                //如果是发送HTTPS请求  
                if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                {
                    ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(CheckValidationResult);
                    request = WebRequest.Create(url) as HttpWebRequest;
                    //创建证书文件
                    System.Security.Cryptography.X509Certificates.X509Certificate objx509
                        = new System.Security.Cryptography.X509Certificates.X509Certificate(System.Windows.Forms.Application.StartupPath + @"\\licensefile\zjs.cer");
                    //添加到请求里
                    request.ClientCertificates.Add(objx509);
                    request.ProtocolVersion = HttpVersion.Version10;
                }
                else
                {
                    request = WebRequest.Create(url) as HttpWebRequest;
                }

                request.Method = "POST";//传输方式
                request.ContentType = "application/x-www-form-urlencoded";//协议                
                request.UserAgent = DefaultUserAgent;//请求的客户端浏览器信息,默认IE                
                request.Timeout = 6000;//超时时间，写死6秒

                //随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空
                if (cookies != null)
                {
                    request.CookieContainer = new CookieContainer();
                    request.CookieContainer.Add(cookies);
                }

                //如果需求POST传数据，转换成utf-8编码
                byte[] data = System.Text.Encoding.UTF8.GetBytes(parameters);
                request.ContentLength = data.Length;

                stream = request.GetRequestStream();
                stream.Write(data, 0, data.Length);

                stream.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }
            }

            return request.GetResponse() as HttpWebResponse;
        }

        private static bool CheckValidationResult(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors errors)
        {
            return true; //总是接受  
        }
    }


}
#endregion

#region System.Net.Mail
namespace System.Net.Mail
{
    /// <summary>
    /// 邮件类
    /// </summary>
    public class EMail
    {
        /// <summary>
        /// 发送邮件
        /// 163测试通过
        /// </summary>
        /// <param name="receiveAddresses">收件人，多个收件人用','隔开</param>
        /// <param name="mailSubject">邮件主题</param>
        /// <param name="mailBody">邮件正文</param>
        /// <param name="mailHost">发送服务器地址</param>
        /// <param name="sendAddress">发件人</param>
        /// <param name="password">发件人密码</param>
        /// <param name="IsBodyHtml">指示正文是否为HTML，可选</param>
        /// <param name="attachmentFiles">附件，可选</param>
        /// <returns></returns>
        public bool Send(
           string receiveAddresses,
           string mailSubject,
           string mailBody,
           string mailHost,
           string sendAddress,
           string password,
           bool IsBodyHtml = false,
           List<string> attachmentFiles = null)
        {
            if (!sendAddress.Contains('@')) return false;
            string displayName = sendAddress.Trim().Split('@')[0];

            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();

            client.Host = mailHost;//client.Port =
            client.UseDefaultCredentials = true;
            client.Credentials = new System.Net.NetworkCredential(displayName, password);
            client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;

            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();

            message.To.Add(receiveAddresses);
            message.From = new System.Net.Mail.MailAddress(sendAddress, displayName, System.Text.Encoding.UTF8);
            message.Subject = mailSubject;
            message.Body = mailBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = IsBodyHtml;
            message.Priority = System.Net.Mail.MailPriority.High;

            if (attachmentFiles != null)
                foreach (string attFile in attachmentFiles)
                {
                    if (System.IO.File.Exists(attFile))
                    {
                        System.Net.Mail.Attachment data = new System.Net.Mail.Attachment(attFile);
                        message.Attachments.Add(data);
                    }
                }

            client.Send(message);
            return true;
        }
    }
}
#endregion

#region System.Reflection
namespace System.Reflection
{
    public static class SystemReflection
    {
        /// <summary>
        /// 获取此程序集中定义的类型。
        /// </summary>
        /// <param name="dllPath"></param>
        /// <returns></returns>
        public static Type[] GetTypes(this string dllPath)
        {
            if (!File.Exists(dllPath)) throw new FileNotFoundException(string.Format("找不到文件：{0}。请检查配置", dllPath));
            string ext = Path.GetExtension(dllPath).ToUpper();
            if (!ext.Equals(".DLL")) throw new DllNotFoundException("此文件不是DLL：" + dllPath);

            Type[] result = new Type[0];
            try
            {
                Assembly assemble = Assembly.LoadFile(dllPath);
                result = assemble.GetTypes();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return result;
        }

        /// <summary>
        /// 由dll路径和类全名获取实例
        /// </summary>
        /// <param name="dllPath"></param>
        /// <param name="typeName">要查找的类型的 System.Type.FullName。</param>
        /// <param name="ignoreCase">如果为 true，则忽略类型名的大小写；否则，为 false。</param>
        /// <returns></returns>
        public static object GetInstance(this string dllPath, string typeName, bool ignoreCase)
        {
            if (!File.Exists(dllPath)) throw new FileNotFoundException(string.Format("找不到文件：{0}。请检查配置", dllPath));
            string ext = Path.GetExtension(dllPath).ToUpper();
            if (!ext.Equals(".DLL")) throw new DllNotFoundException("此文件不是DLL：" + dllPath);
            Assembly assemble = Assembly.LoadFile(dllPath);
            return assemble.CreateInstance(typeName, ignoreCase);
        }

    }
}
#endregion

#region System.Windows.Forms
namespace System.Windows.Forms
{
    /// <summary>
    /// System.Windows.Forms
    /// </summary>
    public static class ExtendSystemWindowsForms
    {
        /// <summary>
        /// 绑定数据（前提是Control的Tag属性值和数据源对象的属性名一致，不区分大小写）
        /// </summary>
        /// <param name="ctrlCollect"></param>
        /// <param name="data">数据源对象</param>
        public static void BindingObjectData(this Control.ControlCollection ctrlCollect, object data)
        {
            foreach (Control ctrl in ctrlCollect)
            {
                ctrl.BindingObjectData(data);
            }
        }

        public static void BindingObjectData(this Control ctrl, object data)
        {
            PropertyInfo[] props = data.GetType().GetProperties();

            string ctrlBindingStr = string.Empty;
            if (ctrl.Tag != null) ctrlBindingStr = ctrl.Tag.ToString().Trim();
            IEnumerable<PropertyInfo> queryResult = props.Where((prop, b) =>
            {
                return prop.Name.ToLower().Equals(ctrlBindingStr.ToLower());
            });
            if (queryResult.Count() > 0)
            {
                PropertyInfo prop = queryResult.First();
                Type curType = ctrl.GetType();
                string ctrlPropertyName = curType.GetCurCtrlValuePropertyName();
                Binding binding = ctrl.DataBindings[ctrlPropertyName];
                if (binding != null) ctrl.DataBindings.Remove(binding);
                ctrl.DataBindings.Add(
                    ctrlPropertyName,
                    data,
                    prop.Name,
                    true,
                    DataSourceUpdateMode.OnValidation);
            }

        }

        public static string GetCurCtrlValuePropertyName(this Type curInsType)
        {
            string propertyName = "Text";
            if (curInsType == null) return propertyName;

            Console.WriteLine(curInsType.FullName);

            string curInsTypeFullName = curInsType.FullName;
            if (curInsTypeFullName.Equals("DevExpress.XtraEditors.TextEdit"))
                propertyName = "Text";

            if (curInsTypeFullName.Equals("DevExpress.XtraEditors.DateEdit"))
                propertyName = "DateTime";

            if (curInsTypeFullName.Equals("DevExpress.XtraEditors.ComboBoxEdit"))
                propertyName = "Text";

            if (curInsTypeFullName.Equals("System.Windows.Forms.CheckBox"))
                propertyName = "Checked";

            if (curInsTypeFullName.Equals("System.Windows.Forms.DateTimePicker"))
                propertyName = "Value";

            return propertyName;
        }

        /// <summary>
        /// 根据扩展名获取ImageIndex
        /// </summary>
        /// <param name="FileIconCollection"></param>
        /// <param name="Extention">文件扩展名（包含句点）</param>
        /// <returns></returns>
        public static int GetImageIndex(this ImageList FileIconCollection, string Extention)
        {
            string EXT = Extention.ToUpper();
            int result = 0;

            if (!FileIconCollection.Images.Keys.Contains(EXT))
            {
                System.Drawing.Icon icon = null;
                System.Drawing.IconHelper.GetFileIcon(EXT, out icon, false);
                FileIconCollection.Images.Add(EXT, icon);
            }
            result = FileIconCollection.Images.Keys.IndexOf(EXT);
            return result;
        }

        /// <summary>
        /// 保存所有图标
        /// </summary>
        /// <param name="FileIconCollection"></param>
        /// <param name="SaveFolder">保存文件夹</param>
        public static void SaveImages(this ImageList FileIconCollection, string SaveFolder)
        {
            for (int i = 0; i < FileIconCollection.Images.Count; i++)
            {
                System.Drawing.Image img = FileIconCollection.Images[i];
                img.Save(SaveFolder + "\\" + FileIconCollection.Images.Keys[i], System.Drawing.Imaging.ImageFormat.Icon);
            }
        }

        /// <summary>
        /// 使文本始终滚动到最底部
        /// </summary>
        /// <param name="edit"></param>
        public static void ScrollToBottomAlways(this TextBoxBase edit)
        {
            edit.TextChanged += TextBox_TextChanged;
        }

        /// <summary>
        /// 取消文本滚动到最底部
        /// </summary>
        /// <param name="edit"></param>
        public static void ScrollToBottomCancel(this TextBoxBase edit)
        {
            edit.TextChanged -= TextBox_TextChanged;
        }

        private static void TextBox_TextChanged(object sender, EventArgs e)
        {
            TextBoxBase edit = sender as TextBoxBase;
            edit.SelectionStart = edit.Text.Length;
            edit.ScrollToCaret();
        }

        /// <summary>
        /// 探测文件夹是否存在。在Text属性变化时，判断Text是不是文件夹，如果不是则弹出提示窗口
        /// </summary>
        /// <param name="edit"></param>
        public static void CheckValueIsFolderOrNotAlways(this TextBoxBase edit)
        {
            edit.TextChanged += TextBox_CheckIsFolderOrNot;
        }

        /// <summary>
        /// 取消文件夹是否存在的探测
        /// </summary>
        /// <param name="edit"></param>
        public static void CheckValueIsFolderOrNotCancel(this TextBoxBase edit)
        {
            edit.TextChanged -= TextBox_CheckIsFolderOrNot;
        }

        private static void TextBox_CheckIsFolderOrNot(object sender, EventArgs e)
        {
            TextBoxBase edit = sender as TextBoxBase;
            if (string.IsNullOrEmpty(edit.Text)) return;
            if (!System.IO.Directory.Exists(edit.Text))
            {
                MessageBox.Show("不存在路径" + edit.Text, "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                edit.Text = string.Empty;
            }
        }

        /// <summary>
        /// 探测文件是否存在。在Text属性变化时，判断Text是不是文件，如果不是则弹出提示窗口
        /// </summary>
        /// <param name="edit"></param>
        public static void CheckValueIsFileOrNotAlways(this TextBoxBase edit)
        {
            edit.TextChanged += TextBox_CheckIsFileOrNot;
        }

        /// <summary>
        /// 取消文件是否存在的探测
        /// </summary>
        /// <param name="edit"></param>
        public static void CheckValueIsFileOrNotCancel(this TextBoxBase edit)
        {
            edit.TextChanged -= TextBox_CheckIsFileOrNot;
        }

        private static void TextBox_CheckIsFileOrNot(object sender, EventArgs e)
        {
            TextBoxBase edit = sender as TextBoxBase;
            if (string.IsNullOrEmpty(edit.Text)) return;
            if (!System.IO.File.Exists(edit.Text))
            {
                MessageBox.Show("不存在文件" + edit.Text, "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                edit.Text = string.Empty;
            }
        }

        /// <summary>
        /// 弹出选择文件夹对话框
        /// </summary>
        /// <param name="edit"></param>
        /// <param name="description"></param>
        public static void GetFolderPath(this TextBoxBase edit, string description = "")
        {
            edit.MouseClick -= FolderText_MouseClick;
            edit.MouseClick += FolderText_MouseClick;

            FolderBrowserDialog browser = new FolderBrowserDialog() { Description = description };
            if (System.IO.Directory.Exists(edit.Text)) browser.SelectedPath = edit.Text;
            if (browser.ShowDialog() == DialogResult.OK) edit.Text = browser.SelectedPath;
        }

        static void FolderText_MouseClick(object sender, MouseEventArgs e)
        {
            TextBoxBase edit = sender as TextBoxBase;
            if (edit != null) edit.SelectAll();
        }

        /// <summary>
        /// 弹出选择文件对话框
        /// </summary>
        /// <param name="edit"></param>
        /// <param name="title"></param>
        public static void GetFilePath(this TextBoxBase edit, string title, string filter)
        {
            OpenFileDialog openfile = new OpenFileDialog()
            {
                Multiselect = false,
                Title = title,
                Filter = filter
            };
            if (openfile.ShowDialog() == DialogResult.OK)
                edit.Text = openfile.FileName;
        }

        /// <summary>
        /// 弹出保存文件对话框
        /// </summary>
        /// <param name="edit"></param>
        /// <param name="title"></param>
        /// <param name="filter"></param>
        public static void GetFileSavePath(this TextBoxBase edit, string title, string filter)
        {
            SaveFileDialog savefile = new SaveFileDialog()
            {
                Title = title,
                Filter = filter
            };
            if (savefile.ShowDialog() == DialogResult.OK)
                edit.Text = savefile.FileName;
        }
    }

    /// <summary>
    /// 窗口帮助类
    /// 创建人：谢成磊
    /// 创建时间：2015.08.25
    /// </summary>
    public static class WinFormHelper
    {
        /// <summary>
        /// 移动窗口
        /// 使用方法：在窗体_MouseMove事件中添加“WinFormHelper.MoveFormBody(this.Handle);”
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static bool MoveFormBody(IntPtr hwnd)
        {
            if (ReleaseCapture())
            {
                if (SendMessage(hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0))
                {
                    return true;
                }
            }
            return false;
        }

        #region 私有字段和方法

        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_MOVE = 0xF010;
        private const int HTCAPTION = 0x0002;

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern bool SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

        #endregion

    }

}
#endregion
