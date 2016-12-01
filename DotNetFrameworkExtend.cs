using System.Collections.Generic;
using System.Configuration;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Win32;

/// .Net Framework 功能扩展
/// 创建人：谢成磊
/// 创建时间：2016/06/05
/// 最后修改时间：2016/11/16
/// 2016/11/16 添加数据库扩展 System.Data 命名空间

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
        public static void Open(string filePath)
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

        /// <summary>
        /// 把字符串解析为double
        /// </summary>
        /// <param name="s"></param>
        /// <param name="value">输出的数值</param>
        public static double ParseToDouble(this string s)
        {
            double tmp = double.MinValue;
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
            int? result = null;
            double tmp = double.MinValue;
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
            string[] dms = dmsString.Split("°′″".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            if (dms.Length < 3) return double.NaN;
            double d = dms[0].ParseToDouble();
            double m = dms[1].ParseToDouble();
            double s = dms[2].ParseToDouble();
            return d + m / 60.0 + s / 3600.0;
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
        /// 字符串解密
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
        /// 属性赋值(只有属性名称一致才会赋值)
        /// </summary>
        /// <param name="tarObj">目标对象</param>
        /// <param name="srcObj">源对象</param>
        public static void GetAttributesValueFrom(this object tarObj, object srcObj)
        {
            if (srcObj == null) return;

            PropertyInfo[] src = srcObj.GetType().GetProperties();
            PropertyInfo[] tar = tarObj.GetType().GetProperties();

            for (int ii = 0; ii < tar.Length; ii++)
            {
                PropertyInfo tarProp = tar[ii];
                string tName = tarProp.Name;

                IEnumerable<PropertyInfo> srcProps = src.Where(e => { return e.Name == tName; });
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
        public enum DataBaseType
        {
            DEFAULT = 1,
            ORACLE = 1,
            MSACCESS = 2,
            MSSQL = 3,
            SQLITE = 4
        }

        /// <summary>
        /// 创建数据库连接
        /// </summary>
        /// <param name="connectStr"></param>
        /// <param name="type">MSSQL|Oracle|MSAccess</param>
        /// <returns></returns>
        public static IDbConnection CreateDB(this DataBaseType type, string connectStr)
        {
            switch (type)
            {
                case DataBaseType.SQLITE:
                    string SqliteDllPath = System.IO.Path.Combine(Environment.CurrentDirectory, "System.Data.SQLite.dll");
                    IDbConnection dbCnn = SqliteDllPath.GetInstance("System.Data.SQLite.SQLiteConnection", true) as IDbConnection;
                    return dbCnn;

                case DataBaseType.MSSQL:
                    return new System.Data.SqlClient.SqlConnection(connectStr);

                case DataBaseType.MSACCESS:
                    return new System.Data.OleDb.OleDbConnection(connectStr);

                case DataBaseType.ORACLE:
                    return new System.Data.OracleClient.OracleConnection(connectStr);
            }
            return null;
        }

        /// <summary>
        /// 创建数据库数据适配器，默认Oracle
        /// </summary>
        /// <param name="type">MSSQL|Oracle|MSAccess</param>
        /// <returns></returns>
        public static IDbDataAdapter CreateDBDataAdapter(this DataBaseType type)
        {
            switch (type)
            {
                case DataBaseType.SQLITE:
                    string SqliteDllPath = System.IO.Path.Combine(Environment.CurrentDirectory, "System.Data.SQLite.dll");
                    IDbDataAdapter dbAdpt = SqliteDllPath.GetInstance("System.Data.SQLite.SQLiteDataAdapter", true) as IDbDataAdapter;
                    return dbAdpt;

                case DataBaseType.MSSQL:
                    return new System.Data.SqlClient.SqlDataAdapter();

                case DataBaseType.MSACCESS:
                    return new System.Data.OleDb.OleDbDataAdapter();

                case DataBaseType.ORACLE:
                    return new System.Data.OracleClient.OracleDataAdapter();

                default:
                    return null;
            }
        }

        /// <summary>
        /// 根据数据库类型获取数据参数
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="value"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IDataParameter CreateDBDataParameter(this DataBaseType type, string paramName, object value)
        {
            switch (type)
            {
                case DataBaseType.SQLITE:
                    string SqliteDllPath = System.IO.Path.Combine(Environment.CurrentDirectory, "System.Data.SQLite.dll");
                    IDataParameter dbPara = SqliteDllPath.GetInstance("System.Data.SQLite.SQLiteParameter", true) as IDataParameter;
                    return dbPara;

                case DataBaseType.MSSQL:
                    return new System.Data.SqlClient.SqlParameter(paramName, value);

                case DataBaseType.MSACCESS:
                    return new System.Data.OleDb.OleDbParameter(paramName, value);

                case DataBaseType.ORACLE:
                    return new System.Data.OracleClient.OracleParameter(paramName, value);

                default:
                    return null;
            }
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

                case DataBaseType.MSACCESS:
                case DataBaseType.MSSQL:
                    return '@';

                default:
                    return char.MinValue;
            }
        }
        #endregion

        #region IDbConnection
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
        /// 获得表字段及其属性
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
                case DataBaseType.MSSQL:
                    sql = "SELECT [sc].[name] AS FIELDNAME, "
                        + "[st].[name] AS DATATYPE, "
                        + "[sc].[length] AS DATALENGTH "
                        + "FROM [syscolumns] AS [sc] "
                        + "INNER JOIN [sysobjects] AS [so] ON [so].[id] = [sc].[id] "
                        + "INNER JOIN [systypes] AS [st] ON [sc].[xtype] = [st].[xtype] "
                        + "WHERE [so].[name] = '{0}'";
                    break;
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
            dbConn.Close();
            return result;
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

            int count = cmd.ExecuteNonQuery();

            bool result = count > 0 ? true : false;

            if (!notClose) dbConn.Close();

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

        public static string GetSelectSQL(this DataBaseType dbType, string tableName, string whereClause, string selectedFields = "*")
        {
            string result = "select {0} from {1} where {2}";
            return string.Format(result, selectedFields, tableName, whereClause);
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

        /// <summary>
        /// 转换为字节数组
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this Image image)
        {
            MemoryStream ms = new MemoryStream();
            image.Save(ms, ImageFormat.Bmp);
            byte[] bs = ms.ToArray();
            ms.Close();
            return bs;
        }

        /// <summary>
        /// 转换为Image类型
        /// </summary>
        /// <param name="bs"></param>
        /// <returns></returns>
        public static Image ToImage(this byte[] bs)
        {
            MemoryStream ms = new MemoryStream(bs);
            Bitmap bmp = new Bitmap(ms);
            ms.Close();
            return bmp;
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
            a = 6378137.0,
            b = 6356752.314245179,
            e,
            r,
            l,
            m;
        private static long
            rate = 1000000000;

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
            Assembly assemble = Assembly.LoadFile(dllPath);
            return assemble.GetTypes();
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