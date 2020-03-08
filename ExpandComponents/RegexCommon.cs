using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace ExpandComponents
{
    /// <summary>
    /// 常用正则表达式验证
    /// </summary>
    public static class RegexCommon
    {
        private static Regex reg { get; set; }

        /// <summary>
        /// 判断是否是纯数字（例：123456）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static string RegIsNumber(this string str) {
            if (string.IsNullOrEmpty(str)) return "";
            reg = new Regex(@"^[0 - 9]*$");
            if (reg.IsMatch(str)) return str;
            else throw new Exception(str+"不是纯数字类型。");
        }
        /// <summary>
        /// 判断是否是日期（例：2000-01-01）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static string RegIsDate(this string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            reg = new Regex(@"^\d{4}(-|/|.)\d{1,2}(-|/|.)\d{1,2}$");
            if (reg.IsMatch(str)) return str;
            else throw new Exception(str + "不是日期类型。");
        }
        /// <summary>
        /// 判断是否是日期时间（例：2000-01-01 00:00:00）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static string RegIsDateTime(this string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            reg = new Regex(@"^\d{4}(-|/|.)\d{1,2}(-|/|.)\d{1,2}");
            if (reg.IsMatch(str)) return str;
            else throw new Exception(str + "不是日期类型。");
        }
        /// <summary>
        /// 判断是否是11位手机号码
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static string RegIsPhone(this string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            reg = new Regex(@"^(13[0-9]|14[5|7]|15[0|1|2|3|5|6|7|8|9]|18[0|1|2|3|5|6|7|8|9])\d{8}$");
            if (reg.IsMatch(str)) return str;
            else throw new Exception(str + "不是手机号码。");
        }
    }
}