using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace ExpandComponents
{
    /// <summary>
    /// Cookie配置
    /// </summary>
    public class CookieConfigure
    {
        /// <summary>
        /// 该Cookie是否已存在
        /// </summary>
        /// <param name="_name">Cookie名</param>
        /// <param name="Value">返回的Cookie值</param>
        /// <returns></returns>
        public static bool IsExist(string _name, out string Value)
        {
            Value = GetCookie(_name);
            return !string.IsNullOrWhiteSpace(Value);
        }

        /// <summary>
        /// 清除所有Cookie
        /// </summary>
        public static void ClearAll()
        {
            HttpContext.Current.Response.Cookies.Clear();
            int count = HttpContext.Current.Request.Cookies.Count;
            for (int i = 0; i < count; i++)
            {
                string name = HttpContext.Current.Request.Cookies[i].Name;
                HttpCookie httpCookie = new HttpCookie(name);
                httpCookie.Expires = DateTime.Now.AddDays(-10.0);
                HttpContext.Current.Response.Cookies.Add(httpCookie);
            }
        }

        /// <summary>
        /// 清除单个Cookie
        /// </summary>
        /// <param name="_name">需要清除的Cookie名</param>
        public static void Clear(string _name)
        {
            HttpContext.Current.Response.Cookies.Clear();
            HttpCookie httpCookie = new HttpCookie(_name);
            httpCookie.Expires = DateTime.Now.AddDays(-10.0);
            HttpContext.Current.Response.Cookies.Add(httpCookie);
        }

        /// <summary>
        /// 使用的编码
        /// </summary>
        private static readonly Encoding enc = Encoding.UTF8;

        /// <summary>
        /// 设置Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="HttpOnly">获取或设置一个值，该值指定 Cookie 是否可通过客户端脚本访问。如果 Cookie 具有 HttpOnly 属性且不能通过客户端脚本访问，则为 true；否则为 false。默认值为 false。</param>
        public static void SetCookie(string _name, string _value, bool HttpOnly)
        {
            if (HttpContext.Current.Request.Cookies[_name] == null)
            {
                HttpContext.Current.Request.Cookies.Add(new HttpCookie(_name));
            }
            HttpContext.Current.Request.Cookies[_name].Value = HttpUtility.UrlEncode(_value, enc);
            HttpContext.Current.Request.Cookies[_name].HttpOnly = HttpOnly;
            if (HttpContext.Current.Response.Cookies[_name] == null)
            {
                HttpContext.Current.Response.Cookies.Add(new HttpCookie(_name));
            }
            HttpContext.Current.Response.Cookies[_name].Value = HttpUtility.UrlEncode(_value, enc);
            HttpContext.Current.Response.Cookies[_name].HttpOnly = HttpOnly;
        }

        /// <summary>
        /// 设置Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_ts">过期时间</param>
        /// <param name="HttpOnly">获取或设置一个值，该值指定 Cookie 是否可通过客户端脚本访问。如果 Cookie 具有 HttpOnly 属性且不能通过客户端脚本访问，则为 true；否则为 false。默认值为 false。</param>
        public static void SetCookie(string _name, string _value,TimeSpan _ts, bool HttpOnly)
        {
            if (HttpContext.Current.Request.Cookies[_name] == null)
            {
                HttpContext.Current.Request.Cookies.Add(new HttpCookie(_name));
            }
            HttpContext.Current.Request.Cookies[_name].Value = HttpUtility.UrlEncode(_value, enc);
            HttpContext.Current.Request.Cookies[_name].HttpOnly = HttpOnly;
            if(_ts!=null) HttpContext.Current.Request.Cookies[_name].Expires = DateTime.Now.Add(_ts);
            if (HttpContext.Current.Response.Cookies[_name] == null)
            {
                HttpContext.Current.Response.Cookies.Add(new HttpCookie(_name));
            }
            HttpContext.Current.Response.Cookies[_name].Value = HttpUtility.UrlEncode(_value, enc);
            HttpContext.Current.Response.Cookies[_name].HttpOnly = HttpOnly;
        }

        /// <summary>
        /// 设置客户端允许访问的Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        public static void SetCookie(string _name, string _value)
        {
            SetCookie(_name, _value, false);
        }

        /// <summary>
        /// 设置客户端允许访问的Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_ts">过期时间</param>
        public static void SetCookie(string _name, string _value,TimeSpan _ts)
        {
            SetCookie(_name, _value, _ts, false);
        }
        /// <summary>
        /// 设置客户端允许访问的Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_hours">过期小时数</param>
        public static void SetCookie(string _name, string _value, int _hours)
        {
            SetCookie(_name, _value, new TimeSpan(_hours,0,0), false);
        }

        /// <summary>
        /// 设置客户端允许访问的Cookie,设置分钟过期
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_minutes">过期分钟数</param>
        public static void SetCookieMinutes(string _name, string _value, int _minutes)
        {
            SetCookie(_name, _value, new TimeSpan(0, _minutes, 0), false);
        }

        /// <summary>
        ///设置客户端不允许访问的Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        public static void SetCookieHttpOnly(string _name, string _value)
        {
            SetCookie(_name, _value, true);
        }

        /// <summary>
        ///设置客户端不允许访问的Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_ts">过期时间</param>
        public static void SetCookieHttpOnly(string _name, string _value, TimeSpan _ts)
        {
            SetCookie(_name, _value,_ts, true);
        }
        /// <summary>
        /// 设置客户端允许访问的Cookie,设置分钟过期
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_minutes">过期分钟数</param>
        public static void SetCookieHttpOnlyMinutes(string _name, string _value, int _minutes)
        {
            SetCookie(_name, _value, new TimeSpan(0,_minutes, 0), true);
        }

        /// <summary>
        /// 设置客户端允许访问的Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <param name="_value">值</param>
        /// <param name="_hours">过期小时数</param>
        public static void SetCookieHttpOnly(string _name, string _value, int _hours)
        {
            SetCookie(_name, _value, new TimeSpan(_hours, 0, 0), true);
        }

        /// <summary>
        /// 获取Cookie
        /// </summary>
        /// <param name="_name">键</param>
        /// <returns>值</returns>
        public static string GetCookie(string _name)
        {
            var temp = HttpContext.Current.Request.Cookies[_name];
            return temp != null ? HttpUtility.UrlDecode(temp.Value, enc) : null;
        }
    }
}