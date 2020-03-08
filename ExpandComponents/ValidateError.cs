using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace ExpandComponents
{
    /********************************************************************************

   ** 类名称： ValidateError

   ** 描述：数据验证错误友好提示

   ** 引用： System.Web.dll、System.Web.Mvc.dll

   ** 作者： LW

   *********************************************************************************/
    /// <summary>
    /// 数据验证拓展,错误友好提示
    /// </summary>
    public static class ValidateError
    {

        /// <summary>
        /// 是否包含某列名,如果包含则校验通过，如果不包含则抛出<see cref="Exception"/>异常，异常可自定义
        /// </summary>
        /// <param name="ColumnName">列名</param>
        /// <param name="dt">一个<see cref="DataTable"/>数据表</param>
        /// <param name="ErrorMsg">必须包含索引{0}，{0}为不含列名时文字提示所在位置。默认:数据表中不存在:{0}列名</param>
        public static CheckResult CheckExistColumns(this DataTable dt, string ColumnName, string ErrorMsg = "")
        {
            var check = new CheckResult();
            var _title = "";
            if (string.IsNullOrEmpty(ColumnName)) return check;
            if (dt == null || !dt.Columns.Contains(ColumnName))
            {
                _title += $"【{ColumnName}】";
            }
            if (_title != "")
            {
                var _msg = string.Format("数据表中不存在以下列名:{0}", _title);
                if (!string.IsNullOrEmpty(ErrorMsg)) _msg = string.Format(ErrorMsg, _title);
                check.IsPass = false;
                check.ErrorMsg = _msg;
                return check;
            }
            return check;
        }
        /// <summary>
        /// 是否包含某列名,提示可自定义
        /// </summary>
        /// <param name="ColumnName">列名数组</param>
        /// <param name="dt">一个<see cref="DataTable"/>数据表</param>
        /// <param name="ErrorMsg">必须包含索引{0}，{0}为不含列名时文字提示所在位置。默认:数据表中不存在:{0}列名</param>
        public static CheckResult CheckExistColumns(this DataTable dt, string[] ColumnArr, string ErrorMsg = "")
        {
            var check = new CheckResult();
            var _title = "";
            if (ColumnArr == null) return check;
            foreach (var item in ColumnArr)
            {
                if (dt == null || !dt.Columns.Contains(item))
                {
                    _title += $"【{item}】";
                }
            }
            if (_title != "")
            {
                var _msg = string.Format("数据表中不存在以下列名:{0}", _title);
                if (!string.IsNullOrEmpty(ErrorMsg)) _msg = string.Format(ErrorMsg, _title);
                check.IsPass = false;
                check.ErrorMsg = _msg;
                return check;
            }
            return check;
        }
        /// <summary>
        /// 字符分割验证是否包含某列名，异常可自定义
        /// </summary>
        /// <param name="ColumnNameList">列名数组</param>
        /// <param name="dt">一个<see cref="DataTable"/>数据表</param>
        /// <param name="ErrorMsg">必须包含索引{0}，{0}为不含列名时文字提示所在位置。默认:数据表中不存在:{0}列名</param>
        public static CheckResult CheckSplitExistColumns(this DataTable dt, string ColumnNameList, string ErrorMsg = "", char split = ',')
        {
            var check = new CheckResult();
            var _title = "";
            if (string.IsNullOrEmpty(ColumnNameList)) return check;
            foreach (var item in ColumnNameList.Split(split))
            {
                if (dt == null || !dt.Columns.Contains(item))
                {
                    _title += $"【{item}】";
                }
            }
            if (_title != "")
            {
                var _msg = string.Format("数据表中不存在以下列名:{0}", _title);
                if (!string.IsNullOrEmpty(ErrorMsg)) _msg = string.Format(ErrorMsg, _title);
                check.IsPass = false;
                check.ErrorMsg = _msg;
                return check;
            }
            return check;
        }


        /// <summary>
        /// 字符分割验证非空列是否为空,提示可自定义
        /// </summary>
        /// <param name="ColumnNameList">列名数组</param>
        /// <param name="dt">一个<see cref="DataTable"/>数据表</param>
        /// <param name="ErrorMsg">必须包含索引{0}，{0}为不含列名时文字提示所在位置。默认:列名:{0}数据不能为空</param>
        public static void CheckRowsSplitIsNull(this DataRow dr, string ColumnNameList, string ErrorMsg = "", char split = ',')
        {
            var _title = "";
            if (string.IsNullOrEmpty(ColumnNameList)) return;
            foreach (var item in ColumnNameList.Split(split))
            {
                if (dr == null || dr[item] == null || string.IsNullOrEmpty(dr[item].ToString()))
                {
                    _title += $"【{item}】";
                }
            }
            if (_title != "")
            {
                var _msg = string.Format("列名:{0}数据不能为空", _title);
                if (!string.IsNullOrEmpty(ErrorMsg)) _msg = string.Format(ErrorMsg, _title);
                var _err = new Exception(_msg);
                throw _err;
            }
        }
        /// <summary>
        /// 字符分割验证非空列是否为空,提示可自定义
        /// </summary>
        /// <param name="ColumnNameList">列名数组</param>
        /// <param name="dt">一个<see cref="DataTable"/>数据表</param>
        /// <param name="ErrorMsg">必须包含索引{0}，{0}为不含列名时文字提示所在位置。默认:列名:{0}数据不能为空</param>
        public static void CheckRowsIsNull(this DataRow dr, string[] ColumnArr, string ErrorMsg = "")
        {
            var _title = "";
            if (ColumnArr == null || ColumnArr.Length == 0) return;
            foreach (var item in ColumnArr)
            {
                if (dr == null || dr[item] == null || string.IsNullOrEmpty(dr[item].ToString()))
                {
                    _title += $"【{item}】";
                }
            }
            if (_title != "")
            {
                var _msg = string.Format("列名:{0}数据不能为空", _title);
                if (!string.IsNullOrEmpty(ErrorMsg)) _msg = string.Format(ErrorMsg, _title);
                var _err = new Exception(_msg);
                throw _err;
            }
        }
        /// <summary>
        /// 验证非空列是否为空,提示可自定义
        /// </summary>
        /// <param name="ColumnNameList">列名数组</param>
        /// <param name="dt">一个<see cref="DataTable"/>数据表</param>
        /// <param name="ErrorMsg">必须包含索引{0}，{0}为不含列名时文字提示所在位置。默认:列名:{0}数据不能为空</param>
        public static void CheckRowsIsNull(this DataRow dr, string ColumnName, string ErrorMsg = "")
        {
            var _title = "";
            if (string.IsNullOrEmpty(ColumnName)) return;
            if (dr == null || dr[ColumnName] == null || string.IsNullOrEmpty(dr[ColumnName].ToString()))
            {
                _title += $"【{ColumnName}】";
            }
            if (_title != "")
            {
                var _msg = string.Format("列名:{0}数据不能为空", _title);
                if (!string.IsNullOrEmpty(ErrorMsg)) _msg = string.Format(ErrorMsg, _title);
                var _err = new Exception(_msg);
                throw _err;
            }
        }
        
        /// <summary>
        /// 将数据转为指定数据类型
        /// </summary>
        /// <typeparam name="T"><see cref="struct"/>常见类型</typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static T? ParseValue<T>(this object value) where T : struct
        {
            T? result = null;
            if (value != null && !(value is DBNull))
            {
                try
                {
                    result = Convert.ChangeType(value, typeof(T)) as T?;
                }
                catch
                {
                    var _err = new Exception($"无法将“{value}”转为{typeof(T)}类型");
                    throw _err;
                }

            }
            return result;
        }
    }
    /// <summary>
    /// 验证结果
    /// </summary>
    public class CheckResult
    {
        /// <summary>
        /// 是否通过校验
        /// </summary>
        public bool IsPass { get; set; } = true;
        /// <summary>
        /// 校验不通过是错误提示
        /// </summary>
        public string ErrorMsg { get; set; }
    }
}