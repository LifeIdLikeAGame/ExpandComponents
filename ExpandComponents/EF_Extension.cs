using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Web;

namespace ExpandComponents
{
    /// <summary>
    /// EF查询延伸类     基于EntityFramework
    /// </summary>
    public static class EF_Extension
    {
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public static string SqlConnectionStr { get; set; }

        /// <summary>
        /// 设置指定字段的排序查询数据
        /// </summary>
        /// <param name="query">当前Linq语句</param>
        /// <param name="sortFieldName">需要排序的字段名</param>
        /// <param name="order">排序规则(ASC|DESC)，默认：ASC升序</param>
        /// <returns></returns>
        public static IQueryable<T> SetQueryableOrder<T>(this IQueryable<T> query, string sortFieldName, string order = "ASC")
        {
            if (string.IsNullOrEmpty(sortFieldName)) throw new Exception("必须指定排序字段!");
            //根据属性名获取属性
            PropertyInfo sortProperty = typeof(T).GetProperty(sortFieldName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (sortProperty == null) throw new Exception("查询对象中不存在排序字段" + sortFieldName + "！");
            //创建表达式变量参数
            ParameterExpression param = Expression.Parameter(typeof(T), "t");
            Expression body = param;
            if (Nullable.GetUnderlyingType(body.Type) != null)
                body = Expression.Property(body, "Value");
            //创建一个访问属性的表达式
            body = Expression.MakeMemberAccess(body, sortProperty);
            LambdaExpression keySelectorLambda = Expression.Lambda(body, param);
            string queryMethod = order.ToUpper() == "DESC" ? "OrderByDescending" : "OrderBy";
            query = query.Provider.CreateQuery<T>(Expression.Call(typeof(Queryable), queryMethod,
                                                               new Type[] { typeof(T), body.Type },
                                                               query.Expression,
                                                               Expression.Quote(keySelectorLambda)));
            return query;
        }
        /// <summary>
        /// 设置多个指定字段的排序查询数据
        /// </summary>
        /// <param name="sortFieldString">需要排序的字段,字段间以逗号隔开(例：“FieldA,FieldB”)</param>
        /// <param name="order">排序规则(ASC|DESC)，默认：ASC升序</param>
        /// <returns></returns>
        public static IQueryable<T> SetQueryableOrderArray<T>(this IQueryable<T> query, string sortFieldString, string order = "ASC")
        {
            var sortFieldArray = sortFieldString.Split(',');
            return query.SetQueryableOrderArray(sortFieldArray, order);
        }
        /// <summary>
        /// 设置多个指定字段的排序查询数据
        /// </summary>
        /// <param name="sortFieldArray">排序字段数组</param>
        /// <param name="order">排序规则(ASC|DESC)，默认：ASC升序</param>
        /// <returns></returns>
        public static IQueryable<T> SetQueryableOrderArray<T>(this IQueryable<T> query, string[] sortFieldArray, string order = "ASC")
        {

            //创建表达式变量参数
            var parameter = Expression.Parameter(typeof(T), "t");

            if (sortFieldArray.Length == 0) throw new Exception("必须指定排序字段!");
            if (sortFieldArray != null && sortFieldArray.Length > 0)
            {
                for (int i = 0; i < sortFieldArray.Length; i++)
                {
                    //根据属性名获取属性
                    var property = typeof(T).GetProperty(sortFieldArray[i], BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    if (property == null) throw new Exception("查询对象中不存在排序字段" + sortFieldArray[i] + "！");
                    //创建一个访问属性的表达式
                    var propertyAccess = Expression.MakeMemberAccess(parameter, property);
                    var orderByExp = Expression.Lambda(propertyAccess, parameter);


                    string OrderName = order.ToUpper() == "DESC" ? "OrderByDescending" : "OrderBy";


                    MethodCallExpression resultExp = Expression.Call(typeof(Queryable), OrderName, new Type[] { typeof(T), property.PropertyType }, query.Expression, Expression.Quote(orderByExp));
                    query = query.Provider.CreateQuery<T>(resultExp);
                }
            }
            return query;
        }

        /// <summary>
        /// 批量插入方法
        /// (调用该方法需要注意，字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">数据集</param>
        /// <param name="tableName">数据库表名</param>
        /// <param name="connectstring">连接字符串</param>
        public static int BulkInsert<T>(this IList<T> list,string connectstring = "",string tableName="") where T : class
        {
            if (string.IsNullOrWhiteSpace(connectstring)) connectstring = SqlConnectionStr;
            if (string.IsNullOrWhiteSpace(connectstring)) throw new Exception("连接字符串为空");
            var _tableName = typeof(T).Name;
            SqlConnection conn = new SqlConnection(connectstring);
            using (var bulkCopy = new SqlBulkCopy(conn))
            {
                bulkCopy.BatchSize = list.Count;
                bulkCopy.DestinationTableName = _tableName;

                var table = new DataTable();
                var props = TypeDescriptor.GetProperties(typeof(T), new Attribute[] { })
                    //Dirty hack to make sure we only have system data types 
                    //i.e. filter out the relationships/collections
                    .Cast<PropertyDescriptor>()
                    .Where(propertyInfo => propertyInfo.PropertyType.Namespace.Equals("System"))
                    .ToArray();
                var props2 = TypeDescriptor.GetProperties(typeof(T), new Attribute[] { new NotMappedAttribute() })
                       .Cast<PropertyDescriptor>()
                       .Where(propertyInfo => propertyInfo.PropertyType.Namespace.Equals("System"))
                       .ToArray();
                var arr = props2.Select(a => a.Name).ToArray();
                foreach (var propertyInfo in props)
                {
                    bulkCopy.ColumnMappings.Add(propertyInfo.Name, propertyInfo.Name);
                    table.Columns.Add(propertyInfo.Name, Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType);
                }

                var values = new object[props.Length];
                foreach (var item in list)
                {
                    var columns = new List<object>();
                    for (var i = 0; i < values.Length; i++)
                    {
                        var name = props[i].Name;
                        if (arr.Contains(name)) continue;
                        columns.Add(props[i].GetValue(item));
                    }
                    var colArr = columns.ToArray();
                    table.Rows.Add(colArr);
                }

                bulkCopy.WriteToServer(table);
                return list.Count;
            }
        }

        /// <summary>
        /// 海量数据插入方法
        /// (调用该方法需要注意，DataTable中的字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <param name="connectstring">数据连接字符串</param>
        /// <param name="table">内存表数据</param>
        /// <param name="tableName">目标数据库表的名称</param>
        public static int BulkInsert(this DataTable table, string connectstring, string tableName)
        {
            if (table != null && table.Rows.Count > 0)
            {
                if (string.IsNullOrWhiteSpace(connectstring)) connectstring = SqlConnectionStr;
                if (string.IsNullOrWhiteSpace(connectstring)) throw new Exception("连接字符串为空");
                using (SqlBulkCopy bulk = new SqlBulkCopy(connectstring))
                {
                    bulk.BatchSize = 1000;
                    bulk.BulkCopyTimeout = 100;
                    bulk.DestinationTableName = tableName;
                    bulk.WriteToServer(table);
                }
                return table.Rows.Count;
            }
            return 0;
        }

        /// <summary>
        /// 海量数据插入方法
        /// (调用该方法需要注意，<see cref="T"/>中的字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <param name="connectstring">数据连接字符串</param>
        /// <param name="list">数据集</param>
        public static int BulkInserts<T>(this List<T> list, string connectstring="")
        {
            if (string.IsNullOrWhiteSpace(connectstring)) connectstring = SqlConnectionStr;
            if (string.IsNullOrWhiteSpace(connectstring)) throw new Exception("连接字符串为空");
            var tableName = typeof(T).Name;
            var dt = list.ToDataTable();
            return BulkInsert(dt, connectstring, tableName);
        }


        /// <summary>
        /// Convert a List{T} to a DataTable.
        /// </summary>
        private static DataTable ToDataTable<T>(this List<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }

        /// <summary>
        /// Determine of specified type is nullable
        /// </summary>
        private static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        /// <summary>
        /// Return underlying type if type is Nullable otherwise return the type
        /// </summary>
        private static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }
       

    }
}