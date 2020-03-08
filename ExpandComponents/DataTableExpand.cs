using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExpandComponents
{
    /// <summary>
    /// 数据集拓展
    /// </summary>
   public static class  DataExpand
    {
        /// <summary>
        ///   将DataTable转实体
        /// </summary>
        /// <typeparam name="T">泛型类型参数。</typeparam>
        /// <param name="this">当前DataTable</param>
        /// <returns>作为一个IEnumerable &lt;T &gt;</returns>
        public static IEnumerable<T> ToEntities<T>(this DataTable @this) where T : new()
        {
            Type type = typeof(T);
            PropertyInfo[] properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            FieldInfo[] fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance);

            var list = new List<T>();

            foreach (DataRow dr in @this.Rows)
            {
                var entity = new T();

                foreach (PropertyInfo property in properties)
                {
                    if (@this.Columns.Contains(property.Name))
                    {
                        Type valueType = property.PropertyType;
                        property.SetValue(entity, dr[property.Name].To(valueType), null);
                    }
                }

                foreach (FieldInfo field in fields)
                {
                    if (@this.Columns.Contains(field.Name))
                    {
                        Type valueType = field.FieldType;
                        field.SetValue(entity, dr[field.Name].To(valueType));
                    }
                }

                list.Add(entity);
            }

            return list;
        }
        

        /// <summary>
        /// 类型转换
        /// </summary>
        /// 例子
        /// var result1 = value.To int ();/ /返回1;
        /// var result2 = value.To int? ();/ /返回1;
        /// var result3 = null . to int? ();/ /返回null;
        /// var result t4 = dbNullValue.To int？ ();/ /返回null;
        /// 泛型类型参数
        /// <param name="@this">当前</param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static object To(this Object @this, Type type)
        {
            if (@this != null)
            {
                Type targetType = type;

                if (@this.GetType() == targetType)
                {
                    return @this;
                }

                TypeConverter converter = TypeDescriptor.GetConverter(@this);
                if (converter != null)
                {
                    if (converter.CanConvertTo(targetType))
                    {
                        return converter.ConvertTo(@this, targetType);
                    }
                }

                converter = TypeDescriptor.GetConverter(targetType);
                if (converter != null)
                {
                    if (converter.CanConvertFrom(@this.GetType()))
                    {
                        return converter.ConvertFrom(@this);
                    }
                }

                if (@this == DBNull.Value)
                {
                    return null;
                }
            }

            return @this;
        }
    }
}
