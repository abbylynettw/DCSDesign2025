using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace WinformUI.CADHelper
{
    /// <summary>
    /// 实体类反射
    /// </summary>
    public class Reflection
    {
        /// <summary>
        /// 反射得到实体类的字段名称和值
        /// var dict = GetProperties(model);
        /// </summary>
        /// <typeparam name="T">实体类</typeparam>
        /// <param name="t">实例化</param>
        /// <returns></returns>
        public static Dictionary<object, object> GetProperties<T>(T t)
        {
            var ret = new Dictionary<object, object>();
            if (t == null) { return null; }
            PropertyInfo[] properties = t.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
            if (properties.Length <= 0) { return null; }
            foreach (PropertyInfo item in properties)
            {

                if (item.PropertyType.IsValueType || item.PropertyType.Name.StartsWith("String"))
                {
                    string name = item.Name;
                    object value = item.GetValue(t, null);
                    ret.Add(name, value);
                }

            }
            return ret;
        }
  
        /// <summary>
        /// 反射获取指定名称的字段值
        /// </summary>
        /// <typeparam name="T">实体类</typeparam>
        /// <param name="t">实例化</param>
        /// <param name="Name">字段名</param>
        /// <returns>返回值</returns>
        public static object GetPropertieOne<T>(T t, string Name)
        {
            object Value = false;
            if (t == null) { return Value; }
            PropertyInfo[] properties = t.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
            if (properties.Length <= 0) { return Value; }
            foreach (PropertyInfo item in properties)
            {
                if (item.Name.ToUpper().Equals(Name.ToUpper()))
                {
                    Value = item.GetValue(t, null);
                }
            }
            return Value;
        }

    }
}
