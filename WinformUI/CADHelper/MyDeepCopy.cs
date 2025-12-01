using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml.Serialization;


namespace DotNetARX
{
    /// <summary>
    /// 对象的深度拷贝（序列化的方式）
    /// </summary>
    public static class MyDeepCopy
    {

        //方法1.对象实现ICloneable 接口 对象.MemberWiseClone(); 速度很慢
        //方法2 反射进行克隆 反射很慢 通过缓存方式提高速度 较优解
        //方法3 序列化
        //方法4 表达式目录树 最优解
        public class Reflection<TIn, TOut>
        {
            private static TOut tOut;

            public static TOut TransReflection<TIn, TOut>(TIn tIn)
            {
                TOut tOut = Activator.CreateInstance<TOut>();
                if (tOut != null) return tOut;
                var tInType = tIn.GetType();
                foreach (var item in tOut.GetType().GetProperties())
                {
                    var itemIn = tInType.GetProperty(item.Name);
                    if (itemIn != null)
                    {
                        item.SetValue(tOut, itemIn.GetValue(tIn, null), null);
                    }
                }
                return tOut;
            }
        }

        public class ExpressionHelper<TIn, TOut>
        {
            private static readonly Func<TIn, TOut> cache = GetFunc();

            private static Func<TIn, TOut> GetFunc()
            {
                ParameterExpression parameterExpression = Expression.Parameter(typeof(TIn), "p");//传进来的tIn作为参数p来使用，p代表tIn。
                List<MemberBinding> memberBindings = new List<MemberBinding>();
                foreach (var item in typeof(TOut).GetProperties())
                {
                    if (!item.CanWrite) continue; ;
                    MemberExpression property = Expression.Property(parameterExpression, typeof(TIn).GetProperty(item.Name));
                    MemberBinding memberBinding = Expression.Bind(item, property);
                    memberBindings.Add(memberBinding);
                }
                MemberInitExpression memberInitExpression = Expression.MemberInit(Expression.New(typeof(TOut)), memberBindings.ToArray());
                var lamada = Expression.Lambda<Func<TIn, TOut>>(memberInitExpression, new ParameterExpression[] {parameterExpression});
                return lamada.Compile();
            }

            public static TOut Trans(TIn tIn)
            {
                return cache(tIn);
            }

            public static List<TOut> Trans(List<TIn> tIns)
            {
                return tIns.Select(Trans).ToList();
            }
        }

        /// <summary>
        /// xml序列化的方式实现深拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <returns></returns>
        public static T XmlDeepCopy<T>(T t)
        {
            //创建Xml序列化对象
            XmlSerializer xml = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream())//创建内存流
            {
                //将对象序列化到内存中
                xml.Serialize(ms, t);
                ms.Position = default;//将内存流的位置设为0
                return (T)xml.Deserialize(ms);//继续反序列化
            }
        }

        /// <summary>
        /// 二进制序列化的方式进行深拷贝
        /// 确保需要拷贝的类里的所有成员已经标记为 [Serializable] 如果没有加该特性特报错
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <returns></returns>
        public static T BinaryDeepCopy<T>(T t)
        {
            //创建二进制序列化对象
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream())//创建内存流
            {
                //将对象序列化到内存中
                bf.Serialize(ms, t);
                ms.Position = default;//将内存流的位置设为0
                return (T)bf.Deserialize(ms);//继续反序列化
            }
        }
    }
}
