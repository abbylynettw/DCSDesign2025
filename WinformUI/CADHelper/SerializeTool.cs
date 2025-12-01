using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace WinformUI.CADHelper
{
    public static class SerializeTools
    {
        /// <summary>
        /// 序列化
        /// </summary>
        /// <param name="obj"></param>  
        /// <returns></returns>
        public static string Serialize(object obj)
        {
            MemoryStream ms = new MemoryStream();
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(ms, obj);
            return Convert.ToBase64String(ms.ToArray());
        }
        /// <summary>
        /// 反序列化
        /// </summary>
        /// <param name="modelBase64"></param>
        /// <returns></returns>
        public static T DeSerialize<T>(string modelBase64) where T : class, new()
        {
            //  System.Reflection.Assembly[] ass = AppDomain.CurrentDomain.GetAssemblies();
            byte[] bytes = Convert.FromBase64String(modelBase64);
            MemoryStream ms = new MemoryStream(bytes);
            BinaryFormatter bf = new BinaryFormatter();
            bf.Binder = new UBinder();
            return bf.Deserialize(ms) as T;
        }


        /// <summary>
        /// 反序列化
        /// </summary>
        /// <param name="modelBase64"></param>
        /// <returns></returns>
        public static object DeSerializeObj(string modelBase64)
        {
            //  System.Reflection.Assembly[] ass = AppDomain.CurrentDomain.GetAssemblies();
            byte[] bytes = Convert.FromBase64String(modelBase64);
            MemoryStream ms = new MemoryStream(bytes);
            BinaryFormatter bf = new BinaryFormatter();
            bf.Binder = new UBinder();
            return bf.Deserialize(ms);
        }

        /// <summary>
        /// 克隆对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <returns></returns>
        public static T CloneObj<T>(T t) where T : class, new()
        {
            try
            {
                return DeSerialize<T>(Serialize(t));
            }
            catch (Exception ex)
            {               
                throw;
            }

        }
        public class UBinder : SerializationBinder
        {
            public override Type BindToType(string assemblyName, string typeName)
            {
                switch (typeName)
                {
                    //case "Model.Assembly.AssemblyTemplate":
                    //    return typeof(Model.Assembly.AssemblyTemplate);
                    //case "Model.Assembly.AssemblyTemplateItem":
                    //    return typeof(Model.Assembly.AssemblyTemplateItem);
                    //case "Model.Module.ModulePropertyView":
                    //    return typeof(Model.Module.ModulePropertyView);
                    //case "Model.Module.ModulePropertyTemplate":
                    //    return typeof(Model.Module.ModulePropertyTemplate);
                    //case "Model.Legend.LegendEntity":
                    //    return typeof(Model.Legend.LegendEntity);
                    //case "Model.Module.ModuleProperty":
                    //    return typeof(Model.Module.ModuleProperty);
                    default:
                        return Type.GetType(typeName);
                }

            }
        }
    }
}
