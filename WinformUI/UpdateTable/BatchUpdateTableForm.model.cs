using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace WinformUI.UpdateTable
{
    public class Model
    {
        // 映射配置类
        [Serializable]
        public class MappingConfig
        {
            public bool Enabled { get; set; }
            public string DwgKeyword { get; set; }//dwg-PageType
            public string CADTableTitle { get; set; }//CAD表格名
            public string ExcelSheet { get; set; }//Excel表单
        }

        // 用于XML序列化的包装类
        [Serializable]
        [XmlRoot("MappingConfigurations")]
        public class MappingCollection
        {
            [XmlArray("Mappings")]
            [XmlArrayItem("Mapping")]
            public List<MappingConfig> Mappings { get; set; }
        }
    }
    public enum LogLevel
    {
        Info,
        Warning,
        Error,
        Success
    }
}
