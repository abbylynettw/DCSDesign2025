
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyOffice
{
    public enum RuleType
    {
        相等,
        包含
    }
    // 替换规则类
    public class ReplaceRule
    {
        public RuleType Type { get; set; }
        public string FindText { get; set; }
        public string ReplaceText { get; set; }
    }

    // 处理选项类
    public class ProcessingOptions
    {
        // DWG选项
        public bool ProcessDwgText { get; set; }
        public bool ProcessDwgTableContent { get; set; }
        public bool ProcessDwgBlockAttributes { get; set; }

        public bool ProcessDowngradeA { get; set; } //降低成A版本

        // Word选项
        public bool 项目名称修改 { get; set; }
        public string 项目名称值 { get; set; }

        public bool 项目编号修改 { get; set; }
        public string 项目编号值 { get; set; }


        public bool 内部编码修改 { get; set; }
        public string 内部编码值 { get; set; }
        public bool 页眉内部编码修改 { get; set; }

        public bool 页眉版本状态修改 { get; set; }
        public string 页眉版本状态值 { get; set; }

        public bool 页眉总页码修改 { get; set; }
        public string 页眉总页码值 { get; set; }

        public bool 外部编码修改 { get; set; }
        public bool 标题替换 { get; set; }
        public bool 页眉标题替换 { get; set; }
        public bool 修改记录重置 { get; set; }
        public bool 编校审批表格重置 { get; set; }
    

        // 文件名选项
        public bool ProcessFolderName { get; set; }
        public bool ProcessDwgFileName { get; set; }
        public bool ProcessWordFileName { get; set; }
      
    }


    public class OuterCodeAndTitle
    {
       public string 外部编码 { get; set; }
       public string 文件标题 { get; set; }
    }

    public class MyCell
    {
        public int rowIndex { get; set; }
        public int colIndex { get; set; }
        public string value { get; set; }
    }
}
