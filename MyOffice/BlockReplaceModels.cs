using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyOffice
{
    /// <summary>
    /// 块替换模型类 - 适用于单个块替换配置
    /// 用于定义块定义替换操作中使用的数据模型
    /// </summary>
    public class BlockReplaceModel
    {
        /// <summary>
        /// 原块名称
        /// </summary>
        public string OriginalBlockName { get; set; }

        /// <summary>
        /// 新块名称
        /// </summary>
        public string NewBlockName { get; set; }

        /// <summary>
        /// 是否保留原有属性
        /// </summary>
        public bool PreserveAttributes { get; set; } = true;

        /// <summary>
        /// 是否保留位置
        /// </summary>
        public bool PreservePosition { get; set; } = true;

        /// <summary>
        /// 是否保留旋转
        /// </summary>
        public bool PreserveRotation { get; set; } = true;

        /// <summary>
        /// 是否保留缩放
        /// </summary>
        public bool PreserveScale { get; set; } = true;

        /// <summary>
        /// 是否创建备份
        /// </summary>
        public bool CreateBackup { get; set; } = true;
    }

    /// <summary>
    /// 块替换结果模型 - 通用结果类
    /// </summary>
    public class BlockReplaceResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 替换数量
        /// </summary>
        public int ReplaceCount { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 处理的文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 处理时间
        /// </summary>
        public DateTime ProcessTime { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// 批处理配置模型 - 兼容性保留
    /// </summary>
    public class BatchReplaceConfig
    {
        /// <summary>
        /// 源文件路径列表
        /// </summary>
        public List<string> SourceFiles { get; set; } = new List<string>();

        /// <summary>
        /// 输出目录
        /// </summary>
        public string OutputDirectory { get; set; }

        /// <summary>
        /// 替换配置列表
        /// </summary>
        public List<BlockReplaceModel> ReplaceConfigs { get; set; } = new List<BlockReplaceModel>();

        /// <summary>
        /// 是否备份原文件
        /// </summary>
        public bool CreateBackup { get; set; } = true;

        /// <summary>
        /// 是否覆盖输出文件
        /// </summary>
        public bool OverwriteOutput { get; set; } = false;
    }
}