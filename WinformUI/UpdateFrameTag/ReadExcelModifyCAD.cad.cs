using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System;
using WinformUI.CADHelper;
using System.Linq;
using Newtonsoft.Json.Linq;

/// <summary>
/// CAD图框标签更新处理类 - 优化版本
/// </summary>
public class CADFrameUpdater
{
    // 定义日志事件和事件参数类
    public class LogEventArgs : EventArgs
    {
        public string Message { get; set; }
        public bool IsError { get; set; }

        public LogEventArgs(string message, bool isError = false)
        {
            Message = message;
            IsError = isError;
        }
    }

    // 声明日志事件
    public event EventHandler<LogEventArgs> LogMessageGenerated;

    // 触发日志事件的方法
    protected virtual void OnLogMessage(string message, bool isError = false)
    {
        LogMessageGenerated?.Invoke(this, new LogEventArgs(message, isError));
    }

    /// <summary>
    /// 更新CAD图框标签 - 通过打开文档方式
    /// </summary>
    /// <param name="filePath">CAD文件路径</param>
    /// <param name="globalMarks">全局标记字典（应用于所有图纸）</param>
    /// <param name="specificMarks">图纸特定标记字典（仅应用于匹配的图纸）</param>
    /// <returns>是否成功更新</returns>
    public async Task<bool> UpdateCADFrameAttributes(string filePath, Dictionary<string, string> globalMarks, Dictionary<string, Dictionary<string, string>> specificMarks)
    {
        try
        {
            OnLogMessage($"开始处理CAD文件: {Path.GetFileName(filePath)}");

            string fileName = Path.GetFileNameWithoutExtension(filePath);
            bool hasSpecificMarks = specificMarks.Keys.Where(s => fileName.Contains(s)).Count() > 0;

            if (hasSpecificMarks)
            {
                OnLogMessage($"文件 {fileName} 存在特定标记配置");
            }

            // 先创建一个合并的标记字典，特定标记优先级高于全局标记
            Dictionary<string, string> mergedMarks = new Dictionary<string, string>();

            // 添加全局标记
            if (globalMarks != null && globalMarks.Count > 0)
            {
                foreach (var mark in globalMarks)
                {
                    mergedMarks[mark.Key] = mark.Value.TrimStart().TrimEnd();
                }
            }

            // 添加并覆盖特定标记（如果有）
            if (hasSpecificMarks)
            {
                // 找到fileName包含的第一个键
                string matchedKey = specificMarks.Keys.FirstOrDefault(s => fileName.Contains(s));
                if (matchedKey != null && specificMarks[matchedKey] != null)
                {
                    foreach (var mark in specificMarks[matchedKey])
                    {
                        mergedMarks[mark.Key] = mark.Value.TrimStart().TrimEnd();
                    }
                }
            }

            // 使用文档方式更新CAD
            bool result = UpdateCADWithDocument(filePath, mergedMarks);

            if (result)
            {
                OnLogMessage($"文件 {fileName} 已成功更新");
            }
            else
            {
                OnLogMessage($"文件 {fileName} 更新失败", true);
            }

            return result;
        }
        catch (Exception ex)
        {
            string errorMessage = $"处理文件 {Path.GetFileName(filePath)} 时出错: {ex.Message}";
            OnLogMessage(errorMessage, true);
            return false;
        }
    }

    /// <summary>
    /// 通过打开文档的方式更新CAD
    /// </summary>
    private bool UpdateCADWithDocument(string filePath, Dictionary<string, string> attributeValues)
    {
        try
        {
            DocumentCollection docs = Application.DocumentManager;

            // 打开文档
            var currentDoc = docs.Open(filePath, false);
            Application.DocumentManager.MdiActiveDocument = currentDoc;
           
            using (currentDoc.LockDocument())
            {
                bool anyUpdated = false;
                using (var util = new DrawingUtility(currentDoc.Database))
                {
                    // 确保在图纸空间
                    if (util.db.TileMode == true)
                    {
                        util.db.TileMode = false;
                        util.editor.SwitchToPaperSpace();
                    }

                    // 设置合适的视图
                    currentDoc.SendCommand("zoom e ");

                    // 获取所有布局
                    var layouts = util.db.GetAllLayouts();
                    if (layouts.Count == 0)
                    {
                        OnLogMessage($"文件 {Path.GetFileNameWithoutExtension(filePath)} 中未找到任何布局", true);
                        return false;
                    }

                    // 只处理第一个布局
                    string layoutName = layouts[0].LayoutName;
                    OnLogMessage($"处理布局: {layoutName}");

                    // 查找CNCS块
                    var blockRefs = util.GetBlocksInLayout(layoutName, "CNCS");
                    if (blockRefs.Count == 0)
                    {
                        OnLogMessage($"布局 {layoutName} 中未找到CNCS块", true);
                        return false;
                    }

                    OnLogMessage($"布局 {layoutName} 中找到 {blockRefs.Count} 个CNCS块");

                   

                    // 遍历所有块引用
                    foreach (var blockRef in blockRefs)
                    {
                        // 使用原有的属性更新方法，保留消息通知
                        bool blockUpdated = UpdateBlockAttributesInDocument(util, blockRef, attributeValues);
                        if (blockUpdated)
                        {
                            anyUpdated = true;
                        }
                    }
                    util.Flush();
                    // 如果有更新，保存文档
                    if (anyUpdated)
                    {
                        // 如果有正在运行的命令，发送ESC终止
                        if (currentDoc.CommandInProgress != "")
                        {
                            currentDoc.SendCommand("\x03\x03");  // 发送ESC键
                        }

                        // 保存文档
                        currentDoc.SendCommand("_QSAVE ");
                    }
                }
            }

            // 关闭文档
            if (currentDoc.IsActive)
            {
                currentDoc.CloseAndSave(filePath);
            }

            return true;
        }
        catch (Exception ex)
        {
            OnLogMessage($"打开或操作文档时出错: {ex.Message}", true);
            return false;
        }
    }

    /// <summary>
    /// 在打开的文档中更新块属性
    /// </summary>
    private bool UpdateBlockAttributesInDocument(DrawingUtility util, BlockReference blockRef, Dictionary<string, string> attributeValues)
    {
        bool anyUpdated = false;

        // 遍历块参照的属性
        foreach (ObjectId attId in blockRef.AttributeCollection)
        {
            // 获取块参照属性对象
            AttributeReference attRef = attId.GetObject(OpenMode.ForRead) as AttributeReference;
            if (attRef != null)
            {
                string tagName = attRef.Tag;

                // 检查此标签是否需要更新
                if (attributeValues.TryGetValue(tagName, out string tagValue))
                {
                    string currentValue = attRef.TextString;

                    // 只有在值不同时才更新
                    if (currentValue != tagValue)
                    {
                        OnLogMessage($"  更新属性 {tagName}: \"{currentValue}\" -> \"{tagValue}\"");

                        attRef.UpgradeOpen();
                        // 更新文本内容
                        attRef.TextString = tagValue;
                        attRef.DowngradeOpen();

                        anyUpdated = true;
                    }
                }
            }
        }

        return anyUpdated;
    }

    /// <summary>
    /// 批量更新多个CAD文件的图框标签
    /// </summary>
    public async Task<(int Success, int Failure)> BatchUpdateCADFrames(List<string> filePaths, Dictionary<string, string> globalMarks, Dictionary<string, Dictionary<string, string>> specificMarks)
    {
        int successCount = 0;
        int failureCount = 0;

        foreach (string filePath in filePaths)
        {
            try
            {
                bool success = await UpdateCADFrameAttributes(filePath, globalMarks, specificMarks);
                if (success) successCount++;
                else failureCount++;
            }
            catch (Exception ex)
            {
                failureCount++;
                OnLogMessage($"处理文件 {Path.GetFileName(filePath)} 时发生异常: {ex.Message}", true);
            }
        }

        return (successCount, failureCount);
    }
}