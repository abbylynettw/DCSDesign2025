using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.GraphicsInterface;
using log4net;
using MyOffice;
using WinformUI.CADHelper;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace WinformUI
{
    /// <summary>
    /// DWG文件处理器 - 负责处理DWG文件内容的替换
    /// </summary>
    public class DwgProcessor
    {
        // 日志记录器
        private static readonly ILog Log = LogManager.GetLogger(typeof(DwgProcessor));

        // 日志构建器 - 保留用于向调用者返回处理日志
        private StringBuilder _logBuilder;

        public DwgProcessor(StringBuilder logBuilder)
        {
            _logBuilder = logBuilder;
        }

        /// <summary>
        /// 处理DWG文件 - 返回Task以保持异步调用接口不变，但内部不使用Task.Run
        /// </summary>
        public Task<bool> ProcessDwgFile(string filePath, List<ReplaceRule> replaceRules, ProcessingOptions options,StringBuilder log)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                Log.Info($"开始处理DWG文件: {fileName}");
                _logBuilder.AppendLine($": -----------------------------");
                _logBuilder.AppendLine($"开始处理DWG文件: {fileName}");
             
                // 直接在当前线程处理，不使用Task.Run
                bool result;
                try
                {
                    result = ProcessDwgFileInternal(filePath, replaceRules, options, log);
                }
                catch (Exception ex)
                {
                    string errorMessage = $"处理DWG文件时出错: {ex.Message}";
                    Log.Error(errorMessage, ex);
                    _logBuilder.AppendLine(errorMessage);
                    _logBuilder.AppendLine($"堆栈跟踪: {ex.StackTrace}");
                    result = false;
                }

                if (result)
                {
                    Log.Info($"【成功】DWG文件处理成功: {fileName}");
                    _logBuilder.AppendLine($"DWG文件处理成功: {fileName}");
                }
                else
                {
                    Log.Warn($"【失败】DWG文件处理失败: {fileName}");
                    _logBuilder.AppendLine($"DWG文件处理失败: {fileName}");
                }

                // 将结果包装在已完成的Task中返回
                return Task.FromResult(result);
            }
            catch (Exception ex)
            {
                string errorMessage = $"【失败】处理DWG文件时出错: {ex.Message}";
                Log.Error(errorMessage, ex);
                _logBuilder.AppendLine(errorMessage);
                return Task.FromResult(false);
            }
        }

        private bool ProcessDwgFileInternal(string filePath, List<ReplaceRule> replaceRules, ProcessingOptions options, StringBuilder log)
        {
            Log.Debug($"开始处理DWG文件: {Path.GetFileName(filePath)}");

            try
            {
                // 获取文档集合
                DocumentCollection docs = Application.DocumentManager;

                // 打开文档
                var currentDoc = docs.Open(filePath, false);
                Application.DocumentManager.MdiActiveDocument = currentDoc;

                using (currentDoc.LockDocument())
                {
                    int replacementCount = 0;
                    bool success = false;

                    // 在单独的代码块中使用DrawingUtility，确保使用完后立即释放
                    using (DrawingUtility util = new DrawingUtility(currentDoc.Database))
                    {
                        try
                        {
                            // 处理文本内容
                            if (options.ProcessDwgText)
                            {
                                replacementCount += ProcessDwgTexts(util, replaceRules, log);
                            }

                            // 处理表格内容
                            if (options.ProcessDwgTableContent)
                            {
                                replacementCount += ProcessDwgTables(util, replaceRules, log);
                            }

                            // 处理块属性
                            if (options.ProcessDwgBlockAttributes)
                            {
                                replacementCount += ProcessDwgBlockAttributes(util, replaceRules, log);
                            }

                            if (options.ProcessDowngradeA)
                            {
                                replacementCount += ProcessDowngradeA(util, log);
                            }

                            // 提交事务
                            util.Commit();
                            success = true;

                            // 保存更改
                            if (success && replacementCount > 0)
                            {
                                // 如果有正在运行的命令，发送ESC终止
                                if (currentDoc.CommandInProgress != "")
                                {
                                    currentDoc.SendCommand("\x03\x03");  // 发送ESC键
                                }

                                // 保存文档
                                currentDoc.SendCommand("_QSAVE ");

                                string resultMessage = $"替换完成，共替换 {replacementCount} 处内容";
                                Log.Info(resultMessage);
                                _logBuilder.AppendLine(resultMessage);
                            }
                        }
                        catch (Exception ex)
                        {
                            string errorMessage = $"处理DWG文件内容时出错: {ex.Message}";
                            Log.Error(errorMessage, ex);
                            _logBuilder.AppendLine(errorMessage);

                            Log.Debug("中止事务");
                            util.Abort();
                            throw; // 重新抛出异常，允许外层catch捕获
                        }
                    } // 在这里DrawingUtility会被Dispose
                   
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
                string errorMessage = $"读取或打开DWG文件时出错: {ex.Message}";
                Log.Error(errorMessage, ex);
                _logBuilder.AppendLine(errorMessage);
                return false;
            }
        }
        /// <summary>
        /// 内部处理DWG文件的方法
        /// </summary>
        private bool ProcessDwgFileInternalOld(string filePath, List<ReplaceRule> replaceRules, ProcessingOptions options,StringBuilder log)
        {
            Log.Debug($"开始内部处理DWG文件: {Path.GetFileName(filePath)}");

            // 创建一个新的数据库对象
            using (Database db = new Database(false, true))
            {
                try
                {
                    // 将DWG文件读入数据库
                    Log.Debug($"----读取DWG文件到数据库: {filePath}");
                    db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, "");

                    int replacementCount = 0;
                    bool success = false;

                    // 在单独的代码块中使用DrawingUtility，确保使用完后立即释放
                    {
                        using (DrawingUtility util = new DrawingUtility(db, false))
                        {
                            try
                            {
                                // 处理文本内容
                                if (options.ProcessDwgText)
                                {                                   
                                    replacementCount += ProcessDwgTexts(util, replaceRules, log);
                                }

                                // 处理表格内容
                                if (options.ProcessDwgTableContent)
                                {                                   
                                    replacementCount += ProcessDwgTables(util, replaceRules,log);
                                }

                                // 处理块属性
                                if (options.ProcessDwgBlockAttributes)
                                {                                   
                                    replacementCount += ProcessDwgBlockAttributes(util, replaceRules,log);
                                }

                                if (options.ProcessDowngradeA)
                                {
                                    replacementCount += ProcessDowngradeA(util, log);
                                }

                                // 提交事务 - 重要：事务必须在使用完后立即提交                               
                                util.Commit();
                                success = true;
                            }
                            catch (Exception ex)
                            {
                                string errorMessage = $"处理DWG文件内容时出错: {ex.Message}";
                                Log.Error(errorMessage, ex);
                                _logBuilder.AppendLine(errorMessage);

                                Log.Debug("中止事务");
                                util.Abort();
                                throw; // 重新抛出异常，允许外层catch捕获
                            }
                        } // 在这里DrawingUtility会被Dispose
                    }

                    // 只有当所有操作成功时才保存DWG文件
                    if (success)
                    {
                        // 保存DWG文件
                        Log.Debug($"保存DWG文件: {filePath}");
                        db.SaveAs(filePath, DwgVersion.Current);

                        string resultMessage = $"替换完成，共替换 {replacementCount} 处内容";
                        Log.Info(resultMessage);
                        _logBuilder.AppendLine(resultMessage);
                        return true;
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    string errorMessage = $"读取或打开DWG文件时出错: {ex.Message}";
                    Log.Error(errorMessage, ex);
                    _logBuilder.AppendLine(errorMessage);
                    return false;
                }
            }
        }
        /// <summary>
        /// 处理降低A版本 - 修改目录页表格和图框块属性
        /// </summary>
        private int ProcessDowngradeA(DrawingUtility util, StringBuilder log)
        {
            int replacementCount = 0;
            log.AppendLine("---开始执行降低A版本处理:");

            try
            {
                // 检查是否为目录页
                string fileName = System.IO.Path.GetFileName(util.db.Filename ?? "");
                bool isDirectoryPage = fileName.Contains("DI-001") || fileName.Contains("DI-002") || fileName.Contains("DI-003");
                log.AppendLine($"当前文件: {fileName}，是否为目录页: {isDirectoryPage}");

                // 只有目录页才处理表格
                if (isDirectoryPage)
                {
                    log.AppendLine("---处理目录页表格:");
                    List<Table> tables = util.GetAllEntity<Table>();
                    log.AppendLine($"找到 {tables.Count} 个表格对象");

                    foreach (Table table in tables)
                    {
                        log.AppendLine($"处理表格: {table.ObjectId} (行数: {table.Rows.Count}, 列数: {table.Columns.Count})");
                        table.UpgradeOpen();
                        table.RemoveDataLink();

                        // 查找版本列的索引
                        int revColumnIndex = -1;
                        for (int col = 0; col < table.Columns.Count; col++)
                        {
                            string headerText = table.Cells[0, col].GetTextString(FormatOption.IgnoreMtextFormat);
                            if (headerText.Contains("版本")){revColumnIndex = col;log.AppendLine($"找到版本列，列索引: {col}, 列标题: \"{headerText}\"");
                                break; }
                        }

                        // 只处理版本列
                        if (revColumnIndex >= 0)
                        {
                            int tableReplacements = 0;
                            for (int row = 1; row < table.Rows.Count; row++) // 从第二行开始（跳过表头）
                            {
                                string cellText = table.Cells[row, revColumnIndex].GetTextString(FormatOption.IgnoreMtextFormat);
                                string oldValue = cellText;
                                if (string.IsNullOrEmpty(cellText)) continue;
                                // 记录并替换为A
                                table.Cells[row, revColumnIndex].SetValue("A", ParseOption.ParseOptionNone);
                                tableReplacements++;

                                log.AppendLine($"  行 {row}, 列 {revColumnIndex}: 替换 \"{oldValue}\" -> \"A\"");
                            }

                            log.AppendLine($"表格处理完成，共替换 {tableReplacements} 个单元格");
                            replacementCount += tableReplacements;
                        }
                        else
                        {
                            log.AppendLine("未找到版本列，跳过此表格");
                        }

                        table.DowngradeOpen();
                    }
                }

                // 处理图框块属性
                log.AppendLine("---处理图框块属性:");
                BlockReference cncsBlock = null;
                List<BlockReference> paperSpaceBlocks = util.GetAllEntityInPaperSpace<BlockReference>();
                log.AppendLine($"找到 {paperSpaceBlocks.Count} 个布局空间块引用");

                foreach (BlockReference br in paperSpaceBlocks)
                    if (br.Name.ToUpper() == "CNCS")
                    {
                        cncsBlock = br;
                        log.AppendLine($"找到CNCS图框块，ID: {cncsBlock.ObjectId}");
                        break;
                    }

                if (cncsBlock != null)
                {
                    cncsBlock.UpgradeOpen();
                    int blockReplacements = 0;

                    log.AppendLine("处理图框块属性:");
                    foreach (ObjectId attId in cncsBlock.AttributeCollection)
                    {
                        AttributeReference attRef = attId.GetObject(OpenMode.ForWrite) as AttributeReference;
                        if (attRef == null) continue;

                        string tag = attRef.Tag;
                        string oldValue = attRef.TextString;

                        if (tag == "REV" || tag == "REV2")
                        {
                            attRef.TextString = "A";
                            blockReplacements++;
                            log.AppendLine($"  属性 \"{tag}\": 替换 \"{oldValue}\" -> \"A\"");
                        }
                        if (tag=="DATE_A")
                        {// 获取当前日期格式为"月/年"                            
                            attRef.TextString = DateTime.Now.ToString("MM/yy"); ;
                            blockReplacements++;
                            log.AppendLine($"  属性 \"{tag}\": 替换 \"{oldValue}\" -> \"{attRef.TextString}\"");
                        }
                        else if (tag.Contains("_B") || tag.Contains("_C") || tag.Contains("_D") ||
                                 tag.Contains("_E") || tag.Contains("_F"))
                        {
                            attRef.TextString = "";
                            blockReplacements++;
                            log.AppendLine($"  属性 \"{tag}\": 清空 \"{oldValue}\" -> \"\"");
                        }
                    }

                    log.AppendLine($"图框块处理完成，共修改 {blockReplacements} 个属性");
                    replacementCount += blockReplacements;

                    cncsBlock.DowngradeOpen();
                }
                else
                {
                    log.AppendLine("未找到CNCS图框块，跳过处理");
                }

                string resultMessage = $"【成功】版本降低完成，共修改 {replacementCount} 处内容";
                log.AppendLine(resultMessage);
                Log.Info(resultMessage);
            }
            catch (Exception ex)
            {
                string errorMessage = $"【失败】降低版本处理出错: {ex.Message}";
                log.AppendLine(errorMessage);
                log.AppendLine($"异常堆栈: {ex.StackTrace}");
                Log.Error(errorMessage, ex);
            }

            log.AppendLine("---降低A版本处理结束");
            return replacementCount;
        }
        /// <summary>
        /// 处理DWG文件中的文本内容
        /// </summary>
        private int ProcessDwgTexts(DrawingUtility util, List<ReplaceRule> replaceRules, StringBuilder log)
        {
            int replacementCount = 0;
            log.AppendLine("---开始处理DWG文本内容:");

            try
            {
                // 处理单行文本
                List<DBText> dbTexts = util.GetAllEntity<DBText>();
               
                foreach (DBText dbText in dbTexts)
                {
                    try
                    {
                        string originalText = dbText.TextString;
                        string newText = ApplyReplaceRules(originalText, replaceRules);

                        if (originalText != newText)
                        {
                            dbText.UpgradeOpen();
                            dbText.TextString = newText;
                            dbText.DowngradeOpen();
                            replacementCount++;

                            string logMessage = $"替换单行文本: \"{originalText}\" -> \"{newText}\"";
                            log.AppendLine(logMessage);
                            Log.Debug(logMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"处理单行文本时出错: {ex.Message}";
                        log.AppendLine(errorMsg);
                        Log.Error(errorMsg, ex);
                    }
                }

                // 处理多行文本
                List<MText> mTexts = util.GetAllEntity<MText>();
              
                foreach (MText mText in mTexts)
                {
                    try
                    {
                        string originalText = mText.Contents;
                        string newText = ApplyReplaceRules(originalText, replaceRules);

                        if (originalText != newText)
                        {
                            mText.UpgradeOpen();
                            mText.Contents = newText;
                            mText.DowngradeOpen();
                            replacementCount++;

                            string shortOriginal = originalText.Length > 20 ? originalText.Substring(0, 20) + "..." : originalText;
                            string shortNew = newText.Length > 20 ? newText.Substring(0, 20) + "..." : newText;
                            string logMessage = $"替换多行文本: \"{shortOriginal}\" -> \"{shortNew}\"";

                            log.AppendLine(logMessage);
                            Log.Debug(logMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"处理多行文本时出错: {ex.Message}";
                        log.AppendLine(errorMsg);
                        Log.Error(errorMsg, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"处理文本对象时出错: {ex.Message}";
                log.AppendLine(errorMsg);
                Log.Error(errorMsg, ex);
            }

            log.AppendLine($"文本内容处理完成，共替换 {replacementCount} 处");
            return replacementCount;
        }

        /// <summary>
        /// 处理DWG文件中的表格内容
        /// </summary>
        private int ProcessDwgTables(DrawingUtility util, List<ReplaceRule> replaceRules, StringBuilder log)
        {
            int replacementCount = 0;
            log.AppendLine("---开始处理DWG表格内容:");

            try
            {
                // 处理表格
                List<Table> tables = util.GetAllEntity<Table>();
              
                int tableIndex = 0;
                foreach (Table table in tables)
                {
                    tableIndex++;
                    try
                    {                       
                        table.UpgradeOpen();
                        table.RemoveDataLink();

                        int tableCellReplaced = 0;
                        // 遍历表格每个单元格
                        for (int row = 0; row < table.Rows.Count; row++)
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                try
                                {
                                    string originalText = table.Cells[row, col].TextString;
                                    string newText = ApplyReplaceRules(originalText, replaceRules);

                                    if (originalText != newText)
                                    {
                                        table.Cells[row, col].TextString = newText;
                                        replacementCount++;
                                        tableCellReplaced++;

                                        string shortOriginal = originalText.Length > 20 ? originalText.Substring(0, 20) + "..." : originalText;
                                        string shortNew = newText.Length > 20 ? newText.Substring(0, 20) + "..." : newText;
                                        string logMessage = $"  单元格[{row},{col}]: \"{shortOriginal}\" -> \"{shortNew}\"";

                                        log.AppendLine(logMessage);
                                        Log.Debug(logMessage);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string errorMsg = $"处理表格单元格[{row},{col}]时出错: {ex.Message}";
                                    log.AppendLine(errorMsg);
                                    Log.Error(errorMsg, ex);
                                }
                            }
                        }

                        log.AppendLine($"表格 #{tableIndex} 处理完成，替换了 {tableCellReplaced} 个单元格");
                        table.DowngradeOpen();
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"处理表格 #{tableIndex} 时出错: {ex.Message}";
                        log.AppendLine(errorMsg);
                        Log.Error(errorMsg, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"处理表格对象时出错: {ex.Message}";
                log.AppendLine(errorMsg);
                Log.Error(errorMsg, ex);
            }

            log.AppendLine($"表格内容处理完成，共替换 {replacementCount} 处");
            return replacementCount;
        }

        /// <summary>
        /// 处理DWG文件中的块属性
        /// </summary>
        private int ProcessDwgBlockAttributes(DrawingUtility util, List<ReplaceRule> replaceRules, StringBuilder log)
        {
            int replacementCount = 0;
            log.AppendLine("---开始处理DWG块属性:");

            try
            {
                // 处理块参照中的属性
                List<BlockReference> blockRefs = util.GetAllEntityInPaperSpace<BlockReference>();               

                int blockCount = 0;
                foreach (BlockReference blockRef in blockRefs)
                {
                    blockCount++;
                    try
                    {
                        int blockAttribsReplaced = 0;

                        foreach (ObjectId attId in blockRef.AttributeCollection)
                        {
                            try
                            {
                                // 获取块参照属性对象
                                AttributeReference attRef = attId.GetObject(OpenMode.ForRead) as AttributeReference;
                                if (attRef != null)
                                {
                                    // 替换属性文本
                                    string originalText = attRef.TextString;
                                    string newText = ApplyReplaceRules(originalText, replaceRules);

                                    if (originalText != newText)
                                    {
                                        attRef.UpgradeOpen();
                                        attRef.TextString = newText;
                                        attRef.DowngradeOpen();
                                        replacementCount++;
                                        blockAttribsReplaced++;
                                        string logMessage = $"  属性[{attRef.Tag}]: \"{originalText}\" -> \"{newText}\"";
                                        log.AppendLine(logMessage);
                                        Log.Debug(logMessage);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                string errorMsg = $"处理块属性时出错: {ex.Message}";
                                log.AppendLine(errorMsg);
                                Log.Error(errorMsg, ex);
                            }
                        }

                        if (blockAttribsReplaced > 0)
                        {
                            log.AppendLine($"块参照 #{blockCount} 处理完成，替换了 {blockAttribsReplaced} 个属性");
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"处理块参照 #{blockCount} 时出错: {ex.Message}";
                        log.AppendLine(errorMsg);
                        Log.Error(errorMsg, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"处理块参照对象时出错: {ex.Message}";
                log.AppendLine(errorMsg);
                Log.Error(errorMsg, ex);
            }

            log.AppendLine($"块属性处理完成，共替换 {replacementCount} 处");
            return replacementCount;
        }      

        /// <summary>
        /// 应用替换规则到文本
        /// </summary>
        private string ApplyReplaceRules(string text, List<ReplaceRule> rules)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            string result = text;

            foreach (ReplaceRule rule in rules)
            {
                // 跳过空规则
                if (string.IsNullOrEmpty(rule.FindText))
                    continue;

                try
                {
                    switch (rule.Type)
                    {
                        case RuleType.包含:
                            // 包含匹配：将所有匹配的子字符串替换为新文本
                            result = result.Replace(rule.FindText, rule.ReplaceText);
                            break;

                        case RuleType.相等:
                            // 完全匹配：只有当文本完全匹配时才替换
                            if (result == rule.FindText)
                            {
                                result = rule.ReplaceText;
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"应用替换规则时出错: {rule.FindText} -> {rule.ReplaceText}", ex);
                }
            }

            return result;
        }
    }
}