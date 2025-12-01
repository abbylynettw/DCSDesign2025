using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using log4net;
using MyOffice;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinformUI;
using WinformUI.CADHelper;
using WinformUI.UpdateTable;
using static System.Windows.Forms.Design.AxImporter;
using static WinformUI.UpdateTable.Model;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using DataTable = System.Data.DataTable;

/// <summary>
/// CAD到Excel的数据导出处理类
/// </summary>
public class CADTableExporter
{
    // 日志记录器
    private static readonly ILog Log = LogManager.GetLogger(typeof(CADTableExporter));
    // 日志构建器
    private StringBuilder _logBuilder;
    private ExcelProcessor excelProcessor;
    // 合并单元格支持开关
    public bool EnableMergeCells { get; set; } = false;

    // 定义日志事件和事件参数类
    public class LogEventArgs : EventArgs
    {
        public string Message { get; set; }
        public InfoType InfoType { get; set; }

        public LogEventArgs(string message, InfoType infoType = InfoType.Info)
        {
            Message = message;
            InfoType = infoType;
        }
    }
    public enum InfoType
    {
        Info,
        Warning,
        Error,
        Success
    }
    // 声明日志事件
    public event EventHandler<LogEventArgs> LogMessageGenerated;

    // 触发日志事件的方法
    protected virtual void OnLogMessage(string message, InfoType infoType = InfoType.Info)
    {
        LogMessageGenerated?.Invoke(this, new LogEventArgs(message, infoType));
        _logBuilder.AppendLine(message);
    }

    public CADTableExporter(StringBuilder logBuilder = null)
    {
        _logBuilder = logBuilder ?? new StringBuilder();
        excelProcessor = new ExcelProcessor();
    }

    #region Excel to CAD
    public async Task<bool> ExcelToCAD(List<MappingConfig> mappings, List<string> processFiles)
    {
        try
        {
            #region 已检验代码

            #endregion
            OnLogMessage("开始处理Excel到CAD的数据更新");
            if (mappings == null || mappings.Count == 0)
            {
                OnLogMessage("映射配置为空，无法继续处理", InfoType.Error);
                return false;
            }

            // 过滤出DWG文件和Excel文件
            var dwgs = processFiles.Where(p => Path.GetExtension(p).ToLower() == ".dwg").ToList();
            var excels = processFiles.Where(p => Path.GetExtension(p).ToLower() == ".xlsx" || Path.GetExtension(p).ToLower() == ".xls")
                                     .Where(x => Path.GetFileNameWithoutExtension(x).ToLower() == "index").ToList();

            if (dwgs.Count == 0)
            {
                OnLogMessage("未找到DWG文件，无法继续处理", InfoType.Error);
                return false;
            }

            if (excels.Count == 0)
            {
                OnLogMessage("未找到Excel文件，无法继续处理", InfoType.Error);
                return false;
            }

            // 使用第一个Excel文件
            string excelFilePath = excels[0];
            OnLogMessage($"读取Excel文件: {Path.GetFileName(excelFilePath)}");

            // 创建字典，用于存储Excel数据
            Dictionary<string, DataTable> excelDataBySheetName = [];

            // 读取Excel数据，根据两种匹配关系
            foreach (var mapping in mappings.Where(m => m.Enabled))
            {
                try
                {
                    // 读取Excel工作表数据并存储
                    if (!string.IsNullOrEmpty(mapping.ExcelSheet))
                    {
                        DataTable sheetData = await excelProcessor.ReadExcelSheet(excelFilePath, mapping.ExcelSheet);
                        // 如果启用合并单元格功能，读取合并单元格信息
                        if (EnableMergeCells && sheetData != null && sheetData.Rows.Count > 0)
                        {
                            var mergedCells = await excelProcessor.ReadExcelMergedCells(excelFilePath, mapping.ExcelSheet);
                            if (mergedCells != null && mergedCells.Count > 0)
                            {
                                sheetData.ExtendedProperties["MergedCells"] = mergedCells;
                                OnLogMessage($"  从Excel工作表 '{mapping.ExcelSheet}' 读取了 {mergedCells.Count} 个合并单元格");
                            }
                        }
                        if (sheetData != null && sheetData.Rows.Count > 0)
                        {
                            // 使用Excel工作表名称作为键
                            excelDataBySheetName[mapping.ExcelSheet] = sheetData;
                            OnLogMessage($"  从Excel工作表 '{mapping.ExcelSheet}' 中读取了 {sheetData.Rows.Count} 行数据");
                        }
                        else
                        {
                            OnLogMessage($"  Excel工作表 '{mapping.ExcelSheet}' 中没有数据或工作表不存在", InfoType.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    string errorMessage = $"读取Excel工作表 '{mapping.ExcelSheet}' 时出错: {ex.Message}";
                    Log.Error(errorMessage, ex);
                    OnLogMessage(errorMessage, InfoType.Error);
                }
            }

            // 检查是否有读取到数据
            if (excelDataBySheetName.Count == 0)
            {
                OnLogMessage("未从Excel文件中读取到任何数据，无法继续处理", InfoType.Error);
                return false;
            }

            // 第一步：计算每个pageType和tableTitle的总容量（总行数）
            Dictionary<string, int> totalRowsByPageType = [];
            Dictionary<string, int> totalRowsByTableTitle = [];
            // 获取文档集合
            DocumentCollection docs = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            // 第一次遍历：统计总容量
            foreach (var filePath in dwgs)
            {
                var currentDoc = docs.Open(filePath, false);
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = currentDoc;
                using (currentDoc.LockDocument())
                {
                    try
                    {
                        //// 将DWG文件读入数据库
                        //db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, "");
                        using (DrawingUtility util =new DrawingUtility(currentDoc.Database))
                        {
                            try
                            {
                                // 获取页面类型
                                string pageType = util.GetDwgSummary("PAGETYPE");
                                // 获取所有表格
                                var tables = util.GetAllEntity<Table>();                               

                                // 基于页面类型计算容量
                                if (!string.IsNullOrEmpty(pageType))
                                {
                                    var typeMapping = mappings.FirstOrDefault(m => m.DwgKeyword == pageType && m.Enabled);
                                    if (typeMapping != null && excelDataBySheetName.ContainsKey(typeMapping.ExcelSheet))
                                    {
                                        if (tables.Count >= 1)
                                        {
                                            // 表格数据行数（减去表头）
                                            int tableRows = tables[0].Rows.Count - 1;

                                            // 累加页面类型的总行数
                                            if (!totalRowsByPageType.ContainsKey(pageType))
                                            {
                                                totalRowsByPageType[pageType] = 0;
                                            }
                                            totalRowsByPageType[pageType] += tableRows;
                                            OnLogMessage($"  统计: 页面类型 '{pageType}' 的CAD文件 {Path.GetFileName(filePath)} 有 {tableRows} 行数据容量");
                                        }
                                    }
                                }
                                // 基于表格标题计算容量                        
                                else if (tables.Count > 0)
                                {
                                    // 创建一个集合来跟踪已处理过的表格标题
                                    HashSet<string> processedTableTitles = [];
                                    // 遍历文件中的所有表格
                                    for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
                                    {
                                        Table currentTable = tables[tableIndex];

                                        // 获取表格的第一行作为表格标题
                                        string tableTitle = null;
                                        if (currentTable.Rows.Count > 0 && currentTable.Columns.Count > 0)
                                        {
                                            tableTitle = currentTable.Cells[0, 0].GetTextString(FormatOption.IgnoreMtextFormat).Trim();
                                        }

                                        if (!string.IsNullOrEmpty(tableTitle))
                                        {
                                            // 检查这个表格标题是否已经处理过
                                            if (processedTableTitles.Contains(tableTitle)) continue;

                                            var titleMapping = mappings.FirstOrDefault(m => m.CADTableTitle == tableTitle && m.Enabled);
                                            if (titleMapping != null && excelDataBySheetName.ContainsKey(titleMapping.ExcelSheet))
                                            {
                                                // 表格数据行数（减去表头）
                                                int tableRows = currentTable.Rows.Count - 1;
                                                // 累加表格标题的总行数
                                                if (!totalRowsByTableTitle.ContainsKey(tableTitle))
                                                {
                                                    totalRowsByTableTitle[tableTitle] = 0;
                                                }
                                                totalRowsByTableTitle[tableTitle] += tableRows;

                                                OnLogMessage($"  统计: 表格标题 '{tableTitle}' 的CAD文件 {Path.GetFileName(filePath)} 有 {tableRows} 行数据容量");
                                                // 标记这个表格标题为已处理
                                                processedTableTitles.Add(tableTitle);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                string errorMessage = $"统计DWG文件容量时出错: {ex.Message}";
                                Log.Error(errorMessage, ex);
                                OnLogMessage(errorMessage, InfoType.Error);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMessage = $"读取或打开DWG文件时出错: {ex.Message}";
                        Log.Error(errorMessage, ex);
                        OnLogMessage(errorMessage, InfoType.Error);
                    }
                }
                // 关闭文档
                if (currentDoc.IsActive)
                {
                    currentDoc.CloseAndSave(filePath);
                }
            }           
            // 替换为:
            PrintStatisticsByMapping(mappings, excelDataBySheetName, totalRowsByPageType, totalRowsByTableTitle);
            // 为每个页面类型创建已使用行数的字典
            Dictionary<string, int> usedRowsByPageType = [];
            // 为每个表格标题创建已使用行数的字典
            Dictionary<string, int> usedRowsByTableTitle = [];

            // 处理每个DWG文件
            int successCount = 0;

            // 第二次遍历：执行更新
            foreach (var filePath in dwgs)
            {
              
                bool fileUpdated = false;
             
                // 打开文档
                var currentDoc = docs.Open(filePath, false);
                Application.DocumentManager.MdiActiveDocument = currentDoc;

                using (currentDoc.LockDocument())
                {
                    try
                    {
                        //// 将DWG文件读入数据库
                        //db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, "");

                        using (DrawingUtility util = new(currentDoc.Database))
                        {
                            try
                            {
                                // 获取页面类型
                                string pageType = util.GetDwgSummary("PAGETYPE");

                                // 获取所有表格
                                var tables = util.GetAllEntity<Table>();                               
                               
                                // 检查是否有表格
                                if (tables.Count > 0)
                                {
                                    OnLogMessage($"【开始】{Path.GetFileName(filePath)}|页面类型: {pageType}");
                                    // 普通DWG文件处理

                                    // 方式1: 根据DwgKeyword(PageType)匹配
                                    if (!string.IsNullOrEmpty(pageType))
                                    {
                                        var typeMapping = mappings.FirstOrDefault(m => m.DwgKeyword == pageType && m.Enabled);
                                        if (typeMapping != null && excelDataBySheetName.ContainsKey(typeMapping.ExcelSheet))
                                        {
                                            OnLogMessage($" ----> 匹配的Excel工作表 '{typeMapping.ExcelSheet}'");

                                            if (tables.Count >= 1)
                                            {
                                                // 获取当前页面类型已使用的行数
                                                if (!usedRowsByPageType.ContainsKey(pageType))
                                                {
                                                    usedRowsByPageType[pageType] = 0;
                                                }

                                                // 检查Excel数据是否已经全部用完
                                                if (usedRowsByPageType[pageType] >= excelDataBySheetName[typeMapping.ExcelSheet].Rows.Count)
                                                {
                                                    OnLogMessage($"  Excel工作表 '{typeMapping.ExcelSheet}' 的数据已全部处理完毕，清空剩余的CAD表格行");

                                                    // 清空整个表格（保留表头）
                                                   
                                                    bool tableUpdated = ClearTableData(tables[0]);
                                                    if (tableUpdated)
                                                    {
                                                        fileUpdated = true;
                                                    }
                                                    continue;
                                                }

                                                // 检查是否有足够的Excel数据行来填充这个表格
                                                int remainingExcelRows = excelDataBySheetName[typeMapping.ExcelSheet].Rows.Count - usedRowsByPageType[pageType];
                                                // 获取CAD表格数据行数（排除表头）
                                                int cadTableDataRows = tables[0].Rows.Count - 1;

                                                if (remainingExcelRows <= 0)
                                                {
                                                    OnLogMessage($"  Excel工作表 '{typeMapping.ExcelSheet}' 中没有更多数据行可用于页面类型 '{pageType}'");

                                                    // 清空整个表格（保留表头）
                                                    bool tableUpdated = ClearTableData(tables[0]);
                                                    if (tableUpdated)
                                                    {
                                                        fileUpdated = true;
                                                    }
                                                    continue;
                                                }

                                                // 检查是否超出了总容量
                                                int totalCapacity = totalRowsByPageType[pageType];
                                                int excelTotalRows = excelDataBySheetName[typeMapping.ExcelSheet].Rows.Count;

                                                if (excelTotalRows > totalCapacity && usedRowsByPageType[pageType] + cadTableDataRows > totalCapacity)
                                                {
                                                    // 仅更新到总容量为止
                                                    int rowsToProcess = totalCapacity - usedRowsByPageType[pageType];
                                                    if (rowsToProcess <= 0)
                                                    {
                                                        // 已经没有剩余容量，清空表格
                                                        OnLogMessage($"  页面类型 '{pageType}' 的总容量已用完，清空剩余的CAD表格");
                                                        bool tbUpdated = ClearTableData(tables[0]);
                                                        if (tbUpdated)
                                                        {
                                                            fileUpdated = true;
                                                        }
                                                        continue;
                                                    }

                                                    OnLogMessage($"  警告: 即将到达页面类型 '{pageType}' 的总容量上限，本表格仅处理 {rowsToProcess} 行");

                                                    // 更新表格数据
                                                    bool tableUpdated = UpdateTableData(
                                                        tables[0],
                                                        excelDataBySheetName[typeMapping.ExcelSheet],
                                                        usedRowsByPageType[pageType],
                                                        rowsToProcess, EnableMergeCells
                                                    );

                                                    if (tableUpdated)
                                                    {
                                                        //再提示

                                                        fileUpdated = true;
                                                        // 更新已使用的行数
                                                        usedRowsByPageType[pageType] += rowsToProcess;
                                                        OnLogMessage($"  页面类型 '{pageType}' 已使用 {usedRowsByPageType[pageType]} 行Excel数据，达到总容量上限");
                                                    }
                                                }
                                                else
                                                {
                                                    // 正常更新
                                                    int rowsToProcess = Math.Min(remainingExcelRows, cadTableDataRows);

                                                    // 更新表格数据
                                                    bool tableUpdated = UpdateTableData(
                                                        tables[0],
                                                        excelDataBySheetName[typeMapping.ExcelSheet],
                                                        usedRowsByPageType[pageType],
                                                        rowsToProcess, EnableMergeCells
                                                    );

                                                    if (tableUpdated)
                                                    {
                                                        fileUpdated = true;
                                                    }
                                                    // 更新已使用的行数
                                                    usedRowsByPageType[pageType] += rowsToProcess;
                                                    OnLogMessage($"  页面类型 '{pageType}' 已使用 {usedRowsByPageType[pageType]} 行Excel数据");
                                                }
                                            }
                                            else
                                            {
                                                //OnLogMessage($"  文件中未找到表格，跳过");
                                            }
                                        }
                                        else
                                        {
                                            //OnLogMessage($"  未找到页面类型 '{pageType}' 对应的匹配配置或Excel数据", InfoType.Error);
                                        }
                                    }
                                    // 方式2: 根据CADTableTitle匹配                                 
                                    else if (tables.Count > 0)
                                    {
                                        bool anyTableUpdated = false;

                                        // 创建一个集合来跟踪已处理过的表格标题
                                        HashSet<string> processedTableTitles = new HashSet<string>();

                                        // 处理文件中的每个表格
                                        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
                                        {
                                            Table currentTable = tables[tableIndex];

                                            // 获取表格的第一行作为表格标题
                                            string tableTitle = null;
                                            if (currentTable.Rows.Count > 0 && currentTable.Columns.Count > 0)
                                            {
                                                tableTitle = currentTable.Cells[0, 0].GetTextString(FormatOption.IgnoreMtextFormat).Trim();
                                                OnLogMessage($"  表格 #{tableIndex + 1} 检测到表格标题: {tableTitle}");
                                            }

                                            if (!string.IsNullOrEmpty(tableTitle))
                                            {
                                                // 检查这个表格标题是否已经处理过
                                                if (processedTableTitles.Contains(tableTitle))
                                                {
                                                    //OnLogMessage($"  表格 #{tableIndex + 1} 标题 '{tableTitle}' 已经处理过，跳过此表格");
                                                    continue;
                                                }

                                                var titleMapping = mappings.FirstOrDefault(m => m.CADTableTitle == tableTitle && m.Enabled);
                                                if (titleMapping != null && excelDataBySheetName.ContainsKey(titleMapping.ExcelSheet))
                                                {
                                                    OnLogMessage($"  表格 #{tableIndex + 1} 根据表格标题 '{tableTitle}' 找到匹配的Excel工作表 '{titleMapping.ExcelSheet}'");

                                                    // 获取Excel数据
                                                    var excelData = excelDataBySheetName[titleMapping.ExcelSheet];

                                                    // 检查Excel数据是否为空
                                                    if (excelData.Rows.Count == 0)
                                                    {
                                                        OnLogMessage($"  表格 #{tableIndex + 1} Excel工作表 '{titleMapping.ExcelSheet}' 没有数据，清空CAD表格");

                                                        // 清空整个表格（保留表头）                                                      
                                                        bool tableTitleUpdated = ClearTableData(currentTable);
                                                        if (tableTitleUpdated)
                                                        {
                                                            anyTableUpdated = true;
                                                        }
                                                        // 标记这个表格标题为已处理
                                                        processedTableTitles.Add(tableTitle);
                                                        continue; // 继续处理下一个表格
                                                    }

                                                    // 获取CAD表格数据行数（排除表头）和Excel行数中的较小值
                                                    int cadTableDataRows = currentTable.Rows.Count - 1;
                                                    int rowsToProcess = Math.Min(excelData.Rows.Count, cadTableDataRows);

                                                    // 使用按位置匹配的方法更新表格数据                                                  
                                                    bool tableUpdated = UpdateTableByTitle(
                                                        currentTable,
                                                        excelData,
                                                        0, // 始终从Excel第一行开始
                                                        rowsToProcess
                                                    );

                                                    if (tableUpdated)
                                                    {
                                                        anyTableUpdated = true;
                                                        OnLogMessage($"  表格 #{tableIndex + 1} 表格标题 '{tableTitle}' 已更新，使用了Excel工作表 '{titleMapping.ExcelSheet}' 的 {rowsToProcess} 行数据");
                                                    }

                                                    // 标记这个表格标题为已处理
                                                    processedTableTitles.Add(tableTitle);
                                                }
                                                else
                                                {
                                                    // OnLogMessage($"  表格 #{tableIndex + 1} 未找到表格标题 '{tableTitle}' 对应的匹配配置或Excel数据");
                                                }
                                            }
                                            else
                                            {
                                                //OnLogMessage($"  表格 #{tableIndex + 1} 未能获取表格标题，跳过");
                                            }
                                        }

                                        // 如果任意表格被更新，则设置文件已更新标志
                                        if (anyTableUpdated)
                                        {
                                            fileUpdated = true;
                                        }
                                    }
                                }

                                // 提交事务 - 只有在有更新时才保存
                                if (fileUpdated)
                                {
                                    util.Commit();
                                    // 如果有正在运行的命令，发送ESC终止
                                    if (currentDoc.CommandInProgress != "")
                                    {
                                        currentDoc.SendCommand("\x03\x03");  // 发送ESC键
                                    }

                                    // 保存文档
                                    currentDoc.SendCommand("_QSAVE ");
                                    successCount++;
                                }
                                else
                                {
                                    util.Abort();
                                }
                            }
                            catch (Exception ex)
                            {
                                string errorMessage = $"处理DWG文件内容时出错: {ex.Message}";
                                Log.Error(errorMessage, ex);
                                OnLogMessage(errorMessage, InfoType.Error);
                                util.Abort();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMessage = $"读取或打开DWG文件时出错: {ex.Message}";
                        Log.Error(errorMessage, ex);
                        OnLogMessage(errorMessage, InfoType.Error);
                    }
                }
                // 关闭文档
                if (currentDoc.IsActive)
                {
                    currentDoc.CloseAndSave(filePath);
                }
            }

            OnLogMessage($"Excel到CAD数据更新完成，成功更新 {successCount} 个DWG文件",InfoType.Success);
            Log.Info($"Excel到CAD数据更新完成，成功更新 {successCount} 个DWG文件");

            return true;
        }
        catch (Exception ex)
        {
            string errorMessage = $"Excel到CAD数据更新过程发生错误: {ex.Message}";
            Log.Error(errorMessage, ex);
            OnLogMessage(errorMessage, InfoType.Error);
            return false;
        }
    }

    // 打印统计结果 - 以mapping配置为基准的提示
    private void PrintStatisticsByMapping(
        List<MappingConfig> mappings,
        Dictionary<string, DataTable> excelDataBySheetName,
        Dictionary<string, int> totalRowsByPageType,
        Dictionary<string, int> totalRowsByTableTitle)
    {
        OnLogMessage("========== 根据前端配置计算匹配情况 ==========");

        // 遍历所有启用的mapping
        foreach (var mapping in mappings.Where(m => m.Enabled))
        {
            // 检查Excel数据是否存在
            if (!excelDataBySheetName.ContainsKey(mapping.ExcelSheet))
            {
                OnLogMessage($"配置项 '{{ExcelSheet: {mapping.ExcelSheet}}}' 未找到对应的Excel工作表数据", InfoType.Error);
                continue;
            }

            // 获取Excel数据行数
            int excelRows = excelDataBySheetName[mapping.ExcelSheet].Rows.Count;

            // 根据PageType或TableTitle类型分别处理
            if (!string.IsNullOrEmpty(mapping.DwgKeyword))
            {
                // PageType类型的mapping
                string pageType = mapping.DwgKeyword;

                // 获取此页面类型的总容量
                int totalCapacity = totalRowsByPageType.ContainsKey(pageType) ? totalRowsByPageType[pageType] : 0;

                // 输出此配置的统计信息         
                OnLogMessage($"配置项 '{{PageType: {pageType}, ExcelSheet: {mapping.ExcelSheet}}}'| Excel：{excelRows} 行|CAD: {totalCapacity} 行");
                // 检查是否找到了匹配的CAD文件
                if (totalCapacity == 0)
                {
                    OnLogMessage($"  - 警告: 未找到匹配页面类型 '{pageType}' 的CAD文件", InfoType.Error);
                }
                else if (excelRows > totalCapacity)
                {
                    // 容量不足警告
                    string errorMsg = $"  - 警告: Excel工作表 '{mapping.ExcelSheet}' 的数据行数({excelRows})超出了" +
                                     $"页面类型 '{pageType}' 的所有CAD表格总容量({totalCapacity})，" +
                                     $"超出了 {excelRows - totalCapacity} 行，需要手动修改CAD表格增加新页。";
                    OnLogMessage(errorMsg, InfoType.Error);
                }
                else
                {
                    //OnLogMessage($"  - 容量充足，可处理全部 {excelRows} 行Excel数据");
                }
            }
            else if (!string.IsNullOrEmpty(mapping.CADTableTitle))
            {
                // TableTitle类型的mapping
                string tableTitle = mapping.CADTableTitle;

                // 获取此表格标题的总容量
                int totalCapacity = totalRowsByTableTitle.ContainsKey(tableTitle) ? totalRowsByTableTitle[tableTitle] : 0;

                // 输出此配置的统计信息
                OnLogMessage($"配置项 '{{TableTitle: {tableTitle}, ExcelSheet: {mapping.ExcelSheet}}}'| Excel：{excelRows} 行|CAD: {totalCapacity} 行");            

                // 检查是否找到了匹配的CAD文件
                if (totalCapacity == 0)
                {
                    OnLogMessage($"  - 警告: 未找到匹配表格标题 '{tableTitle}' 的CAD文件", InfoType.Error);
                }
                else if (excelRows > totalCapacity)
                {
                    // 容量不足警告
                    string errorMsg = $"  - 警告: Excel工作表 '{mapping.ExcelSheet}' 的数据行数({excelRows})超出了" +
                                     $"表格标题 '{tableTitle}' 的所有CAD表格总容量({totalCapacity})，" +
                                     $"超出了 {excelRows - totalCapacity} 行，需要手动修改CAD表格增加新页。";
                    OnLogMessage(errorMsg, InfoType.Error);
                }
                else
                {
                    OnLogMessage($"  - 容量充足，可处理全部 {excelRows} 行Excel数据");
                }
            }
        }

        OnLogMessage("==============================");
    }

    /// <summary>
    /// 更新表格数据，以CAD表格结构为准，并记录每个更新的单元格
    /// 如果Excel数据行少于CAD表格行，则清空多余的CAD行
    /// 完全替换原有单元格内容（包括格式）
    /// </summary>
    private bool UpdateTableData(Table table, DataTable excelData, int startRow = 0, int maxRowsToProcess = int.MaxValue,bool EnableMerge=false)
    {
        if (table == null || excelData == null)
        {
            return false;
        }

        try
        {
            table.UpgradeOpen();
            // 处理合并单元格信息
            List<MergedCellInfo> excelMergedCells = new List<MergedCellInfo>();
            if (EnableMergeCells && excelData.ExtendedProperties.ContainsKey("MergedCells"))
            {
                excelMergedCells = excelData.ExtendedProperties["MergedCells"] as List<MergedCellInfo> ?? new List<MergedCellInfo>();
                OnLogMessage($"  检测到 {excelMergedCells.Count} 个Excel合并单元格信息");
            }
            // 获取表头映射
            Dictionary<int, string> headerMapping = new Dictionary<int, string>();
            for (int col = 0; col < table.Columns.Count; col++)
            {
                string headerText = table.Cells[0, col].GetTextString(FormatOption.IgnoreMtextFormat).Trim();
                if (!string.IsNullOrEmpty(headerText))
                {
                    headerMapping.Add(col, headerText);
                }
            }

            // 检查Excel数据是否包含表头
            bool hasMatchingColumns = false;
            foreach (var header in headerMapping.Values)
            {
                if (excelData.Columns.Contains(header))
                {
                    hasMatchingColumns = true;
                    break;
                }
                else
                {
                    OnLogMessage($"  Excel数据中缺少列 '{header}'", InfoType.Error);
                }
            }

            if (!hasMatchingColumns)
            {
                OnLogMessage($"  Excel数据中没有匹配的列，无法更新表格", InfoType.Error);
                return false;
            }

            // 计算实际要处理的行数（CAD表格行数减去表头，Excel可用行数，和指定的最大处理行数中的最小值）
            int cadTableDataRows = table.Rows.Count - 1; // 减去表头行
            int availableExcelRows = excelData.Rows.Count - startRow;
            int rowsToProcess = Math.Min(Math.Min(cadTableDataRows, availableExcelRows), maxRowsToProcess);

            OnLogMessage($" [CAD表格]: {table.Rows.Count} 行, {table.Columns.Count} 列");
            OnLogMessage($" [Excel表格]: {excelData.Rows.Count} 行, {excelData.Columns.Count} 列，起始行为 {startRow + 1}，将处理 {rowsToProcess} 行");

            int updatedCells = 0;
            // 从第二行开始更新数据（跳过表头）
            for (int row = 1; row <= cadTableDataRows; row++)
            {
                bool rowUpdated = false;

                // 如果这一行有对应的Excel数据
                if (row <= rowsToProcess)
                {
                    // 计算对应的Excel数据行索引
                    int excelRowIndex = startRow + (row - 1);

                    // 更新每一列
                    foreach (var colMapping in headerMapping)
                    {
                        int colIndex = colMapping.Key;
                        string header = colMapping.Value;

                        if (excelData.Columns.Contains(header))
                        {
                            // 获取当前CAD单元格的值
                            string currentValue = table.Cells[row, colIndex].GetTextString(FormatOption.IgnoreMtextFormat).Trim();

                            // 获取要设置的新值
                            string newValue = "";
                            if (excelRowIndex >= 0 && excelRowIndex < excelData.Rows.Count)
                            {
                                object cellValue = excelData.Rows[excelRowIndex][header];
                                newValue = (cellValue == null || cellValue == DBNull.Value) ? "" : cellValue.ToString().Trim();
                            }

                            // 仅在值不同时更新
                            if (currentValue != newValue)
                            {
                                try
                                {
                                    // 记录要更新的单元格信息
                                    OnLogMessage($"  更新单元格[{row},{colIndex}]({header}): \"{currentValue}\" -> \"{newValue}\"");

                                    // 直接使用TextString属性完全替换单元格内容
                                    table.Cells[row, colIndex].SetValue(newValue, ParseOption.ParseOptionNone);
                                    updatedCells++;
                                    rowUpdated = true;
                                }
                                catch (Exception ex)
                                {
                                    OnLogMessage($"  更新单元格失败: {ex.Message}", InfoType.Error);
                                }
                            }
                        }
                    }

                    // 记录每行的更新状态
                    //if (rowUpdated)
                    //{
                    //    OnLogMessage($"  第 {row} 行(对应Excel第 {excelRowIndex + 1} 行)已更新");
                    //}
                }
                else
                {
                    // 清空CAD表格中多余的行
                    foreach (var colMapping in headerMapping)
                    {
                        int colIndex = colMapping.Key;
                        string header = colMapping.Value;
                        string currentValue = table.Cells[row, colIndex].GetTextString(FormatOption.IgnoreMtextFormat).Trim();

                        // 只有在当前值不为空时才清空
                        if (!string.IsNullOrEmpty(currentValue))
                        {
                            try
                            {
                                OnLogMessage($"  清空多余单元格[{row},{colIndex}]({header}): \"{currentValue}\" -> \"\"");
                                // 直接使用TextString属性清空单元格内容
                                table.Cells[row, colIndex].TextString = "";
                                updatedCells++;
                                rowUpdated = true;
                            }
                            catch (Exception ex)
                            {
                                OnLogMessage($"  清空单元格失败: {ex.Message}", InfoType.Error);
                            }
                        }
                    }

                    if (rowUpdated)
                    {
                        OnLogMessage($"  第 {row} 行已清空（Excel数据不足）");
                    }
                }
            }
            if (updatedCells > 0)
            {
                OnLogMessage($"【结果】表格更新完成，共更新了 {updatedCells} 个单元格", InfoType.Success);
            }
            else OnLogMessage($"【结果】表格未更新,未找到有变化的单元格", InfoType.Warning);

            // 应用合并单元格到CAD表格
            if (EnableMergeCells)
            {
               

                // 总是先取消所有现有的合并单元格
                UnmergeAllCells(table);
                OnLogMessage($"  已取消CAD表格中的所有现有合并单元格");

                // 如果Excel有合并单元格信息，则应用这些合并
                if (excelMergedCells.Count > 0)
                {
                    int appliedMerges = ApplyMergedCellsToCADTable(table, excelMergedCells, startRow);
                    OnLogMessage($"  应用了 {appliedMerges} 个Excel合并单元格到CAD表格");
                }
                else
                {
                    OnLogMessage($"  Excel中无合并单元格，CAD表格保持无合并状态");
                }
                         
            }
            table.DowngradeOpen();
            return true;
        }
        catch (Exception ex)
        {
            string errorMessage = $"更新表格数据时出错: {ex.Message}";
            Log.Error(errorMessage, ex);
            OnLogMessage(errorMessage, InfoType.Error);
            return false;
        }
    }

    /// <summary>
    /// 清空表格数据（保留表头）
    /// </summary>
    private bool ClearTableData(Table table)
    {
        if (table == null)
        {
            return false;
        }

        try
        {
            table.UpgradeOpen();
            int updatedCells = 0;
            // 从第二行开始清空数据（跳过表头）
            for (int row = 1; row < table.Rows.Count; row++)
            {
                bool rowUpdated = false;

                for (int col = 0; col < table.Columns.Count; col++)
                {
                    string currentValue = table.Cells[row, col].GetTextString(FormatOption.IgnoreMtextFormat).Trim();

                    // 只有在当前值不为空时才清空
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        try
                        {
                            OnLogMessage($"  即将清空单元格[{row},{col}]: \"{currentValue}\" -> \"\"");
                            table.Cells[row, col].SetValue("", ParseOption.ParseOptionNone);
                            updatedCells++;
                            rowUpdated = true;
                        }
                        catch (Exception ex)
                        {
                            OnLogMessage($"  清空单元格[{row},{col}]:失败，改用TextString: {ex.Message}", InfoType.Error);
                            table.Cells[row, col].TextString = "";
                            updatedCells++;
                            rowUpdated = true;
                        }
                    }
                }

                if (rowUpdated)
                {
                    OnLogMessage($"  第 {row} 行已清空");
                }
            }
            table.DowngradeOpen();
            OnLogMessage($"表格清空完成，共清空了 {updatedCells} 个单元格");
            return updatedCells > 0;
        }
        catch (Exception ex)
        {
            string errorMessage = $"清空表格数据时出错: {ex.Message}";
            Log.Error(errorMessage, ex);
            OnLogMessage(errorMessage, InfoType.Error);
            return false;
        }
    }
    /// <summary>
    /// 根据表格标题更新表格数据，直接按位置匹配单元格（除第一行外）
    /// 不考虑列名映射，直接使用单元格位置进行匹配
    /// </summary>
    private bool UpdateTableByTitle(Table table, DataTable excelData, int startRow = 0, int maxRowsToProcess = int.MaxValue)
    {
        if (table == null || excelData == null)
        {
            return false;
        }

        try
        {
            table.UpgradeOpen();
            // 处理合并单元格信息
            List<MergedCellInfo> excelMergedCells = new List<MergedCellInfo>();
            if (EnableMergeCells && excelData.ExtendedProperties.ContainsKey("MergedCells"))
            {
                excelMergedCells = excelData.ExtendedProperties["MergedCells"] as List<MergedCellInfo> ?? new List<MergedCellInfo>();
                OnLogMessage($"  检测到 {excelMergedCells.Count} 个Excel合并单元格信息");
            }
            // 计算实际要处理的行数（CAD表格行数减去表头，Excel可用行数，和指定的最大处理行数中的最小值）
            int cadTableDataRows = table.Rows.Count - 1; // 减去表头行
            int availableExcelRows = excelData.Rows.Count - startRow;
            int rowsToProcess = Math.Min(Math.Min(cadTableDataRows, availableExcelRows), maxRowsToProcess);

            OnLogMessage($"开始更新表格数据（按标题匹配），表格有 {table.Rows.Count} 行, {table.Columns.Count} 列");
            OnLogMessage($"Excel数据有 {excelData.Rows.Count} 行, {excelData.Columns.Count} 列，起始行为 {startRow + 1}，将处理 {rowsToProcess} 行");

            // 获取Excel数据表的最大列数
            int excelColCount = excelData.Columns.Count;

            int updatedCells = 0;
            // 从第二行开始更新数据（跳过表头）
            for (int row = 1; row <= cadTableDataRows; row++)
            {
                bool rowUpdated = false;

                // 如果这一行有对应的Excel数据
                if (row <= rowsToProcess)
                {
                    // 计算对应的Excel数据行索引
                    int excelRowIndex = startRow + (row - 1);

                    // 更新每一列
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        // 仅处理Excel列范围内的列
                        if (col < excelColCount)
                        {
                            // 获取当前CAD单元格的值
                            string currentValue = table.Cells[row, col].GetTextString(FormatOption.IgnoreMtextFormat).Trim();

                            // 获取要设置的新值
                            string newValue = "";
                            if (excelRowIndex >= 0 && excelRowIndex < excelData.Rows.Count)
                            {
                                object cellValue = excelData.Rows[excelRowIndex][col];
                                newValue = (cellValue == null || cellValue == DBNull.Value) ? "" : cellValue.ToString().Trim();
                            }

                            // 仅在值不同时更新
                            if (currentValue != newValue)
                            {
                                try
                                {
                                    // 记录要更新的单元格信息
                                    OnLogMessage($"  更新单元格[{row},{col}]: \"{currentValue}\" -> \"{newValue}\"");

                                    // 使用SetValue方法更新单元格内容
                                    table.Cells[row, col].SetValue(newValue, ParseOption.ParseOptionNone);
                                    updatedCells++;
                                    rowUpdated = true;
                                }
                                catch (Exception ex)
                                {
                                    // 如果SetValue失败，记录错误并回退到TextString
                                    OnLogMessage($"  SetValue失败，改用TextString: {ex.Message}", InfoType.Error);
                                    table.Cells[row, col].TextString = newValue;
                                    updatedCells++;
                                    rowUpdated = true;
                                }
                            }
                        }
                    }

                    // 记录每行的更新状态
                    if (rowUpdated)
                    {
                        OnLogMessage($"  第 {row} 行(对应Excel第 {excelRowIndex + 1} 行)已更新");
                    }
                }
                else
                {
                    // 清空CAD表格中多余的行
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        string currentValue = table.Cells[row, col].GetTextString(FormatOption.IgnoreMtextFormat).Trim();

                        // 只有在当前值不为空时才清空
                        if (!string.IsNullOrEmpty(currentValue))
                        {
                            try
                            {
                                OnLogMessage($"  清空多余单元格[{row},{col}]: \"{currentValue}\" -> \"\"");
                                // 使用SetValue方法清空单元格内容
                                table.Cells[row, col].SetValue("", ParseOption.ParseOptionNone);
                                updatedCells++;
                                rowUpdated = true;
                            }
                            catch (Exception ex)
                            {
                                // 如果SetValue失败，记录错误并回退到TextString
                                OnLogMessage($"  SetValue失败，改用TextString: {ex.Message}", InfoType.Error);
                                table.Cells[row, col].TextString = "";
                                updatedCells++;
                                rowUpdated = true;
                            }
                        }
                    }

                    if (rowUpdated)
                    {
                        OnLogMessage($"  第 {row} 行已清空（Excel数据不足）");
                    }
                }
            }

            OnLogMessage($"表格更新完成，共更新了 {updatedCells} 个单元格");
            // 处理合并单元格：如果启用了合并单元格功能，总是先清除现有合并，然后根据Excel数据决定是否重新合并
            if (EnableMergeCells)
            {
                
                // 总是先取消所有现有的合并单元格
                UnmergeAllCells(table);
                OnLogMessage($"  已取消CAD表格中的所有现有合并单元格");

                // 如果Excel有合并单元格信息，则应用这些合并
                if (excelMergedCells.Count > 0)
                {
                    int appliedMerges = ApplyMergedCellsToCADTable(table, excelMergedCells, startRow);
                    OnLogMessage($"  应用了 {appliedMerges} 个Excel合并单元格到CAD表格");
                }
                else
                {
                    OnLogMessage($"  Excel中无合并单元格，CAD表格保持无合并状态");
                }
               
            }
            table.DowngradeOpen();
            return true;
        }
        catch (Exception ex)
        {
            string errorMessage = $"更新表格数据时出错: {ex.Message}";
            Log.Error(errorMessage, ex);
            OnLogMessage(errorMessage, InfoType.Error);
            return false;
        }
    }
    private int ApplyMergedCellsToCADTable(Table table, List<MergedCellInfo> excelMergedCells, int dataStartRow)
    {
        int appliedCount = 0;

        try
        {
            OnLogMessage($"开始应用合并单元格，dataStartRow = {dataStartRow}");

            // 第一步：取消表格中所有现有的合并单元格
            UnmergeAllCells(table);

            foreach (var mergedCell in excelMergedCells)
            {
                // 跳过表头行的合并单元格
                if (mergedCell.StartRow == 0) continue;

                try
                {
                    int cadStartRow = mergedCell.StartRow - dataStartRow;
                    int cadEndRow = mergedCell.EndRow - dataStartRow;
                    int cadStartCol = mergedCell.StartCol;
                    int cadEndCol = mergedCell.EndCol;

                    OnLogMessage($"计算后的CAD位置: [{cadStartRow},{cadStartCol}] 到 [{cadEndRow},{cadEndCol}]");

                    // 确保范围有效
                    if (cadStartRow >= 1 && cadEndRow < table.Rows.Count &&
                        cadStartCol >= 0 && cadEndCol < table.Columns.Count &&
                        cadStartRow <= cadEndRow && cadStartCol <= cadEndCol &&
                        (cadStartRow != cadEndRow || cadStartCol != cadEndCol))
                    {
                        // 显示即将合并的单元格内容
                        string cellContent = table.Cells[cadStartRow, cadStartCol].GetTextString(FormatOption.IgnoreMtextFormat);
                        OnLogMessage($"即将合并的主单元格[{cadStartRow},{cadStartCol}]内容: '{cellContent}'");

                        // 清空合并区域的其他单元格
                        for (int row = cadStartRow; row <= cadEndRow; row++)
                        {
                            for (int col = cadStartCol; col <= cadEndCol; col++)
                            {
                                if (row == cadStartRow && col == cadStartCol)
                                {
                                    continue; // 主单元格保持内容
                                }
                                else
                                {
                                    table.Cells[row, col].TextString = "";
                                }
                            }
                        }

                        // 执行合并操作
                        CellRange range = CellRange.Create(table, cadStartRow, cadStartCol, cadEndRow, cadEndCol);
                        table.MergeCells(range);

                        OnLogMessage($"    CAD表格合并单元格成功: [{cadStartRow},{cadStartCol}] 到 [{cadEndRow},{cadEndCol}]");
                        appliedCount++;
                    }
                    else
                    {
                        OnLogMessage($"合并范围超出表格边界或无效，跳过", InfoType.Warning);
                    }
                }
                catch (Exception ex)
                {
                    OnLogMessage($"    应用单个合并单元格失败: {ex.Message}", InfoType.Warning);
                }
            }
        }
        catch (Exception ex)
        {
            OnLogMessage($"应用合并单元格到CAD表格失败: {ex.Message}", InfoType.Error);
        }

        return appliedCount;
    }

    /// <summary>
    /// 取消表格中所有现有的合并单元格
    /// </summary>
    private void UnmergeAllCells(Table table)
    {
        try
        {
            OnLogMessage("开始取消表格中的现有合并单元格");
            int unmergedCount = 0;

            // 遍历所有单元格，查找已合并的单元格
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    try
                    {
                        var cell = table.Cells[row, col];

                        // 检查单元格是否已合并
                        if (cell.IsMerged.Value)
                        {
                            // 获取合并范围
                            var mergeRange = cell.GetMergeRange();
                            if (mergeRange != null)
                            {
                                // 只在主单元格（左上角）执行取消合并
                                if (row == mergeRange.TopRow && col == mergeRange.LeftColumn)
                                {
                                    OnLogMessage($"  取消合并单元格: [{mergeRange.TopRow},{mergeRange.LeftColumn}] 到 [{mergeRange.BottomRow},{mergeRange.RightColumn}]");

                                    // 取消合并
                                    CellRange range = CellRange.Create(table, mergeRange.TopRow, mergeRange.LeftColumn,
                                                                     mergeRange.BottomRow, mergeRange.RightColumn);
                                    table.UnmergeCells(range);
                                    unmergedCount++;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OnLogMessage($"  检查/取消单元格[{row},{col}]合并时出错: {ex.Message}", InfoType.Warning);
                    }
                }
            }

            OnLogMessage($"取消合并完成，共取消了 {unmergedCount} 个合并区域");
        }
        catch (Exception ex)
        {
            OnLogMessage($"取消表格合并单元格时出错: {ex.Message}", InfoType.Error);
        }
    }
    #endregion

    #region CAD to Excel
    /// <summary>
    /// 将CAD表格数据导出到Excel
    /// </summary>
    /// <param name="mappings">映射配置列表</param>
    /// <param name="processFiles">要处理的CAD文件列表</param>
    /// <returns>是否成功</returns>
    public async Task<bool> CADToExcel(List<MappingConfig> mappings, List<string> processFiles)
    {
        try
        {
            OnLogMessage("开始处理CAD到Excel的数据导出");

            if (mappings == null || mappings.Count == 0)
            {
                OnLogMessage("映射配置为空，无法继续处理", InfoType.Error);
                return false;
            }

            // 过滤出DWG文件和Excel文件
            var dwgs = processFiles.Where(p => Path.GetExtension(p).ToLower() == ".dwg").ToList();
            var excels = processFiles.Where(p => Path.GetExtension(p).ToLower() == ".xlsx" || Path.GetExtension(p).ToLower() == ".xls")
                                    .Where(x => Path.GetFileNameWithoutExtension(x).ToLower() == "index").ToList();

            if (dwgs.Count == 0 || excels.Count == 0)
            {
                OnLogMessage(dwgs.Count == 0 ? "未找到DWG文件" : "未找到Excel文件", InfoType.Error);
                return false;
            }

            // 使用第一个Excel文件
            string excelFilePath = excels[0];
            OnLogMessage($"准备更新Excel文件: {Path.GetFileName(excelFilePath)}");

            // 第一步：收集所有表格数据，按工作表分组
            var pageTypeTableCollection = new Dictionary<string, List<TableData>>();  // PageType类型的表格集合
            var tableTitleCollection = new Dictionary<string, List<TableData>>();     // TableTitle类型的表格集合

            // 收集所有DWG文件的表格数据
            foreach (var filePath in dwgs)
            {               
                CollectTablesFromDwg(filePath, mappings, pageTypeTableCollection, tableTitleCollection);
            }

            OnLogMessage($"收集完成,PageType类型：{pageTypeTableCollection.Count}个工作表，TableTitle类型：{tableTitleCollection.Count}个工作表");

            // 第二步：处理数据并更新到Excel
            bool result = await UpdateExcelWithTables(excelFilePath, pageTypeTableCollection, tableTitleCollection);

            return result;
        }
        catch (Exception ex)
        {
            string errorMessage = $"CAD到Excel数据导出过程发生错误: {ex.Message}";
            Log.Error(errorMessage, ex);
            OnLogMessage(errorMessage, InfoType.Error);
            return false;
        }
    }

    // 修改TableData类，存储表格内容而不是表格对象
    private class TableData
    {
        public List<List<string>> TableContent { get; set; }  // 表格内容，二维数组
        public List<MergedCellInfo> MergedCells { get; set; } 
        public string FilePath { get; set; }                 // 来源文件路径
        public string TargetSheet { get; set; }              // 目标Excel工作表
        public int RowCount { get; set; }                    // 行数
        public int ColumnCount { get; set; }                 // 列数
    }
    // 在事务内复制表格数据
    private void CollectTablesFromDwg(string filePath, List<MappingConfig> mappings,
        Dictionary<string, List<TableData>> pageTypeCollection, Dictionary<string, List<TableData>> tableTitleCollection)
    {
        try
        {
            using (Database db = new Database(false, true))
            {
                db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, "");

                using (DrawingUtility util = new DrawingUtility(db, false))
                {
                    // 获取页面类型
                    string pageType = util.GetDwgSummary("PAGETYPE");                   
                    // 获取所有表格
                    var tables = util.GetAllEntity<Table>();
                    if (tables.Count == 0)
                    {
                        //OnLogMessage($"  文件中未找到表格");
                        return;
                    }
                 
                    // 首先尝试通过PageType匹配
                    bool pageTypeMatched = false;
                    if (!string.IsNullOrEmpty(pageType))
                    {
                        var pageTypeMapping = mappings.FirstOrDefault(m => m.DwgKeyword == pageType && m.Enabled);
                        if (pageTypeMapping != null && !string.IsNullOrEmpty(pageTypeMapping.ExcelSheet))
                        {
                            pageTypeMatched = true;
                            string targetSheet = pageTypeMapping.ExcelSheet;
                            OnLogMessage($"{Path.GetFileName(filePath)}|页面类型: {pageType}| 找到 {tables.Count} 个表格---> 匹配到工作表: {targetSheet}");                           
                            // 将所有表格添加到PageType集合
                            foreach (var table in tables)
                            {
                                if (table.Rows.Count > 0 && table.Columns.Count > 0)
                                {
                                    // 复制表格内容
                                    var tableContent = new List<List<string>>();
                                    var mergedCells = new List<MergedCellInfo>();
                                    for (int row = 0; row < table.Rows.Count; row++)
                                    {
                                        var rowData = new List<string>();
                                        for (int col = 0; col < table.Columns.Count; col++)
                                        {
                                            var cell = table.Cells[row, col];
                                            rowData.Add(table.Cells[row, col].GetTextString(FormatOption.IgnoreMtextFormat));
                                            // 检测合并单元格
                                            // 注意：AutoCAD的Table.Cell有IsMerged属性来检测合并状态
                                            if (cell.IsMerged.Value)
                                            {
                                                // 获取合并单元格的范围
                                                var mergeRange = cell.GetMergeRange();
                                                if (mergeRange != null)
                                                {
                                                    var mergedInfo = new MergedCellInfo
                                                    {
                                                        StartRow = mergeRange.TopRow,
                                                        StartCol = mergeRange.LeftColumn,
                                                        EndRow = mergeRange.BottomRow,
                                                        EndCol = mergeRange.RightColumn
                                                    };

                                                    // 避免重复添加相同的合并单元格信息
                                                    if (!mergedCells.Any(m => m.StartRow == mergedInfo.StartRow &&
                                                                              m.StartCol == mergedInfo.StartCol &&
                                                                              m.EndRow == mergedInfo.EndRow &&
                                                                              m.EndCol == mergedInfo.EndCol))
                                                    {
                                                        mergedCells.Add(mergedInfo);
                                                        OnLogMessage($"    检测到合并单元格: [{mergedInfo.StartRow},{mergedInfo.StartCol}] 到 [{mergedInfo.EndRow},{mergedInfo.EndCol}]");
                                                    }
                                                }
                                            }
                                        }
                                        tableContent.Add(rowData);
                                    }

                                    // 将表格添加到PageType集合
                                    if (!pageTypeCollection.ContainsKey(targetSheet))
                                        pageTypeCollection[targetSheet] = new List<TableData>();

                                    pageTypeCollection[targetSheet].Add(new TableData
                                    {
                                        TableContent = tableContent,
                                        FilePath = filePath,
                                        TargetSheet = targetSheet,
                                        RowCount = table.Rows.Count,
                                        ColumnCount = table.Columns.Count,
                                        MergedCells = mergedCells // 添加合并单元格信息
                                    });
                                }
                            }
                        }
                    }

                    // 如果PageType未匹配，则尝试通过表格标题匹配
                    if (!pageTypeMatched)
                    {
                        // 记录已经找到的表格标题，确保每种标题只处理一次
                        HashSet<string> processedTitles = new HashSet<string>();

                        for (int i = 0; i < tables.Count; i++)
                        {
                            Table table = tables[i];
                            if (table.Rows.Count == 0 || table.Columns.Count == 0)
                                continue;

                            // 获取表格标题（第一个单元格）
                            string tableTitle = table.Cells[0, 0].GetTextString(FormatOption.IgnoreMtextFormat).Trim();
                            if (string.IsNullOrEmpty(tableTitle))
                                continue;

                            // 如果已经处理过这个标题，则跳过
                            if (processedTitles.Contains(tableTitle))
                            {
                                OnLogMessage($"  表格 #{i + 1} 标题: {tableTitle} - 已经处理过相同标题的表格，跳过");
                                continue;
                            }

                            // 查找匹配的标题配置
                            var titleMapping = mappings.FirstOrDefault(m => m.CADTableTitle == tableTitle && m.Enabled);
                            if (titleMapping != null && !string.IsNullOrEmpty(titleMapping.ExcelSheet))
                            {
                                string targetSheet = titleMapping.ExcelSheet;
                                OnLogMessage($"  表格 #{i + 1} 标题: {tableTitle}--> 匹配到工作表: {targetSheet}");                               
                                // 复制表格内容
                                var tableContent = new List<List<string>>();
                                for (int row = 0; row < table.Rows.Count; row++)
                                {
                                    var rowData = new List<string>();
                                    for (int col = 0; col < table.Columns.Count; col++)
                                    {
                                        rowData.Add(table.Cells[row, col].GetTextString(FormatOption.IgnoreMtextFormat));
                                    }
                                    tableContent.Add(rowData);
                                }

                                // 将表格添加到TableTitle集合
                                if (!tableTitleCollection.ContainsKey(targetSheet))
                                    tableTitleCollection[targetSheet] = new List<TableData>();

                                tableTitleCollection[targetSheet].Add(new TableData
                                {
                                    TableContent = tableContent,
                                    FilePath = filePath,
                                    TargetSheet = targetSheet,
                                    RowCount = table.Rows.Count,
                                    ColumnCount = table.Columns.Count
                                });

                                // 标记此标题已处理
                                processedTitles.Add(tableTitle);
                            }
                        }
                    }

                    util.Commit();
                }
            }
        }
        catch (Exception ex)
        {
            OnLogMessage($"处理DWG文件时出错: {ex.Message}", InfoType.Error);
        }
    }

    /// <summary>
    /// 合并表格数据到DataTable (PageType类型)
    /// </summary>
    private DataTable MergePageTypeTables(List<TableData> tables)
    {
        if (tables == null || tables.Count == 0)
            return null;

        // 使用第一个表格创建DataTable结构
        var firstTable = tables[0];
        DataTable result = new DataTable(firstTable.TargetSheet);

        // 从第一个表格的第一行获取列名
        var firstRowData = firstTable.TableContent[0];
        for (int col = 0; col < firstTable.ColumnCount; col++)
        {
            string colName = col < firstRowData.Count ? firstRowData[col].Trim() : $"Column{col}";
            if (string.IsNullOrEmpty(colName))
                colName = $"Column{col}";

            // 确保列名唯一
            if (result.Columns.Contains(colName))
            {
                int suffix = 1;
                string uniqueName = $"{colName}_{suffix}";
                while (result.Columns.Contains(uniqueName))
                    uniqueName = $"{colName}_{++suffix}";
                colName = uniqueName;
            }

            result.Columns.Add(colName);
        }

        // 添加所有表格的数据行，第一个表格添加所有行，其余表格跳过第一行（表头）
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
        {
            var currentTable = tables[tableIndex];
            // 确定起始行：第一个表格从第0行开始（包含表头），其他表格从第1行开始（跳过表头）
            int startRow = (tableIndex == 0) ? 0 : 1;

            for (int row = startRow; row < currentTable.TableContent.Count; row++)
            {
                DataRow newRow = result.NewRow();
                var rowData = currentTable.TableContent[row];

                for (int col = 0; col < Math.Min(rowData.Count, result.Columns.Count); col++)
                {
                    newRow[col] = rowData[col];
                }
                result.Rows.Add(newRow);
            }
        }
        // 处理合并单元格信息
        if (EnableMergeCells)
        {
            List<MergedCellInfo> allMergedCells = new List<MergedCellInfo>();
            int currentRowOffset = 0;

            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
            {
                var currentTable = tables[tableIndex];

                if (currentTable.MergedCells != null && currentTable.MergedCells.Count > 0)
                {
                    foreach (var mergedCell in currentTable.MergedCells)
                    {
                        var adjustedMergedCell = new MergedCellInfo
                        {
                            StartRow = mergedCell.StartRow + currentRowOffset,
                            StartCol = mergedCell.StartCol,
                            EndRow = mergedCell.EndRow + currentRowOffset,
                            EndCol = mergedCell.EndCol
                        };
                        allMergedCells.Add(adjustedMergedCell);
                    }
                }

                // 更新行偏移量
                if (tableIndex == 0)
                {
                    currentRowOffset += currentTable.RowCount;
                }
                else
                {
                    currentRowOffset += (currentTable.RowCount - 1);
                }
            }

            if (allMergedCells.Count > 0)
            {
                result.ExtendedProperties["MergedCells"] = allMergedCells;
                OnLogMessage($"合并了 {allMergedCells.Count} 个合并单元格信息");
            }
        }
        return result;
    }

    /// <summary>
    /// 合并表格标题类型的表格
    /// </summary>
    private DataTable MergeTableTitleTables(List<TableData> tables)
    {
        if (tables == null || tables.Count == 0)
            return null;

        // 使用第一个表格创建DataTable结构
        var firstTable = tables[0];
        DataTable result = new DataTable(firstTable.TargetSheet);

        // 为所有列创建列名
        for (int col = 0; col < firstTable.ColumnCount; col++)
        {
            string colName = $"Column{col}";
            result.Columns.Add(colName);
        }

        // 添加第一个表格的所有行
        var tableContent = firstTable.TableContent;
        for (int row = 0; row < tableContent.Count; row++)
        {
            DataRow newRow = result.NewRow();
            var rowData = tableContent[row];

            for (int col = 0; col < Math.Min(rowData.Count, result.Columns.Count); col++)
            {
                newRow[col] = rowData[col];
            }
            result.Rows.Add(newRow);
        }
        // 处理合并单元格信息
        if (EnableMergeCells && firstTable.MergedCells != null && firstTable.MergedCells.Count > 0)
        {
            result.ExtendedProperties["MergedCells"] = firstTable.MergedCells;
            OnLogMessage($"存储了 {firstTable.MergedCells.Count} 个合并单元格信息");
        }

        return result;
    }

    /// <summary>
    /// 将收集的表格数据更新到Excel
    /// </summary>
    private async Task<bool> UpdateExcelWithTables(string excelFilePath,
        Dictionary<string, List<TableData>> pageTypeCollection, Dictionary<string, List<TableData>> tableTitleCollection)
    {
        int updatedSheets = 0;
        bool overallResult = true;

        // 处理PageType类型的表格
        foreach (var entry in pageTypeCollection)
        {
            string sheetName = entry.Key;
            List<TableData> tables = entry.Value;

            OnLogMessage($"【开始】处理PageType类型的工作表 '{sheetName}'，共 {tables.Count} 个表格");

            // 合并表格数据
            DataTable mergedData = MergePageTypeTables(tables);
            if (mergedData == null || mergedData.Rows.Count == 0)
            {
                OnLogMessage($"  工作表 '{sheetName}' 没有有效数据，跳过");
                continue;
            }

            // 更新Excel
            try
            {
                OnLogMessage($"  更新Excel工作表 '{sheetName}'，共 {mergedData.Rows.Count} 行");
                bool result = await excelProcessor.UpdateExcelSheet(excelFilePath, sheetName, mergedData, false,EnableMergeCells);

                if (result)
                {
                    updatedSheets++;
                    OnLogMessage($"  工作表 '{sheetName}' 更新成功",InfoType.Success);
                }
                else
                {
                    overallResult = false;
                    OnLogMessage($"  工作表 '{sheetName}' 更新失败", InfoType.Error);
                }
            }
            catch (Exception ex)
            {
                overallResult = false;
                OnLogMessage($"  更新工作表 '{sheetName}' 时出错: {ex.Message}", InfoType.Error);
            }
        }

        // 处理TableTitle类型的表格
        foreach (var entry in tableTitleCollection)
        {
            string sheetName = entry.Key;
            List<TableData> tables = entry.Value;

            OnLogMessage($"处理TableTitle类型的工作表 '{sheetName}'，共 {tables.Count} 个表格");

            // 对于TableTitle类型，我们只处理第一个匹配的表格
            if (tables.Count > 1)
            {
                OnLogMessage($"  警告: TableTitle类型找到多个匹配表格，只使用第一个");
            }

            // 合并表格数据（实际上只使用第一个表格）
            DataTable tableData = MergeTableTitleTables(tables);
            if (tableData == null || tableData.Rows.Count == 0)
            {
                OnLogMessage($"  工作表 '{sheetName}' 没有有效数据，跳过");
                continue;
            }

            // 更新Excel
            try
            {
                OnLogMessage($"  更新Excel工作表 '{sheetName}'，共 {tableData.Rows.Count} 行");
                bool result = await excelProcessor.UpdateExcelSheet(excelFilePath, sheetName, tableData, true,EnableMergeCells);

                if (result)
                {
                    updatedSheets++;
                    OnLogMessage($"  工作表 '{sheetName}' 更新成功", InfoType.Success);
                }
                else
                {
                    overallResult = false;
                    OnLogMessage($"  工作表 '{sheetName}' 更新失败", InfoType.Error);
                }
            }
            catch (Exception ex)
            {
                overallResult = false;
                OnLogMessage($"  更新工作表 '{sheetName}' 时出错: {ex.Message}", InfoType.Error);
            }
        }

        OnLogMessage($"CAD到Excel数据导出完成，成功更新 {updatedSheets} 个工作表",InfoType.Success);
        return overallResult && updatedSheets > 0;
    }
    #endregion

   
}
