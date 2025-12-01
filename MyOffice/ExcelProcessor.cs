using log4net;
using OfficeOpenXml; // EPPlus命名空间
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyOffice
{
    public class ExcelProcessor
    {
        // 日志记录器
        private static readonly ILog Log = LogManager.GetLogger(typeof(ExcelProcessor));
        // 日志构建器，用于记录处理过程
        private StringBuilder _logBuilder;

        static ExcelProcessor()
        {
            // 设置EPPlus许可证模式为非商业用途
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        /// <summary>
        /// 从Excel文件中读取数据，只匹配与目标类型属性名相同的列
        /// </summary>
        public List<T> GetDataFromExcel<T>(string filePath, int sheetIndex = 0) where T : new()
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"未找到Excel文件: {filePath}");

            List<T> resultList = new List<T>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // 获取指定索引的工作表
                var worksheet = package.Workbook.Worksheets[sheetIndex]; // EPPlus索引从1开始

                if (worksheet == null)
                    throw new ArgumentException($"工作表索引 {sheetIndex} 不存在");

                // 获取数据范围
                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.End.Column;

                if (rowCount <= 1) // 至少需要有标题行和一行数据
                    return resultList;

                // 获取类型T的所有属性
                PropertyInfo[] properties = typeof(T).GetProperties();
                Dictionary<int, PropertyInfo> columnToProperty = new Dictionary<int, PropertyInfo>();

                // 建立列索引到属性的映射
                for (int col = 1; col <= colCount; col++) // EPPlus从1开始
                {
                    string headerName = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(headerName))
                        continue;

                    // 查找匹配的属性
                    foreach (PropertyInfo property in properties)
                    {
                        if (string.Equals(headerName, property.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            columnToProperty[col] = property;
                            break;
                        }

                        // 检查DisplayName特性
                        var displayAttr = property.GetCustomAttribute<System.ComponentModel.DisplayNameAttribute>();
                        if (displayAttr != null && string.Equals(headerName, displayAttr.DisplayName, StringComparison.OrdinalIgnoreCase))
                        {
                            columnToProperty[col] = property;
                            break;
                        }
                    }
                }

                // 处理数据行
                for (int row = 2; row <= rowCount; row++) // 从第2行开始（跳过标题）
                {
                    T item = new T();
                    bool hasData = false;

                    foreach (var mapping in columnToProperty)
                    {
                        int col = mapping.Key;
                        PropertyInfo property = mapping.Value;

                        object cellValue = worksheet.Cells[row, col].Value;
                        if (cellValue != null)
                        {
                            hasData = true;
                            SetPropertyValue(property, item, cellValue);
                        }
                    }

                    if (hasData)
                        resultList.Add(item);
                }
            }

            return resultList;
        }

        private void SetPropertyValue(PropertyInfo property, object obj, object value)
        {
            if (value == null || value == DBNull.Value)
                return;

            Type propertyType = property.PropertyType;
            Type underlyingType = Nullable.GetUnderlyingType(propertyType);
            if (underlyingType != null)
                propertyType = underlyingType;

            try
            {
                if (propertyType == typeof(string))
                    property.SetValue(obj, value.ToString().Trim());
                else if (value is IConvertible)
                    property.SetValue(obj, Convert.ChangeType(value, propertyType));
                else
                    property.SetValue(obj, value);
            }
            catch (Exception ex)
            {
                // 转换失败时记录日志
                Log.Warn($"属性 {property.Name} 转换失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新Excel工作表，支持PageType和TableTitle两种不同更新模式
        /// </summary>
        /// <param name="excelFilePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="dataTable">要更新的数据</param>
        /// <param name="preserveHeader">是否保留第一行标题行（TableTitle模式）</param>
        /// <returns>是否更新成功</returns>
        public async Task<bool> UpdateExcelSheet(string excelFilePath, string sheetName, DataTable dataTable, bool preserveHeader = false,bool enableMergeCells=false)
        {
            try
            {
                // 检查文件是否存在
                FileInfo file = new FileInfo(excelFilePath);
                if (!file.Exists)
                {
                    Log.Error($"更新Excel文件失败: 文件不存在 {excelFilePath}");
                    return false;
                }

                // 锁定文件进行操作，防止多线程访问冲突
                using (var package = new ExcelPackage(file))
                {
                    // 检查工作表是否存在，不存在则创建
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    bool isNewSheet = false;

                    if (worksheet == null)
                    {
                        worksheet = package.Workbook.Worksheets.Add(sheetName);
                        isNewSheet = true;
                        Log.Info($"创建新的工作表: {sheetName}");

                        // 新建工作表时，写入列标题 - 使用DataTable的第一行作为标题
                        if (dataTable != null && dataTable.Rows.Count > 0 && dataTable.Columns.Count > 0)
                        {
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                // 第一行单元格作为标题
                                string headerText = dataTable.Rows[0][col]?.ToString() ?? dataTable.Columns[col].ColumnName;
                                worksheet.Cells[1, col + 1].Value = headerText;
                                worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                            }
                        }
                        else
                        {
                            // 如果没有数据，使用列名
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                                worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                            }
                        }
                    }


                    // 如果DataTable为空或没有列，则只清空非标题行内容
                    if (dataTable == null || dataTable.Columns.Count == 0 || dataTable.Rows.Count == 0)
                    {
                        if (worksheet.Dimension != null && worksheet.Dimension.Rows > 1)
                        {
                            int lastRow = worksheet.Dimension.Rows;
                            int lastCol = worksheet.Dimension.Columns;
                            var dataRange = worksheet.Cells[2, 1, lastRow, lastCol];
                            dataRange.Clear();
                        }

                        await package.SaveAsync();
                        Log.Info($"Excel工作表数据行已清空: {excelFilePath}, 工作表: {sheetName}");
                        return true;
                    }

                    // 获取Excel中的列名及其索引
                    Dictionary<string, int> excelColumns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int excelColCount = worksheet.Dimension?.Columns ?? 0;

                    for (int col = 1; col <= excelColCount; col++)
                    {
                        var headerValue = worksheet.Cells[1, col].Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(headerValue))
                        {
                            excelColumns[headerValue] = col;
                        }
                    }

                    // 处理列映射
                    Dictionary<int, int> columnMapping = new Dictionary<int, int>();

                    if (preserveHeader)
                    {
                        // TableTitle模式：按位置映射，不考虑列名
                        for (int dtCol = 0; dtCol < Math.Min(dataTable.Columns.Count, excelColCount); dtCol++)
                        {
                            columnMapping[dtCol] = dtCol + 1; // Excel列从1开始
                        }

                        // 如果数据表有更多列，为这些列在Excel中创建新列
                        for (int dtCol = excelColCount; dtCol < dataTable.Columns.Count; dtCol++)
                        {
                            int newExcelCol = excelColCount + (dtCol - excelColCount) + 1;
                            columnMapping[dtCol] = newExcelCol;
                        }
                    }
                    else
                    {
                        // PageType模式：按列名映射
                        // 首先尝试使用第一行的值作为列名进行映射
                        if (dataTable.Rows.Count > 0)
                        {
                            for (int dtCol = 0; dtCol < dataTable.Columns.Count; dtCol++)
                            {
                                string cellValue = dataTable.Rows[0][dtCol]?.ToString() ?? "";
                                if (!string.IsNullOrWhiteSpace(cellValue) && excelColumns.TryGetValue(cellValue, out int excelCol))
                                {
                                    columnMapping[dtCol] = excelCol;
                                }
                                else if (excelColumns.TryGetValue(dataTable.Columns[dtCol].ColumnName, out excelCol))
                                {
                                    // 回退到使用列名
                                    columnMapping[dtCol] = excelCol;
                                }
                                else
                                {
                                    // 在Excel中添加新列
                                    int newExcelCol = excelColCount + 1;
                                    excelColCount++;

                                    // 使用CAD表格的第一行单元格作为列标题
                                    string headerText = !string.IsNullOrWhiteSpace(cellValue)
                                        ? cellValue
                                        : dataTable.Columns[dtCol].ColumnName;

                                    worksheet.Cells[1, newExcelCol].Value = headerText;
                                    worksheet.Cells[1, newExcelCol].Style.Font.Bold = true;

                                    columnMapping[dtCol] = newExcelCol;
                                    if (!string.IsNullOrWhiteSpace(headerText))
                                    {
                                        excelColumns[headerText] = newExcelCol;
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 没有数据行的情况，直接使用列名
                            for (int dtCol = 0; dtCol < dataTable.Columns.Count; dtCol++)
                            {
                                string colName = dataTable.Columns[dtCol].ColumnName;
                                if (excelColumns.TryGetValue(colName, out int excelCol))
                                {
                                    columnMapping[dtCol] = excelCol;
                                }
                                else
                                {
                                    // 在Excel中添加新列
                                    int newExcelCol = excelColCount + 1;
                                    excelColCount++;

                                    worksheet.Cells[1, newExcelCol].Value = colName;
                                    worksheet.Cells[1, newExcelCol].Style.Font.Bold = true;

                                    columnMapping[dtCol] = newExcelCol;
                                    excelColumns[colName] = newExcelCol;
                                }
                            }
                        }
                    }

                    if (columnMapping.Count == 0)
                    {
                        Log.Error($"没有找到匹配的列，无法更新数据");
                        return false;
                    }

                    // 清除除标题行外的现有数据
                    if (worksheet.Dimension != null && worksheet.Dimension.Rows > 1)
                    {
                        int lastRow = worksheet.Dimension.Rows;
                        int totalColumns = Math.Max(excelColCount, dataTable.Columns.Count);
                        var dataRange = worksheet.Cells[2, 1, lastRow, totalColumns];
                        dataRange.Clear();
                    }

                    // 写入数据行 - 重要：始终从DataTable的第二行开始（跳过表头行）
                    int startDataRow = 2; // Excel行从1开始，跳过标题行
                    int dataRowsWritten = 0;

                    // 计算起始行 - 对于两种表格类型都跳过第一行（表头行）
                    int startRowIndex = 1; // DataTable的索引从0开始，所以第二行是索引1

                    for (int row = startRowIndex; row < dataTable.Rows.Count; row++)
                    {
                        int targetExcelRow = startDataRow + dataRowsWritten;
                        dataRowsWritten++;

                        // 写入数据
                        foreach (var mapping in columnMapping)
                        {
                            int dtCol = mapping.Key;
                            int excelCol = mapping.Value;

                            var cell = worksheet.Cells[targetExcelRow, excelCol];
                            var value = dataTable.Rows[row][dtCol];

                            // 保存当前单元格的所有样式属性
                            var originalStyle = cell.Style;
                            var originalNumberFormat = cell.Style.Numberformat.Format;

                            // 设置单元格值，无论是否为空
                            if (value != null && value != DBNull.Value)
                            {
                                cell.Value = value;
                            }
                            else
                            {
                                cell.Value = null; // 明确设置为null以清空单元格
                            }

                            // 如果是新增的数值类型，且没有原始格式，则设置默认格式
                            if (string.IsNullOrEmpty(originalNumberFormat))
                            {
                                if (value is DateTime)
                                    cell.Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                                else if (value is decimal || value is double || value is float)
                                    cell.Style.Numberformat.Format = "#,##0.00";
                            }
                            // 否则保持原有格式不变
                        }
                    }
                    // 处理合并单元格
                    if (enableMergeCells && dataTable.ExtendedProperties.ContainsKey("MergedCells"))
                    {
                        var mergedCells = dataTable.ExtendedProperties["MergedCells"] as List<MergedCellInfo>;
                        if (mergedCells != null && mergedCells.Count > 0)
                        {
                            ApplyMergedCells(worksheet, mergedCells, startDataRow);
                            Log.Info($"应用了 {mergedCells.Count} 个合并单元格");
                        }
                    }

                    // 保存更改
                    await package.SaveAsync();
                }

                Log.Info($"Excel更新成功: {excelFilePath}, 工作表: {sheetName}, 行数: {dataTable.Rows.Count - 1}");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error($"更新Excel文件失败: {excelFilePath}, 工作表: {sheetName}, 错误: {ex.Message}", ex);
                return false;
            }
        }

        public async Task<bool> UpdateExcelSheet(string excelFilePath, string sheetName, List<MyCell> cells)
        {
            try
            {
                _logBuilder = new StringBuilder();
                _logBuilder.AppendLine($"开始更新Excel文件: {excelFilePath}, 工作表: {sheetName}");

                // 检查文件是否存在
                FileInfo file = new FileInfo(excelFilePath);
                if (!file.Exists)
                {
                    Log.Error($"Excel文件不存在: {excelFilePath}");
                    return false;
                }

                // 检查单元格列表是否为空
                if (cells == null || cells.Count == 0)
                {
                    _logBuilder.AppendLine($"没有提供需要更新的单元格数据");
                    return false;
                }

                // 锁定文件进行操作，防止多线程访问冲突
                using (var package = new ExcelPackage(file))
                {
                    // 检查工作表是否存在，不存在则创建
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                    {
                        _logBuilder.AppendLine($"工作表 {sheetName} 不存在，创建新工作表");
                        worksheet = package.Workbook.Worksheets.Add(sheetName);
                    }
                    else
                    {
                        _logBuilder.AppendLine($"工作表 {sheetName} 已存在，更新现有工作表");
                    }

                    // 更新单元格数据
                    int updatedCount = 0;
                    foreach (var cell in cells)
                    {
                        if (cell.rowIndex <= 0 || cell.colIndex <= 0)
                        {
                            _logBuilder.AppendLine($"忽略无效单元格坐标: 行={cell.rowIndex}, 列={cell.colIndex}");
                            continue;
                        }

                        var excelCell = worksheet.Cells[cell.rowIndex, cell.colIndex];
                        excelCell.Value = cell.value;

                        // 根据数据类型设置格式
                        if (decimal.TryParse(cell.value, out decimal numValue))
                        {
                            excelCell.Style.Numberformat.Format = "#,##0.00";
                        }
                        else if (DateTime.TryParse(cell.value, out DateTime dateValue))
                        {
                            excelCell.Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        }

                        updatedCount++;
                    }

                    _logBuilder.AppendLine($"工作表 {sheetName} 已更新 {updatedCount} 个单元格");

                    // 自动调整列宽以适应内容
                    if (worksheet.Dimension != null)
                    {
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    }

                    // 保存更改
                    _logBuilder.AppendLine($"保存Excel文件更改");
                    await package.SaveAsync();
                }

                Log.Info($"Excel更新成功: {excelFilePath}, 工作表: {sheetName}, 更新了 {cells.Count} 个单元格");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error($"更新Excel文件失败: {excelFilePath}, 工作表: {sheetName}, 错误: {ex.Message}", ex);
                _logBuilder.AppendLine($"错误: {ex.Message}");
                return false;
            }
            finally
            {
                if (_logBuilder != null && _logBuilder.Length > 0)
                {
                    Log.Debug(_logBuilder.ToString());
                }
            }
        }

        /// <summary>
        /// 读取Excel工作表数据，处理空行并确保读取到最后一行
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>包含工作表数据的DataTable</returns>
        public async Task<DataTable> ReadExcelSheet(string filePath, string sheetName)
        {
            return await Task.Run(() =>
            {
                try
                {
                    _logBuilder = new StringBuilder();
                    _logBuilder.AppendLine($"开始读取Excel文件: {filePath}, 工作表: {sheetName}");

                    // 检查文件是否存在
                    if (!File.Exists(filePath))
                    {
                        Log.Error($"Excel文件不存在: {filePath}");
                        return null;
                    }

                    DataTable dt = new DataTable(sheetName);

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        // 获取指定名称的工作表
                        var worksheet = package.Workbook.Worksheets[sheetName];
                        if (worksheet == null)
                        {
                            Log.Error($"在Excel文件中未找到工作表 '{sheetName}'");
                            return null;
                        }

                        // 如果工作表为空，直接返回空表
                        if (worksheet.Dimension == null)
                        {
                            _logBuilder.AppendLine($"工作表 '{sheetName}' 为空");
                            return dt;
                        }

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        _logBuilder.AppendLine($"工作表维度: {rowCount} 行, {colCount} 列");

                        // 如果只有标题行或者为空，返回空表
                        if (rowCount <= 1)
                        {
                            _logBuilder.AppendLine($"工作表只有标题行或为空");
                            return dt;
                        }

                        // 读取标题行，创建DataTable列
                        for (int col = 1; col <= colCount; col++)
                        {
                            var headerCell = worksheet.Cells[1, col];
                            string headerText = headerCell.Value?.ToString().Trim() ?? $"Column{col}";

                            // 避免列名重复
                            if (!dt.Columns.Contains(headerText))
                            {
                                dt.Columns.Add(headerText);
                            }
                            else
                            {
                                // 如果列名重复，添加序号
                                int suffix = 1;
                                string uniqueHeaderText = $"{headerText}_{suffix}";
                                while (dt.Columns.Contains(uniqueHeaderText))
                                {
                                    suffix++;
                                    uniqueHeaderText = $"{headerText}_{suffix}";
                                }
                                dt.Columns.Add(uniqueHeaderText);
                            }
                        }

                        // 创建一个字典记录每一行的数据，保持原始行号映射
                        Dictionary<int, DataRow> allRows = new Dictionary<int, DataRow>();
                        int nonEmptyRowCount = 0;

                        // 读取所有数据行，包括空行
                        for (int row = 2; row <= rowCount; row++)
                        {
                            DataRow dataRow = dt.NewRow();
                            bool hasData = false;

                            // 读取每一列的数据
                            for (int col = 1; col <= colCount; col++)
                            {
                                if (col > dt.Columns.Count)
                                    break;

                                var cell = worksheet.Cells[row, col];
                                var value = cell.Value;

                                if (value != null)
                                {
                                    // 处理不同类型的数据
                                    if (value is DateTime dateTime)
                                    {
                                        dataRow[col - 1] = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        dataRow[col - 1] = value.ToString().Trim();
                                    }
                                    hasData = true;
                                }
                            }

                            // 记录行数据，无论是否为空行
                            allRows[row - 2] = dataRow; // 行索引从0开始
                            if (hasData)
                            {
                                nonEmptyRowCount++;
                            }
                        }

                        // 按顺序将所有行添加到DataTable中
                        for (int i = 0; i < rowCount - 1; i++) // 跳过表头行
                        {
                            if (allRows.ContainsKey(i))
                            {
                                dt.Rows.Add(allRows[i]);
                            }
                        }

                        _logBuilder.AppendLine($"从工作表 '{sheetName}' 读取了 {dt.Rows.Count} 行数据 (其中非空行: {nonEmptyRowCount} 行)");
                    }

                    if (_logBuilder != null && _logBuilder.Length > 0)
                    {
                        Log.Debug(_logBuilder.ToString());
                    }

                    return dt;
                }
                catch (Exception ex)
                {
                    Log.Error($"读取Excel工作表 '{sheetName}' 时出错: {ex.Message}", ex);
                    throw;
                }
            });
        }

        /// <summary>
        /// 读取Excel单元格数据（专用于CA002表格），从第4行开始读取所有单元格
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>单元格数据列表</returns>
        public async Task<List<MyCell>> ReadExcelCells(string filePath, string sheetName)
        {
            return await Task.Run(() =>
            {
                try
                {
                    _logBuilder = new StringBuilder();
                    _logBuilder.AppendLine($"开始读取Excel文件单元格: {filePath}, 工作表: {sheetName}");

                    // 检查文件是否存在
                    if (!File.Exists(filePath))
                    {
                        Log.Error($"Excel文件不存在: {filePath}");
                        return null;
                    }

                    List<MyCell> cells = new List<MyCell>();

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        // 获取指定名称的工作表
                        var worksheet = package.Workbook.Worksheets[sheetName];
                        if (worksheet == null)
                        {
                            Log.Error($"在Excel文件中未找到工作表 '{sheetName}'");
                            return null;
                        }

                        // 如果工作表为空，直接返回空列表
                        if (worksheet.Dimension == null)
                        {
                            _logBuilder.AppendLine($"工作表 '{sheetName}' 为空");
                            return cells;
                        }

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;
                        _logBuilder.AppendLine($"工作表维度: {rowCount} 行, {colCount} 列");

                        // 定义起始行（从第4行开始）
                        int startRow = 4;

                        // 遍历所有单元格，从第4行开始
                        for (int row = startRow; row <= rowCount; row++)
                        {
                            bool rowHasData = false;

                            for (int col = 1; col <= colCount; col++)
                            {
                                var cell = worksheet.Cells[row, col];
                                string cellValue = "";

                                // 处理单元格值，确保即使为null也创建单元格对象
                                if (cell.Value != null)
                                {
                                    var value = cell.Value;

                                    // 处理不同类型的数据
                                    if (value is DateTime dateTime)
                                    {
                                        cellValue = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    else
                                    {
                                        cellValue = value.ToString().Trim();
                                    }

                                    if (!string.IsNullOrWhiteSpace(cellValue))
                                    {
                                        rowHasData = true;
                                    }
                                }
                                // 单元格为空的情况 - 确保仍然添加到列表中

                                // 添加单元格（无论是否为空值，保留行列结构）
                                cells.Add(new MyCell
                                {
                                    rowIndex = row,
                                    colIndex = col,
                                    value = cellValue  // 空单元格将被表示为空字符串""
                                });

                                // 记录空单元格
                                if (string.IsNullOrWhiteSpace(cellValue))
                                {
                                    _logBuilder.AppendLine($"  空单元格 [{row},{col}] 已添加到列表");
                                }
                            }

                            if (rowHasData)
                            {
                                _logBuilder.AppendLine($"  行 {row} 包含数据");
                            }
                            else
                            {
                                _logBuilder.AppendLine($"  行 {row} 仅包含空单元格");
                            }
                        }

                        _logBuilder.AppendLine($"从工作表 '{sheetName}' 读取了 {cells.Count} 个单元格数据（从第 {startRow} 行开始）");
                        _logBuilder.AppendLine($"包括空单元格和非空单元格");
                    }

                    if (_logBuilder != null && _logBuilder.Length > 0)
                    {
                        Log.Debug(_logBuilder.ToString());
                    }

                    return cells;
                }
                catch (Exception ex)
                {
                    Log.Error($"读取Excel单元格数据 '{sheetName}' 时出错: {ex.Message}", ex);
                    throw;
                }
            });
        }

        /// <summary>
        /// 下载Excel模板并保存
        /// </summary>
        public async Task<bool> DownloadTemplateAsync(string savePath)
        {
            try
            {
                // 使用 EPPlus 创建一个 Excel 文件
                var fileInfo = new FileInfo(savePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    // 创建一个名为 "Write2Drawing" 的工作表
                    var worksheet = package.Workbook.Worksheets.Add("Write2Drawing");

                    // 设置第一行的标题
                    worksheet.Cells[1, 1].Value = "图纸名称";
                    worksheet.Cells[1, 2].Value = "标记_名称";
                    worksheet.Cells[1, 3].Value = "标记_内容";

                    // 设置列宽，确保标题可见
                    worksheet.Column(1).Width = 20;
                    worksheet.Column(2).Width = 20;
                    worksheet.Column(3).Width = 20;

                    // 保存 Excel 文件
                    await package.SaveAsync();
                }

                Log.Info($"模板已保存到: {savePath}");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error($"导出模板时出错: {ex.Message}", ex);
                return false;
            }
        }

        // 在ExcelProcessor类中添加读取合并单元格的方法
        /// <summary>
        /// 读取Excel工作表的合并单元格信息
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="sheetName">工作表名称</param>
        /// <returns>合并单元格信息列表</returns>
        public async Task<List<MergedCellInfo>> ReadExcelMergedCells(string filePath, string sheetName)
        {
            return await Task.Run(() =>
            {
                List<MergedCellInfo> mergedCells = new List<MergedCellInfo>();

                try
                {
                    if (!File.Exists(filePath))
                    {
                        Log.Error($"Excel文件不存在: {filePath}");
                        return mergedCells;
                    }

                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[sheetName];
                        if (worksheet == null)
                        {
                            Log.Error($"在Excel文件中未找到工作表 '{sheetName}'");
                            return mergedCells;
                        }

                        // 遍历所有合并单元格
                        foreach (var mergedRange in worksheet.MergedCells)
                        {
                            // 解析合并单元格范围，例如 "A1:B2"
                            var range = worksheet.Cells[mergedRange];

                            var mergedInfo = new MergedCellInfo
                            {
                                StartRow = range.Start.Row - 1, // 转换为0基索引
                                StartCol = range.Start.Column - 1,
                                EndRow = range.End.Row - 1,
                                EndCol = range.End.Column - 1
                            };

                            mergedCells.Add(mergedInfo);
                            Log.Debug($"检测到Excel合并单元格: [{mergedInfo.StartRow},{mergedInfo.StartCol}] 到 [{mergedInfo.EndRow},{mergedInfo.EndCol}]");
                        }

                        Log.Info($"从工作表 '{sheetName}' 读取了 {mergedCells.Count} 个合并单元格");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"读取Excel合并单元格时出错: {ex.Message}", ex);
                }

                return mergedCells;
            });
        }
        /// <summary>
        /// 应用合并单元格
        /// </summary>
        private void ApplyMergedCells(ExcelWorksheet worksheet, List<MergedCellInfo> mergedCells, int dataStartRow)
        {
            try
            {
                // EPPlus的MergedCells集合不支持Remove操作，我们需要重新创建工作表
                // 或者使用其他方法处理

                // 应用新的合并单元格
                foreach (var merged in mergedCells)
                {
                    // 跳过表头行的合并
                    if (merged.StartRow == 0) continue;

                    try
                    {
                        // 计算Excel中的实际位置
                        int excelStartRow = merged.StartRow + dataStartRow - 1;
                        int excelEndRow = merged.EndRow + dataStartRow - 1;
                        int excelStartCol = merged.StartCol + 1;
                        int excelEndCol = merged.EndCol + 1;

                        // 确保范围有效且不是单个单元格
                        if (excelStartRow > 0 && excelEndRow > 0 && excelStartCol > 0 && excelEndCol > 0 &&
                            excelStartRow <= excelEndRow && excelStartCol <= excelEndCol &&
                            (excelStartRow != excelEndRow || excelStartCol != excelEndCol))
                        {
                            // 直接使用范围对象进行合并
                            var range = worksheet.Cells[excelStartRow, excelStartCol, excelEndRow, excelEndCol];
                            range.Merge = true;

                            Log.Info($"创建合并单元格: [{excelStartRow},{excelStartCol}] 到 [{excelEndRow},{excelEndCol}]");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warn($"创建单个合并单元格失败: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error($"应用合并单元格失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 将数字转换为Excel列字母
        /// </summary>
        private string GetColumnLetter(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }
            return columnName;
        }

    }
    // 在CADTableExporter类中添加合并单元格信息类
    /// <summary>
    /// 合并单元格信息
    /// </summary>
    public class MergedCellInfo
    {
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int EndRow { get; set; }
        public int EndCol { get; set; }

        public override string ToString()
        {
            return $"[{StartRow},{StartCol}] 到 [{EndRow},{EndCol}]";
        }
    }
}