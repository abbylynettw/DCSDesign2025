using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using MyOffice;
using Ookii.Dialogs.WinForms;
using static WinformUI.UpdateTable.Model;

namespace WinformUI.UpdateTable
{
    public partial class BatchUpdateTableForm : Form
    {
        // 硬编码的初始配置
        private List<MappingConfig> defaultMappings =
        [
            new MappingConfig { Enabled = true,CADTableTitle="", DwgKeyword = "INDEX", ExcelSheet = "INDEX" },
            new MappingConfig { Enabled = true,CADTableTitle="", DwgKeyword = "MATERIALIST", ExcelSheet = "MATERIALIST" },
            new MappingConfig { Enabled = true,CADTableTitle="", DwgKeyword = "NAMEPLATE", ExcelSheet = "NAMEPLATE" },

            new MappingConfig { Enabled = true,CADTableTitle="", DwgKeyword = "ALARMLINE", ExcelSheet = "ALARMLINE" },
            new MappingConfig { Enabled = true,CADTableTitle="", DwgKeyword = "POWERLINE", ExcelSheet = "POWERLINE" },
            new MappingConfig { Enabled = true,CADTableTitle="", DwgKeyword = "SIGNALINE", ExcelSheet = "SIGNALINE" },

            new MappingConfig { Enabled = true,CADTableTitle="-E01", DwgKeyword = "", ExcelSheet = "CQ100" },
            new MappingConfig { Enabled = true,CADTableTitle="-E02", DwgKeyword = "", ExcelSheet = "CQ200" },
            new MappingConfig { Enabled = true,CADTableTitle="-E03", DwgKeyword = "", ExcelSheet = "CQ300" },
        ];
        public BatchUpdateTableForm()
        {
            InitializeComponent();          
            InitializeData();
          
            // 为更新按钮添加事件处理
            btnUpdate.Click += BtnUpdate_Click;
        }

        #region 前端变化

        private void InitializeData()
        {
            // 初始化数据网格
            foreach (var mapping in defaultMappings)
            {
                int rowIndex = gridMappings.Rows.Add();
                var row = gridMappings.Rows[rowIndex];
                row.Cells[0].Value = mapping.Enabled;
                row.Cells[1].Value = mapping.DwgKeyword;
                row.Cells[2].Value = mapping.CADTableTitle;
                row.Cells[3].Value = mapping.ExcelSheet;
            }
            // 初始化日志
            AppendLog("系统初始化完成，已加载默认配置。");
        }
        //浏览选择目录
        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (VistaFolderBrowserDialog folderDialog = new())
            {
                folderDialog.Description = "选择图纸集目录";
                folderDialog.UseDescriptionForTitle = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    this.txtFolderPath.Text = folderDialog.SelectedPath;
                    AppendLog($"已选择目录: {folderDialog.SelectedPath}");

                    // 可以在这里扫描目录下的文件
                    int dwgCount = Directory.GetFiles(folderDialog.SelectedPath, "*.dwg", SearchOption.AllDirectories).Length;
                    int xlsxCount = Directory.GetFiles(folderDialog.SelectedPath, "*.xlsx", SearchOption.AllDirectories).Length;

                    AppendLog($"在选择的目录中找到 {dwgCount} 个DWG文件和 {xlsxCount} 个Excel文件。");
                }
            }
        }
        //选择 全选变化
        private void checkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = checkSelectAll.Checked;           
            gridMappings.Rows.Cast<DataGridViewRow>().ToList().ForEach(row => row.Cells[0].Value = isChecked);
            AppendLog(isChecked ? "已选择所有映射项" : "已取消选择所有映射项");
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int rowIndex = gridMappings.Rows.Add();
            var row = gridMappings.Rows[rowIndex];
            row.Cells[0].Value = true;
            row.Cells[1].Value = "新关键字";
            row.Cells[2].Value = "";
            row.Cells[3].Value = "新表名";

            // 可以直接编辑选中项
            gridMappings.CurrentCell = row.Cells[1];
            gridMappings.BeginEdit(true);
            AppendLog("已添加新的映射项，请设置（DWG关键字或表格标题，任选其一）和Excel表名。");
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (gridMappings.SelectedRows.Count > 0)
            {
                string keyword = gridMappings.SelectedRows[0].Cells[1].Value?.ToString();
                string sheetName = gridMappings.SelectedRows[0].Cells[2].Value?.ToString();
                string title = gridMappings.SelectedRows[0].Cells[2].Value?.ToString();

                gridMappings.Rows.Remove(gridMappings.SelectedRows[0]);

                AppendLog($"已删除映射项：{keyword}->{title}-> {sheetName}");
            }
            else MessageBox.Show("请先选择要删除的行", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);            
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "XML文件 (*.xml)|*.xml|所有文件 (*.*)|*.*";
                openFileDialog.Title = "导入映射配置";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        AppendLog($"正在从 {openFileDialog.FileName} 导入配置...");

                        // 从XML文件导入映射配置
                        List<MappingConfig> importedMappings = ImportMappingsFromXml(openFileDialog.FileName);

                        if (importedMappings != null && importedMappings.Count > 0)
                        {
                            // 清空当前网格
                            gridMappings.Rows.Clear();

                            // 添加导入的映射到网格
                            foreach (var mapping in importedMappings)
                            {
                                int rowIndex = gridMappings.Rows.Add();
                                var row = gridMappings.Rows[rowIndex];
                                row.Cells[0].Value = mapping.Enabled;
                                row.Cells[1].Value = mapping.DwgKeyword;
                                row.Cells[2].Value = mapping.CADTableTitle;
                                row.Cells[3].Value = mapping.ExcelSheet;
                            }

                            AppendLog($"成功导入 {importedMappings.Count} 个映射配置。");
                            MessageBox.Show($"成功导入 {importedMappings.Count} 个映射配置。", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            AppendLog("导入的配置为空或格式不正确。");
                            MessageBox.Show("导入的配置为空或格式不正确。", "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"导入配置时出错: {ex.Message}");
                        MessageBox.Show($"导入配置时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private List<MappingConfig> ImportMappingsFromXml(string filePath)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(MappingCollection));

                using (FileStream fs = new(filePath, FileMode.Open))
                {
                    MappingCollection collection = (MappingCollection)serializer.Deserialize(fs);
                    return collection.Mappings;
                }
            }
            catch (Exception ex)
            {
                AppendErr($"XML反序列化错误: {ex.Message}");
                throw;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "XML文件 (*.xml)|*.xml";
                saveFileDialog.Title = "导出映射配置";
                saveFileDialog.FileName = "DWG_Excel_Mapping.xml";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        AppendLog($"正在导出配置到 {saveFileDialog.FileName}...");

                        // 从网格收集当前映射配置
                        List<MappingConfig> currentMappings = new List<MappingConfig>();

                        foreach (DataGridViewRow row in gridMappings.Rows)
                        {
                            if ((row.Cells[1].Value != null|| row.Cells[2].Value != null )&& row.Cells[3].Value != null)
                            {
                                currentMappings.Add(new MappingConfig
                                {
                                    Enabled = row.Cells[0].Value != null && (bool)row.Cells[0].Value,
                                    DwgKeyword = row.Cells[1].Value == null ? "" : row.Cells[1].Value.ToString(),
                                    CADTableTitle = row.Cells[2].Value == null ? "" : row.Cells[2].Value.ToString(),
                                    ExcelSheet = row.Cells[3].Value == null ? "" : row.Cells[3].Value.ToString(),
                                });
                            }
                        }

                        // 导出为XML
                        ExportMappingsToXml(currentMappings, saveFileDialog.FileName);

                        AppendLog($"成功导出 {currentMappings.Count} 个映射配置。");
                        MessageBox.Show($"成功导出 {currentMappings.Count} 个映射配置。", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        AppendErr($"导出配置时出错: {ex.Message}");
                        MessageBox.Show($"导出配置时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportMappingsToXml(List<MappingConfig> mappings, string filePath)
        {
            try
            {
                // 创建包装类实例
                MappingCollection collection = new() { Mappings = mappings };
                // 创建XML序列化器
                XmlSerializer serializer = new XmlSerializer(typeof(MappingCollection));

                // 写入XML文件
                using (FileStream fs = new FileStream(filePath, FileMode.Create))
                {
                    serializer.Serialize(fs, collection);
                }
            }
            catch (Exception ex)
            {
                AppendErr($"XML序列化错误: {ex.Message}");
                throw;      
            }
        }
        #endregion

        #region 日志
        /// <summary>
        /// 添加日志到日志框
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="level">日志级别</param>
        private void AppendLog(string message, LogLevel level = LogLevel.Info)
        {
            // 设置文本框选择位置到末尾
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.SelectionLength = 0;

            // 添加时间戳
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logEntry = $"[{timestamp}] {message}{Environment.NewLine}";

            // 根据日志级别设置颜色
            txtLog.SelectionColor = GetLogLevelColor(level);

            // 添加文本
            txtLog.AppendText(logEntry);

            // 滚动到最底部
            txtLog.ScrollToCaret();
        }

        /// <summary>
        /// 根据日志级别获取对应颜色
        /// </summary>
        /// <param name="level">日志级别</param>
        /// <returns>对应的颜色</returns>
        private Color GetLogLevelColor(LogLevel level)
        {
            return level switch
            {
                LogLevel.Info => Color.Black,
                LogLevel.Warning => Color.Orange,
                LogLevel.Error => Color.Red,
                LogLevel.Success => Color.Green,
                _ => Color.Black
            };
        }

        // 为了与原来的方法兼容，可以保留这些方法
        private void AppendErr(string message)
        {
            AppendLog(message, LogLevel.Error);
        }

        // 新增一个成功日志方法
        private void AppendSuccess(string message)
        {
            AppendLog(message, LogLevel.Success);
        }

        // 新增一个警告日志方法
        private void AppendWarning(string message)
        {
            AppendLog(message, LogLevel.Warning);
        }
        #endregion

        #region 开始更新      

        // Update the BtnUpdate_Click method to implement the Excel to CAD functionality
        private async void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFolderPath.Text)){
                MessageBox.Show("请先选择目录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            // 获取用户选择
            string updateDirection = radioExcelToCAD.Checked ? "Excel → CAD" : "CAD → Excel";
            bool createBackup = checkBackup.Checked;

            // 获取选中的映射项
            List<MappingConfig> selectedMappings = [];
            foreach (DataGridViewRow row in gridMappings.Rows)
            {
                bool isEnabled = row.Cells[0].Value != null && (bool)row.Cells[0].Value;
                if (isEnabled)
                {
                    selectedMappings.Add(new MappingConfig
                    {
                        Enabled = true,
                        DwgKeyword = row.Cells[1].Value?.ToString(),
                        CADTableTitle = row.Cells[2].Value?.ToString(),
                        ExcelSheet = row.Cells[3].Value?.ToString(),
                    });
                }
            }

            if (selectedMappings.Count == 0){
                MessageBox.Show("请至少选择一个映射项", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);return;}

            // 显示确认对话框
            string message = $"您确定要执行以下操作吗？\n\n" + $"更新方向: {updateDirection}\n" +
                           $"创建备份: {(createBackup ? "是" : "否")}\n" + $"选中项目: {selectedMappings.Count} 个\n\n" +
                           "此操作可能会修改文件，请确保已备份重要数据。";

            if (MessageBox.Show(message, "确认操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                // 禁用按钮，防止用户在处理过程中重复点击
                btnUpdate.Enabled = false;

                try
                {
                    // 根据选择的更新方向执行不同的操作
                    if (radioExcelToCAD.Checked)
                    {
                        // Excel到CAD的更新逻辑
                        AppendLog("开始执行Excel到CAD的更新操作...");
                        // 如果需要创建备份
                        if (createBackup)
                        {
                            AppendLog("创建DWG文件备份...");                            
                            FileUtilHelper.BackupFiles(txtFolderPath.Text,"*.dwg","DwgBackup",SearchOption.AllDirectories,true, AppendSuccess, AppendErr);
                            AppendLog("备份完成");
                        }

                        // 获取目录下的所有文件
                        Dictionary<string, List<string>> dic = FileUtilHelper.GetDrawingSets(this.txtFolderPath.Text);

                        if (dic.Count == 0)
                        {
                            AppendErr("未找到有效的图纸集，请检查选择的目录结构");
                            MessageBox.Show("未找到有效的图纸集，请检查选择的目录结构", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        AppendLog($"共找到 {dic.Count} 套图纸集");

                        // 定义成功和失败计数
                        int successCount = 0;
                        int failCount = 0;

                        // 处理每套图纸
                        int currentSet = 0;
                        foreach (var cabinet in dic)
                        {
                            currentSet++;
                            string dirName = Path.GetFileName(cabinet.Key);
                            AppendLog($"[{currentSet}/{dic.Count}] 开始处理图集: {dirName}");

                            // 创建并使用CADTableExporter
                            CADTableExporter exporter = new();

                            exporter.EnableMergeCells = checkMergeCells.Checked;
                            // 订阅日志事件
                            exporter.LogMessageGenerated += Exporter_LogMessageGenerated;

                            try
                            {                                                           
                                bool result = await exporter.ExcelToCAD(selectedMappings, cabinet.Value);

                                if (result)
                                {
                                    successCount++;
                                    AppendSuccess($"图集 {dirName} 处理成功");
                                }
                                else
                                {
                                    failCount++;
                                    AppendErr($"图集 {dirName} 处理失败");
                                }
                            }
                            finally
                            {
                                // 取消订阅事件，防止内存泄漏
                                exporter.LogMessageGenerated -= Exporter_LogMessageGenerated;
                            }
                        }

                        // 汇总处理结果
                        string resultSummary = $"Excel到CAD更新操作完成: 成功 {successCount} 套, 失败 {failCount} 套";
                        if (successCount > 0 && failCount == 0)
                        {
                            AppendSuccess(resultSummary);
                            MessageBox.Show(resultSummary, "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (successCount > 0 && failCount > 0)
                        {
                            AppendWarning(resultSummary);
                            MessageBox.Show(resultSummary + "\n请查看日志了解详细信息", "部分成功", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            AppendErr(resultSummary);
                            MessageBox.Show(resultSummary + "\n请查看日志了解详细信息", "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        // CAD到Excel的更新逻辑
                        AppendLog("开始执行CAD到Excel的更新操作...");

                        // 如果需要创建备份
                        if (createBackup)
                        {
                            AppendLog("创建备份...");                          
                            FileUtilHelper.BackupMultipleFileTypes(txtFolderPath.Text,["*.xlsx", "*.xls"],"Backup",SearchOption.TopDirectoryOnly,false,AppendSuccess,AppendErr);
                            AppendLog("备份完成");
                        }

                        // 获取目录下的所有文件
                        Dictionary<string, List<string>> dic = FileUtilHelper.GetDrawingSets(this.txtFolderPath.Text);

                        if (dic.Count == 0)
                        {
                            AppendErr("未找到有效的图纸集，请检查选择的目录结构");
                            MessageBox.Show("未找到有效的图纸集，请检查选择的目录结构", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        AppendLog($"共找到 {dic.Count} 套图纸集");

                        // 定义成功和失败计数
                        int successCount = 0;
                        int failCount = 0;

                        // 处理每套图纸
                        int currentSet = 0;
                        foreach (var cabinet in dic)
                        {
                            currentSet++;
                            string dirName = Path.GetFileName(cabinet.Key);
                            AppendLog($"[{currentSet}/{dic.Count}] 开始处理图集: {dirName}");

                            // 创建并使用CADTableExporter
                            CADTableExporter exporter = new CADTableExporter();
                            exporter.EnableMergeCells = checkMergeCells.Checked;

                            // 订阅日志事件
                            exporter.LogMessageGenerated += Exporter_LogMessageGenerated;

                            try
                            {
                                // 检查是否启用合并单元格功能
                                bool enableMergeCells = checkMergeCells.Checked;
                                if (enableMergeCells)
                                {
                                    AppendLog("已启用合并单元格功能");
                                }

                                bool result = await exporter.CADToExcel(selectedMappings, cabinet.Value);                               

                                if (result)
                                {
                                    successCount++;
                                    AppendSuccess($"图集 {dirName} 处理成功");
                                }
                                else
                                {
                                    failCount++;
                                    AppendErr($"图集 {dirName} 处理失败");
                                }
                            }
                            finally
                            {
                                // 取消订阅事件，防止内存泄漏
                                exporter.LogMessageGenerated -= Exporter_LogMessageGenerated;
                            }
                        }

                        // 汇总处理结果
                        string resultSummary = $"CAD到Excel更新操作完成: 成功 {successCount} 套, 失败 {failCount} 套";
                        if (successCount > 0 && failCount == 0)
                        {
                            AppendSuccess(resultSummary);
                            MessageBox.Show(resultSummary, "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (successCount > 0 && failCount > 0)
                        {
                            AppendLog(resultSummary);
                            MessageBox.Show(resultSummary + "\n请查看日志了解详细信息", "部分成功", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            AppendErr(resultSummary);
                            MessageBox.Show(resultSummary + "\n请查看日志了解详细信息", "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendErr($"更新操作过程中发生错误: {ex.Message}");
                    MessageBox.Show($"更新操作过程中发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // 恢复按钮状态
                    btnUpdate.Enabled = true;
                }
            }
        }

        // 处理CADTableExporter生成的日志事件
        private void Exporter_LogMessageGenerated(object sender, CADTableExporter.LogEventArgs e)
        {
            // 由于此事件可能在非UI线程上触发，需要使用Invoke确保在UI线程上更新控件
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => {
                    switch (e.InfoType)
                    {
                        case CADTableExporter.InfoType.Info:
                            AppendLog(e.Message);
                            break;
                        case CADTableExporter.InfoType.Warning:
                            AppendWarning(e.Message);
                            break;
                        case CADTableExporter.InfoType.Error:
                            AppendErr(e.Message);
                            break;
                        case CADTableExporter.InfoType.Success:
                            AppendSuccess(e.Message);
                            break;
                        default:
                            break;
                    }                   
                      
                }));
            }
            else
            {
                switch (e.InfoType)
                {
                    case CADTableExporter.InfoType.Info:
                        AppendLog(e.Message);
                        break;
                    case CADTableExporter.InfoType.Warning:
                        AppendWarning(e.Message);
                        break;
                    case CADTableExporter.InfoType.Error:
                        AppendErr(e.Message);
                        break;
                    case CADTableExporter.InfoType.Success:
                        AppendSuccess(e.Message);
                        break;
                    default:
                        break;
                }
            }

            // 确保日志立即刷新到UI
            Application.DoEvents();
        }
        #endregion
    }
}