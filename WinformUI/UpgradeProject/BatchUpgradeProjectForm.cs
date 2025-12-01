using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Ookii.Dialogs.WinForms; // 添加Ookii.Dialogs引用
using System.Diagnostics;
using log4net;
using WinformUI;
using System.Linq;
using MyOffice;

namespace WinformUI
{
    public partial class BatchUpgradeProjectForm:Form
    {
        // 日志记录器
        private static readonly ILog Log = LogManager.GetLogger(typeof(BatchUpgradeProjectForm));

        // 进度状态
        private int _totalFiles = 0;
        private int _processedFiles = 0;
        private int _successFiles = 0;
        private int _failedFiles = 0;

        // 定义规则类型枚举，与Designer中的下拉列表保持一致
      

        public BatchUpgradeProjectForm()
        {
            try
            {
                Log.Info("初始化批量图纸更新窗体");
                InitializeComponent();               
                InitializeUI();               
                SetupEventHandlers();
                Log.Info("批量图纸更新窗体初始化完成");
            }
            catch (Exception ex)
            {
                Log.Error("初始化批量图纸更新窗体失败", ex);
                MessageBox.Show($"初始化窗体失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region 前端变化
        private void InitializeUI()
        {
            // 设置进度条初始状态
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;

            // 设置状态标签初始文本
            lblStatus.Text = "就绪";

            // 初始化文本框状态与复选框状态一致
            txt_projectName.Enabled = chk项目名称.Checked;
            txt_ProjectNo.Enabled = chk项目编号.Checked;          
            txt_InternalCode.Enabled = chk内部编码.Checked;
            txt_versionState.Enabled = chk页眉版本状态.Checked;   
            txt_totalPage.Enabled = chk页眉总页码.Checked;

            // 设置规则类型列默认值为"包含"
            colRuleType.DefaultCellStyle.NullValue = "包含";
        }

        /// <summary>
        /// 设置事件处理程序
        /// </summary>
        private void SetupEventHandlers()
        {
            // 为项目名称设置CheckedChanged事件
            chk项目名称.CheckedChanged += (sender, e) =>{txt_projectName.Enabled = chk项目名称.Checked;};

            // 为项目编号设置CheckedChanged事件
            chk项目编号.CheckedChanged += (sender, e) =>{txt_ProjectNo.Enabled = chk项目编号.Checked;};
           

            // 为内部编码设置CheckedChanged事件
            chk内部编码.CheckedChanged += (sender, e) =>{txt_InternalCode.Enabled = chk内部编码.Checked;};

            // 为页眉版本状态设置CheckedChanged事件
            chk页眉版本状态.CheckedChanged += (sender, e) =>{txt_versionState.Enabled = chk页眉版本状态.Checked;};
          

            // 为页眉总页码设置CheckedChanged事件
            chk页眉总页码.CheckedChanged += (sender, e) =>{ txt_totalPage.Enabled = chk页眉总页码.Checked;};
        }
        private void SetControlsEnabled(bool enabled)
        {
            // 启用/禁用所有控件
            txtFolderPath.Enabled = enabled;
            btnSelectFolder.Enabled = enabled;
            chkBackupFiles.Enabled = enabled;
            dgvReplaceRules.Enabled = enabled;
            btnAddRule.Enabled = enabled;
            btnDeleteRule.Enabled = enabled;
            btnStartProcessing.Enabled = enabled;

            tabControl.Enabled = enabled;
        }
        private void btnAddRule_Click(object sender, EventArgs e)
        {
            try
            {
                // 添加新行，设置默认规则类型为"包含"
                int newRowIndex = dgvReplaceRules.Rows.Add("", "");
                dgvReplaceRules.Rows[newRowIndex].Cells["colRuleType"].Value = "包含";
                Log.Debug("添加了新的替换规则行");
            }
            catch (Exception ex)
            {
                Log.Error("添加替换规则行出错", ex);
                MessageBox.Show($"添加规则时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteRule_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvReplaceRules.SelectedRows.Count > 0)
                {
                    // 创建一个列表存储要删除的行索引
                    List<DataGridViewRow> rowsToDelete = new List<DataGridViewRow>();

                    // 添加所有选中的非新行到列表中
                    foreach (DataGridViewRow row in dgvReplaceRules.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            rowsToDelete.Add(row);
                        }
                    }

                    // 删除收集到的行
                    foreach (DataGridViewRow row in rowsToDelete)
                    {
                        dgvReplaceRules.Rows.Remove(row);
                    }

                    Log.Debug($"删除了 {rowsToDelete.Count} 行替换规则");
                }
                else if (dgvReplaceRules.SelectedCells.Count > 0)
                {
                    // 获取第一个选中单元格所在的行索引
                    int rowIndex = dgvReplaceRules.SelectedCells[0].RowIndex;

                    // 检查是否是新行
                    if (rowIndex >= 0 && !dgvReplaceRules.Rows[rowIndex].IsNewRow)
                    {
                        dgvReplaceRules.Rows.RemoveAt(rowIndex);
                        Log.Debug($"删除了第 {rowIndex + 1} 行替换规则");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("删除替换规则行出错", ex);
                MessageBox.Show($"删除规则时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 在BatchUpgradeProjectForm类中添加以下方法

        #region 规则导入导出功能


        private void btn_import_Click(object sender, EventArgs e)
        {
            try
            {
                // 创建打开对话框
                using (var openDialog = new Ookii.Dialogs.WinForms.VistaOpenFileDialog())
                {
                    openDialog.Title = "导入替换规则";
                    openDialog.Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                    openDialog.CheckFileExists = true;
                    openDialog.Multiselect = false;

                    if (openDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        // 读取文件内容
                        string[] lines = File.ReadAllLines(openDialog.FileName, Encoding.UTF8);

                        if (lines.Length <= 1) // 只有标题行或空文件
                        {
                            MessageBox.Show("文件中没有找到有效的规则。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        List<ReplaceRule> rules = new List<ReplaceRule>();

                        // 跳过标题行，从第二行开始解析
                        for (int i = 1; i < lines.Length; i++)
                        {
                            string line = lines[i].Trim();
                            if (string.IsNullOrEmpty(line))
                                continue;

                            // 解析CSV行
                            string[] fields = ParseCsvLine(line);

                            if (fields.Length >= 3)
                            {
                                RuleType ruleType;
                                // 尝试解析规则类型
                                if (!Enum.TryParse(fields[0], out ruleType))
                                {
                                    ruleType = RuleType.包含; // 默认为"包含"
                                }

                                rules.Add(new ReplaceRule
                                {
                                    Type = ruleType,
                                    FindText = fields[1],
                                    ReplaceText = fields[2]
                                });
                            }
                        }

                        if (rules.Count == 0)
                        {
                            MessageBox.Show("文件中没有找到有效的规则。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        // 确认是否替换现有规则或合并
                        DialogResult result = MessageBox.Show(
                            "是否替换当前规则？\n- 点击是替换现有规则\n- 点击否将新规则添加到现有规则后",
                            "导入规则",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.Cancel)
                            return;

                        // 如果选择替换，清除现有规则
                        if (result == DialogResult.Yes)
                        {
                            dgvReplaceRules.Rows.Clear();
                        }

                        // 添加导入的规则
                        foreach (var rule in rules)
                        {
                            // 使用Add方法创建新行
                            int rowIndex = dgvReplaceRules.Rows.Add();
                            DataGridViewRow row = dgvReplaceRules.Rows[rowIndex];

                            // 设置值
                            row.Cells["colOldText"].Value = rule.FindText;
                            row.Cells["colNewText"].Value = rule.ReplaceText;

                            // 确保规则类型正确设置 - 这是关键修改
                            DataGridViewComboBoxCell comboCell = row.Cells["colRuleType"] as DataGridViewComboBoxCell;
                            if (comboCell != null)
                            {
                                // 确保下拉列表项存在
                                if (!comboCell.Items.Contains(rule.Type.ToString()))
                                {
                                    Log.Warn($"规则类型 '{rule.Type}' 不在下拉列表中，使用默认类型'包含'");
                                    comboCell.Value = "包含";
                                }
                                else
                                {
                                    comboCell.Value = rule.Type.ToString();
                                }
                            }
                            else
                            {
                                // 如果不是下拉列表单元格，直接设置值
                                row.Cells["colRuleType"].Value = rule.Type.ToString();
                            }
                        }

                        Log.Info($"已导入 {rules.Count} 条规则从 {openDialog.FileName}");
                        MessageBox.Show($"成功导入 {rules.Count} 条规则！", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("导入规则时出错", ex);
                MessageBox.Show($"导入规则时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_export_Click(object sender, EventArgs e)
        {
            try
            {
                // 收集当前规则
                List<ReplaceRule> rules = new List<ReplaceRule>();
                foreach (DataGridViewRow row in dgvReplaceRules.Rows)
                {
                    if (!row.IsNewRow && row.Cells["colOldText"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["colOldText"].Value.ToString()))
                    {
                        string oldText = row.Cells["colOldText"].Value.ToString();
                        string newText = row.Cells["colNewText"].Value?.ToString() ?? "";
                        string ruleTypeStr = row.Cells["colRuleType"].Value?.ToString() ?? "包含";
                        RuleType ruleType = (RuleType)Enum.Parse(typeof(RuleType), ruleTypeStr);

                        rules.Add(new ReplaceRule
                        {
                            Type = ruleType,
                            FindText = oldText,
                            ReplaceText = newText
                        });
                    }
                }

                if (rules.Count == 0)
                {
                    MessageBox.Show("没有规则可供导出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 创建保存对话框
                using (var saveDialog = new Ookii.Dialogs.WinForms.VistaSaveFileDialog())
                {
                    saveDialog.Title = "导出替换规则";
                    saveDialog.Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                    saveDialog.DefaultExt = "csv";
                    saveDialog.FileName = "替换规则_" + DateTime.Now.ToString("yyyyMMdd");

                    if (saveDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        // 创建CSV内容
                        StringBuilder csvContent = new StringBuilder();

                        // 添加CSV标题行
                        csvContent.AppendLine("规则类型,查找文本,替换文本");

                        // 添加每条规则
                        foreach (var rule in rules)
                        {
                            // 处理CSV中的特殊字符：如果文本中包含逗号、引号或换行符，需要用引号包围并处理内部引号
                            string oldText = ProcessCsvField(rule.FindText);
                            string newText = ProcessCsvField(rule.ReplaceText);

                            csvContent.AppendLine($"{rule.Type},{oldText},{newText}");
                        }

                        // 保存到文件
                        File.WriteAllText(saveDialog.FileName, csvContent.ToString(), Encoding.UTF8);

                        Log.Info($"已导出 {rules.Count} 条规则到 {saveDialog.FileName}");
                        MessageBox.Show($"成功导出 {rules.Count} 条规则！", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("导出规则时出错", ex);
                MessageBox.Show($"导出规则时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
      
        /// <summary>
        /// 处理CSV字段中的特殊字符
        /// </summary>
        private string ProcessCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            // 如果字段包含逗号、引号或换行符，需要用引号包围
            bool needQuotes = field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r");

            if (needQuotes)
            {
                // 将字段中的引号替换为两个引号（CSV转义规则）
                field = field.Replace("\"", "\"\"");
                // 用引号包围整个字段
                return $"\"{field}\"";
            }

            return field;
        }


        /// <summary>
        /// 解析CSV行，处理引号和逗号等特殊情况
        /// </summary>
        private string[] ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            int position = 0;
            bool insideQuotes = false;
            StringBuilder currentField = new StringBuilder();

            while (position < line.Length)
            {
                char currentChar = line[position];

                // 处理引号
                if (currentChar == '"')
                {
                    if (insideQuotes && position + 1 < line.Length && line[position + 1] == '"')
                    {
                        // 双引号转义，添加单个引号到字段
                        currentField.Append('"');
                        position += 2;
                    }
                    else
                    {
                        // 切换引号状态
                        insideQuotes = !insideQuotes;
                        position++;
                    }
                }
                // 处理分隔符
                else if (currentChar == ',' && !insideQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                    position++;
                }
                // 处理普通字符
                else
                {
                    currentField.Append(currentChar);
                    position++;
                }
            }

            // 添加最后一个字段
            fields.Add(currentField.ToString());

            return fields.ToArray();
        }

        #endregion
        #endregion


        #region 文件选择
        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            try
            {
                // 使用Ookii.Dialogs的VistaFolderBrowserDialog代替FolderBrowserDialog
                using (var folderDialog = new VistaFolderBrowserDialog())
                {
                    folderDialog.Description = "选择包含工程图纸的文件夹";
                    folderDialog.UseDescriptionForTitle = true; // 使用描述作为对话框标题

                    // 如果文本框中已有有效路径，则设置为初始路径
                    if (!string.IsNullOrEmpty(txtFolderPath.Text) && Directory.Exists(txtFolderPath.Text))
                    {
                        folderDialog.SelectedPath = txtFolderPath.Text;
                    }

                    if (folderDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        txtFolderPath.Text = folderDialog.SelectedPath;
                        Log.Debug($"用户选择了文件夹: {folderDialog.SelectedPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("选择文件夹出错", ex);
                MessageBox.Show($"选择文件夹时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //这里写选择一个excel表的功能
        private void btn_selectExcel_Click(object sender, EventArgs e)
        {
            try
            {
                // 使用Ookii.Dialogs的VistaOpenFileDialog来选择Excel文件
                using (var fileDialog = new Ookii.Dialogs.WinForms.VistaOpenFileDialog())
                {
                    fileDialog.Title = "选择Excel文件";
                    fileDialog.Filter = "Excel文件|*.xlsx;*.xls";
                    fileDialog.CheckFileExists = true;
                    fileDialog.Multiselect = false;

                    // 如果文本框中已有有效路径，则设置为初始目录
                    if (!string.IsNullOrEmpty(txt_excel.Text) && File.Exists(txt_excel.Text))
                    {
                        fileDialog.InitialDirectory = Path.GetDirectoryName(txt_excel.Text);
                        fileDialog.FileName = Path.GetFileName(txt_excel.Text);
                    }

                    if (fileDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        txt_excel.Text = fileDialog.FileName;                      
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("选择Excel文件出错", ex);
                MessageBox.Show($"选择Excel文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnStartProcessing_Click(object sender, EventArgs e)
        {
            try
            {
                Log.Info("开始批量处理");

                // 输入验证
                if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
                {
                    Log.Warn("未选择工程图纸文件夹");
                    MessageBox.Show("请选择工程图纸文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!Directory.Exists(txtFolderPath.Text))
                {
                    Log.Warn($"指定的文件夹不存在: {txtFolderPath.Text}");
                    MessageBox.Show("指定的文件夹不存在！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (dgvReplaceRules.Rows.Count < 1 || dgvReplaceRules.Rows[0].Cells["colOldText"].Value == null)
                {
                    Log.Warn("未添加替换规则");
                    MessageBox.Show("请至少添加一条替换规则！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 收集替换规则
                List<ReplaceRule> replaceRules = new List<ReplaceRule>();
                foreach (DataGridViewRow row in dgvReplaceRules.Rows)
                {
                    if (!row.IsNewRow && row.Cells["colOldText"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["colOldText"].Value.ToString()))
                    {
                        string oldText = row.Cells["colOldText"].Value.ToString();
                        string newText = row.Cells["colNewText"].Value?.ToString() ?? "";
                        string ruleTypeStr = row.Cells["colRuleType"].Value?.ToString() ?? "包含"; // 默认为包含
                        RuleType ruleType = (RuleType)Enum.Parse(typeof(RuleType), ruleTypeStr);

                        replaceRules.Add(new ReplaceRule
                        {
                            Type = ruleType,
                            FindText = oldText,
                            ReplaceText = newText
                        });

                        Log.Debug($"添加替换规则: {ruleTypeStr} \"{oldText}\" -> \"{newText}\"");
                    }
                }

                // 收集处理选项

                ProcessingOptions options = new ProcessingOptions
                {
                    // DWG选项
                    ProcessDwgText = chkDwgText.Checked,
                    ProcessDwgTableContent = chkDwgTableContent.Checked,
                    ProcessDwgBlockAttributes = chkDwgBlockAttributes.Checked,
                    ProcessDowngradeA = chkdowngradeA.Checked,

                    // Word选项 - 项目信息
                    项目名称修改 = chk项目名称.Checked,
                    项目名称值 = txt_projectName.Text,

                    项目编号修改 = chk项目编号.Checked,
                    项目编号值 = txt_ProjectNo.Text,

                    // Word选项 - 编码
                    内部编码修改 = chk内部编码.Checked,
                    内部编码值 = txt_InternalCode.Text,

                    外部编码修改 = chk外部编码.Checked,

                    // Word选项 - 页眉
                    页眉内部编码修改 = chk页眉内部编码.Checked,

                    页眉版本状态修改 = chk页眉版本状态.Checked,
                    页眉版本状态值 = txt_versionState.Text,

                    页眉总页码修改 = chk页眉总页码.Checked,
                    页眉总页码值 = txt_totalPage.Text,

                    页眉标题替换 = chk页眉标题.Checked,

                    // Word选项 - 标题和表格
                    标题替换 = chk标题替换.Checked,

                    修改记录重置 = chk修改记录重置.Checked,

                    编校审批表格重置 = chk编校审批表格重置.Checked,

                    // 文件名选项
                    ProcessFolderName = chkFolderName.Checked,
                    ProcessDwgFileName = chkDwgFileName.Checked,
                    ProcessWordFileName = chkWordFileName.Checked,                  
                };

                Log.Info("开始执行批量处理任务");
                Log.Debug($"处理选项: DWG文本={options.ProcessDwgText}, DWG表格={options.ProcessDwgTableContent}, DWG块属性={options.ProcessDwgBlockAttributes}");

                // 禁用UI元素
                SetControlsEnabled(false);

                // 显示等待提示
                lblStatus.Text = "处理中...";
                progressBar.Value = 0;
                string newFolder = ApplyReplaceRules(txtFolderPath.Text, replaceRules);
                // 重置计数器
                _totalFiles = 0;
                _processedFiles = 0;
                _successFiles = 0;
                _failedFiles = 0;

                // 创建日志构建器
                StringBuilder logBuilder = new StringBuilder();
                logBuilder.AppendLine($"开始处理文件夹: {txtFolderPath.Text}");
                logBuilder.AppendLine($"处理时间: {DateTime.Now}");
                logBuilder.AppendLine();

                try
                {
                    List<OuterCodeAndTitle> outerCodes = new ExcelProcessor().GetDataFromExcel<OuterCodeAndTitle>(this.txt_excel.Text);
                    // 开始处理文件
                    await ProcessFilesAsync(txtFolderPath.Text, replaceRules, outerCodes, options, logBuilder);

                    // 显示完成信息
                    string resultMessage = $"批量处理完成！\n\n成功: {_successFiles} 个文件\n失败: {_failedFiles} 个文件";                 
                    MessageBox.Show(resultMessage, "处理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    lblStatus.Text = $"处理完成 | 成功: {_successFiles} | 失败: {_failedFiles}";

                    // 4. 最后处理文件夹名称
                   
                    await ProcessFolderNamesAsync(txtFolderPath.Text, replaceRules, options, logBuilder);
                    // 保存并打开日志
                    string logPath = SaveLog(newFolder, logBuilder);
                    if (!string.IsNullOrEmpty(logPath) && MessageBox.Show(
                        "是否查看处理日志？",
                        "提示",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Process.Start("notepad.exe", logPath);
                    } 
                }
                catch (Exception ex)
                {
                    Log.Error("批量处理过程中发生错误", ex);
                    MessageBox.Show($"处理过程中发生错误：\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lblStatus.Text = "处理出错";

                    // 记录错误到日志并保存
                    logBuilder.AppendLine($"处理过程中发生错误：{ex.Message}");
                    logBuilder.AppendLine($"堆栈跟踪：{ex.StackTrace}");
                    SaveLog(newFolder, logBuilder);
                }
                finally
                {
                    // 重新启用UI元素
                    SetControlsEnabled(true);
                }
            }
            catch (Exception ex)
            {
                Log.Error("启动批量处理出错", ex);
                MessageBox.Show($"启动处理任务时出错：\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion


        // 应用替换规则到文本
        private string ApplyReplaceRules(string text, List<ReplaceRule> rules)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            string result = text;

            foreach (ReplaceRule rule in rules)
            {
                try
                {
                    // 跳过空规则
                    if (string.IsNullOrEmpty(rule.FindText))
                        continue;

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

        #region 进度与日志

        // 更新进度条和状态标签
        private void UpdateProgress()
        {
            if (_totalFiles > 0)
            {
                int progress = (int)(_processedFiles * 100.0 / _totalFiles);
                if (progress > 100) progress = 100;

                // 使用Invoke确保在UI线程上更新控件
                if (progressBar.InvokeRequired)
                {
                    progressBar.Invoke(new Action(() => {
                        progressBar.Value = progress;
                        lblStatus.Text = $"处理中... {_processedFiles}/{_totalFiles} ({progress}%)";
                    }));
                }
                else
                {
                    progressBar.Value = progress;
                    lblStatus.Text = $"处理中... {_processedFiles}/{_totalFiles} ({progress}%)";
                }
            }
        }

        // 保存日志
        private string SaveLog(string folderPath, StringBuilder logBuilder)
        {
            try
            {
                string logDir = Path.Combine(folderPath, "Logs");
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                string logPath = Path.Combine(logDir, $"处理日志_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.WriteAllText(logPath, logBuilder.ToString(), Encoding.UTF8);
                Log.Info($"处理日志已保存至: {logPath}");
                return logPath;
            }
            catch (Exception ex)
            {
                Log.Error("保存处理日志出错", ex);
                return null;
            }
        }
        #endregion

        /// <summary>
        /// 创建文件夹备份
        /// </summary>
        private string BackupFolder(string folderPath, StringBuilder logBuilder)
        {
            try
            {
                // 创建备份文件夹，路径为：原文件夹所在目录/Backup
                string parentDir = Path.GetDirectoryName(folderPath);
                string backupRootDir = Path.Combine(parentDir, "Backup");

                if (!Directory.Exists(backupRootDir)) Directory.CreateDirectory(backupRootDir);              
                // 获取源文件夹名称
                string folderName = Path.GetFileName(folderPath);

                // 创建带时间戳的备份文件夹名称，避免覆盖
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupFolderPath = Path.Combine(backupRootDir, $"{folderName}_{timestamp}");

                // 复制整个目录结构
                CopyDirectory(folderPath, backupFolderPath);

                logBuilder.AppendLine($"已备份文件夹: {folderPath} -> {backupFolderPath}");
                Log.Info($"已备份文件夹: {folderPath} -> {backupFolderPath}");

                return backupFolderPath;
            }
            catch (Exception ex)
            {
                string errorMessage = $"备份文件夹失败: {folderPath}, 错误: {ex.Message}";
                logBuilder.AppendLine(errorMessage);
                Log.Error(errorMessage, ex);
                return null;
            }
        }

        /// <summary>
        /// 递归复制目录及其所有内容
        /// </summary>
        private void CopyDirectory(string sourceDir, string destDir)
        {
            // 创建目标目录
            Directory.CreateDirectory(destDir);

            // 复制所有文件
            foreach (string filePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(filePath);
                string destFilePath = Path.Combine(destDir, fileName);
                File.Copy(filePath, destFilePath, true);
            }

            // 递归复制所有子目录
            foreach (string subDirPath in Directory.GetDirectories(sourceDir))
            {
                string subDirName = Path.GetFileName(subDirPath);
                string destSubDir = Path.Combine(destDir, subDirName);
                CopyDirectory(subDirPath, destSubDir);
            }
        }

        /// <summary>
        /// 处理文件重命名
        /// </summary>
        private async Task<Dictionary<string, string>> RenameFilesAsync(List<string> filesToRename, List<ReplaceRule> replaceRules, ProcessingOptions options, StringBuilder logBuilder)
        {
            return await Task.Run(() =>
            {
                int renamedFiles = 0;
                Dictionary<string, string> oldToNewPaths = new Dictionary<string, string>();

                logBuilder.AppendLine("---开始修改文件名:");
                Log.Info("开始修改文件名");

                foreach (string filePath in filesToRename)
                {
                    try
                    {
                        string extension = Path.GetExtension(filePath).ToLower();
                        string oldFileName = Path.GetFileNameWithoutExtension(filePath);
                        string newFileName = ApplyReplaceRules(oldFileName, replaceRules);

                        if (oldFileName != newFileName)
                        {
                            string directory = Path.GetDirectoryName(filePath);
                            string newPath = Path.Combine(directory, newFileName + extension);
                                                 

                            // 如果目标文件已存在，先删除
                            if (File.Exists(newPath)) File.Delete(newPath);                          

                            // 重命名文件
                            File.Move(filePath, newPath);
                            renamedFiles++;

                            string renameMessage = $"已重命名: {Path.GetFileName(filePath)} -> {Path.GetFileName(newPath)}";
                            logBuilder.AppendLine(renameMessage);                           

                            // 记录旧路径到新路径的映射
                            oldToNewPaths[filePath] = newPath;
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMessage = $"处理文件名出错: {filePath}, 错误: {ex.Message}";
                        logBuilder.AppendLine(errorMessage);
                        Log.Error(errorMessage, ex);
                        _failedFiles++;
                    }

                    // 更新进度
                    _processedFiles++;
                    UpdateProgress();
                }

                logBuilder.AppendLine($"【成功】文件名修改完成，共重命名 {renamedFiles} 个文件");
                logBuilder.AppendLine();
                Log.Info($"文件名处理完成，共重命名 {renamedFiles} 个文件");

                return oldToNewPaths;
            });
        }

        /// <summary>
        /// 处理Word文件内容
        /// </summary>
        private async Task ProcessWordFilesAsync(List<string> wordFiles, List<ReplaceRule> replaceRules, List<OuterCodeAndTitle> excelData, ProcessingOptions options, StringBuilder logBuilder)
        {
            WordProcessor wordProcessor = new WordProcessor(logBuilder);
            logBuilder.AppendLine("---开始处理Word文件内容:");
            Log.Info("开始处理Word文件内容");

            foreach (string filePath in wordFiles)
            {
                try
                {
                    Log.Info($"--处理Word文件: {Path.GetFileName(filePath)}");
                    bool success = await wordProcessor.ProcessWordFile(filePath, replaceRules, excelData, options);

                    if (success)
                    {
                        _successFiles++;
                        Log.Info($"【成功】Word文件处理成功: {Path.GetFileName(filePath)}");
                    }
                    else
                    {
                        _failedFiles++;
                        Log.Warn($"【失败】Word文件处理失败: {Path.GetFileName(filePath)}");
                    }
                }
                catch (Exception ex)
                {
                    _failedFiles++;
                    string errorMessage = $"【失败】处理Word文件内容失败: {Path.GetFileName(filePath)}, 错误: {ex.Message}";
                    logBuilder.AppendLine(errorMessage);
                    Log.Error(errorMessage, ex);
                }

                // 更新进度
                _processedFiles++;
                UpdateProgress();
            }

            logBuilder.AppendLine($"Word文件处理完成，成功: {_successFiles}，失败: {_failedFiles}");
            logBuilder.AppendLine();
        }

        /// <summary>
        /// 处理DWG文件内容
        /// </summary>
        private async Task ProcessDwgFilesAsync(List<string> dwgFiles, List<ReplaceRule> replaceRules, ProcessingOptions options, StringBuilder logBuilder)
        {
            DwgProcessor dwgProcessor = new DwgProcessor(logBuilder);
            logBuilder.AppendLine("---开始处理DWG文件内容:");
            Log.Info("开始处理DWG文件内容");

            foreach (string filePath in dwgFiles)
            {
                try
                {
                    Log.Info($"--处理DWG文件: {Path.GetFileName(filePath)}");
                    bool success = await dwgProcessor.ProcessDwgFile(filePath, replaceRules, options,logBuilder);

                    if (success)
                    {
                        _successFiles++;
                        Log.Info($"【成功】DWG文件处理成功: {Path.GetFileName(filePath)}");
                    }
                    else
                    {
                        _failedFiles++;
                        Log.Warn($"【失败】DWG文件处理失败: {Path.GetFileName(filePath)}");
                    }
                }
                catch (Exception ex)
                {
                    _failedFiles++;
                    string errorMessage = $"【失败】处理DWG文件内容失败: {Path.GetFileName(filePath)}, 错误: {ex.Message}";
                    logBuilder.AppendLine(errorMessage);
                    Log.Error(errorMessage, ex);
                }

                // 更新进度
                _processedFiles++;
                UpdateProgress();
            }

            logBuilder.AppendLine($"DWG文件处理完成，成功: {_successFiles}，失败: {_failedFiles}");
            logBuilder.AppendLine();
        }

        /// <summary>
        /// 处理文件夹名称
        /// </summary>
        /// <summary>
        /// 处理文件夹名称（包括根文件夹）
        /// </summary>
        private async Task<string> ProcessFolderNamesAsync(string rootFolderPath, List<ReplaceRule> replaceRules, ProcessingOptions options, StringBuilder logBuilder)
        {
            return await Task.Run(() =>
            {
                int renamedFolders = 0, failedFolders = 0;
                string updatedRootPath = rootFolderPath;
                logBuilder.AppendLine("---开始修改文件夹名称:");

                // 从最深层文件夹开始处理
                List<string> allSubFolders = [.. Directory.GetDirectories(rootFolderPath, "*", SearchOption.AllDirectories)
            .OrderByDescending(f => f.Count(c => c == Path.DirectorySeparatorChar))];
                logBuilder.AppendLine($"找到 {allSubFolders.Count} 个子文件夹");

                // 处理所有子文件夹
                foreach (string folderPath in allSubFolders)
                {
                    try
                    {
                        string parentDir = Path.GetDirectoryName(folderPath);
                        string oldName = Path.GetFileName(folderPath);
                        string newName = ApplyReplaceRules(oldName, replaceRules);

                        if (oldName != newName)
                        {
                            string newPath = Path.Combine(parentDir, newName);
                            if (Directory.Exists(newPath))
                            {
                                logBuilder.AppendLine($"无法重命名文件夹，目标已存在: {folderPath} -> {newPath}");
                                failedFolders++;
                                continue;
                            }

                            Directory.Move(folderPath, newPath);
                            renamedFolders++;
                            logBuilder.AppendLine($"已重命名文件夹: {oldName} -> {newName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failedFolders++;
                        logBuilder.AppendLine($"处理文件夹出错: {folderPath}, 错误: {ex.Message}");
                        Log.Error($"处理文件夹出错: {folderPath}, {ex.Message}", ex);
                    }
                }

                // 处理根文件夹
                try
                {
                    string parentDir = Path.GetDirectoryName(rootFolderPath);
                    if (parentDir != null) // 确保根文件夹有父目录
                    {
                        string oldName = Path.GetFileName(rootFolderPath);
                        string newName = ApplyReplaceRules(oldName, replaceRules);

                        if (oldName != newName)
                        {
                            string newPath = Path.Combine(parentDir, newName);
                            if (!Directory.Exists(newPath))
                            {
                                Directory.Move(rootFolderPath, newPath);
                                renamedFolders++;
                                logBuilder.AppendLine($"已重命名根文件夹: {oldName} -> {newName}");
                                updatedRootPath = newPath;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    logBuilder.AppendLine($"处理根文件夹出错: {rootFolderPath}, 错误: {ex.Message}");
                    Log.Error($"处理根文件夹出错: {rootFolderPath}, {ex.Message}", ex);
                }

                logBuilder.AppendLine($"文件夹名称修改完成，共重命名 {renamedFolders} 个文件夹，失败 {failedFolders} 个");
                return updatedRootPath;
            });
        }

        /// <summary>
        /// 获取需要处理的文件列表
        /// </summary>
        private List<string> GetFilesToProcess(StringBuilder log, string rootFolderPath, ProcessingOptions options)
        {
            List<string> filesToProcess = new List<string>();

            try
            {
                // 获取支持的文件扩展名
                HashSet<string> supportedExtensions = new(StringComparer.OrdinalIgnoreCase);

                // 添加用户在选项中指定的文件类型
                if (options.ProcessWordFileName)
                {
                    supportedExtensions.Add(".doc");
                    supportedExtensions.Add(".docx");
                }

                if (options.ProcessDwgFileName) supportedExtensions.Add(".dwg");

                // 如果没有选择任何文件类型，则不处理
                if (supportedExtensions.Count == 0) return filesToProcess;               

                // 获取所有文件
                string[] allFiles = Directory.GetFiles(rootFolderPath, "*.*", SearchOption.AllDirectories);

                // 过滤文件
                foreach (string filePath in allFiles)
                {
                    string extension = Path.GetExtension(filePath).ToLower();
                    // 检查是否是支持的文件类型
                    if (supportedExtensions.Contains(extension))
                    {
                        // 检查是否在备份文件夹中（跳过备份文件夹中的文件）
                        if (!filePath.Contains(Path.DirectorySeparatorChar + "Backup" + Path.DirectorySeparatorChar)) filesToProcess.Add(filePath);
                    }
                }

                log.AppendLine($"找到 {filesToProcess.Count} 个需要处理的文件");
            }
            catch (Exception ex)
            {
                Log.Error($"获取文件列表时出错: {ex.Message}", ex);
            }

            return filesToProcess;
        }

        /// <summary>
        /// 主处理方法 - 确保文件夹在最后处理
        /// </summary>
        public async Task ProcessFilesAsync(string rootFolderPath, List<ReplaceRule> replaceRules, List<OuterCodeAndTitle> excelData, ProcessingOptions options, StringBuilder logBuilder)
        {
            try
            {
                if (chkBackupFiles.Checked) BackupFolder(rootFolderPath, logBuilder);//备份文件
               
                // 先获取所有需要处理的文件
                List<string> filesToRename = GetFilesToProcess(logBuilder,rootFolderPath, options);
                List<string> wordFiles = [.. filesToRename.Where(f => Path.GetExtension(f).ToLower() == ".docx" || Path.GetExtension(f).ToLower() == ".doc")];
                List<string> dwgFiles = [.. filesToRename.Where(f => Path.GetExtension(f).ToLower() == ".dwg")];

                // 1. 处理文件重命名
                Dictionary<string, string> renamedFiles = await RenameFilesAsync(filesToRename, replaceRules, options, logBuilder);

                // 2. 处理Word文件内容 (使用重命名后的路径)
                List<string> updatedWordFiles = [.. wordFiles.Select(f => renamedFiles.ContainsKey(f) ? renamedFiles[f] : f)];
                await ProcessWordFilesAsync(updatedWordFiles, replaceRules, excelData, options, logBuilder);

                // 3. 处理DWG文件内容 (使用重命名后的路径)
                List<string> updatedDwgFiles = dwgFiles.Select(f => renamedFiles.ContainsKey(f) ? renamedFiles[f] : f).ToList();
                await ProcessDwgFilesAsync(updatedDwgFiles, replaceRules, options, logBuilder);

              

                logBuilder.AppendLine("所有处理完成!");
                Log.Info("所有处理完成!");
            }
            catch (Exception ex)
            {
                string errorMessage = $"处理过程中发生错误: {ex.Message}";
                logBuilder.AppendLine(errorMessage);
                Log.Error(errorMessage, ex);
            }
        }
    }
}