using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using MyOffice;
using Autodesk.AutoCAD.Windows;
using System.Threading.Tasks;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;

namespace WinformUI.UpdateFrameTag
{
    public partial class ReadExcelModifyCAD : Form
    {
        // 用于存储已添加的文件
        private List<string> addedFiles = new List<string>();
        private Dictionary<string, Dictionary<string, string>> drawingSpecificMarks = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        public ReadExcelModifyCAD()
        {
            InitializeComponent();
            SetupEventHandlers();
            SetupListView();
            InitializeUI();
            SetupListViewControls();
        }

        #region 前端变化
        private void InitializeUI()
        {
            // 初始化UI
            UpdateFileCountLabel();
            AppendLog("系统初始化完成，请添加DWG文件，excel配置信息，操作步骤如下：");
            AppendLog("【1】选择或拖拽dwg；");
            AppendLog("【2】全局图框修改，可以直接从界面设置，也可以在Excel中将图纸名称设置为'ALL'输入标记内容、值，然后导入Excel；");
            AppendLog("【3】单页图框修改，需要在Excel中填写单页图纸名称、标记内容、标记值。");
        }
        private void SetupEventHandlers()
        {
            // 为拖拽功能添加事件处理
            dwgFilesListView.DragEnter += DwgFilesListView_DragEnter;
            dwgFilesListView.DragDrop += DwgFilesListView_DragDrop;
            dwgFilesListView.SelectedIndexChanged += DwgFilesListView_SelectedIndexChanged;
            // 为Excel相关按钮添加事件处理
            browseExcelButton.Click += BrowseExcelButton_Click;
            downloadTemplateButton.Click += DownloadTemplateButton_Click;
            // 为Excel文本框添加拖拽功能
            excelPathTextBox.AllowDrop = true;
            excelPathTextBox.DragEnter += ExcelPathTextBox_DragEnter;
            excelPathTextBox.DragDrop += ExcelPathTextBox_DragDrop;
            // 为开始修改按钮添加事件处理
            startButton.Click += StartButton_Click;

        } 
        private void SetupListView()
        {
            // 设置ListView以支持拖放
            dwgFilesListView.AllowDrop = true;

            // 设置ListView的列
            dwgFilesListView.Columns[0].Width = 250; // 文件名
            dwgFilesListView.Columns[1].Width = 100; // 状态
            dwgFilesListView.Columns[2].Width = 350; // 文件路径
        }        
        #endregion

        #region 文件列表全选和删除功能
        private void button1_Click(object sender, EventArgs e)
        {
            // 使用Ookii.Dialogs.WinForms库的Vista风格文件夹选择对话框
            using (Ookii.Dialogs.WinForms.VistaFolderBrowserDialog folderDialog = new Ookii.Dialogs.WinForms.VistaFolderBrowserDialog())
            {
                folderDialog.Description = "选择包含DWG文件的文件夹";
                folderDialog.UseDescriptionForTitle = true; // 将描述作为对话框标题
                folderDialog.ShowNewFolderButton = true;    // 允许创建新文件夹

                if (folderDialog.ShowDialog(this) == DialogResult.OK)
                {
                    string selectedFolder = folderDialog.SelectedPath;

                    try
                    {
                        // 获取选定文件夹中的所有DWG文件
                        string[] dwgFiles = Directory.GetFiles(selectedFolder, "*.dwg", SearchOption.AllDirectories);

                        // 如果没有找到DWG文件
                        if (dwgFiles.Length == 0)
                        {
                            AppendLog($"在文件夹 '{Path.GetFileName(selectedFolder)}' 中未找到DWG文件", LogLevel.Warning);
                            MessageBox.Show($"在文件夹 '{Path.GetFileName(selectedFolder)}' 中未找到DWG文件", "未找到文件", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        int folderAddedCount = 0;
                        // 添加文件夹中的所有DWG文件
                        foreach (string file in dwgFiles)
                        {
                            // 检查文件是否已经添加过
                            if (!addedFiles.Contains(file))
                            {
                                AddFileToListView(file);
                                folderAddedCount++;
                            }
                        }

                        // 更新文件计数显示
                        UpdateFileCountLabel();

                        // 输出添加结果
                        if (folderAddedCount > 0)
                        {
                            AppendLog($"从文件夹 '{Path.GetFileName(selectedFolder)}' 添加了 {folderAddedCount} 个DWG文件", LogLevel.Success);
                        }
                        else
                        {
                            AppendLog($"文件夹 '{Path.GetFileName(selectedFolder)}' 中的DWG文件已全部添加过", LogLevel.Info);
                        }
                    }
                    catch (Exception ex)
                    {
                        // 处理访问文件夹时可能发生的异常
                        AppendLog($"访问文件夹 '{Path.GetFileName(selectedFolder)}' 时出错: {ex.Message}", LogLevel.Error);
                        MessageBox.Show($"访问文件夹时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        // 在构造函数中设置ListView的键盘和右键菜单支持
        private void SetupListViewControls()
        {
            // 设置ListView的KeyDown事件
            dwgFilesListView.KeyDown += DwgFilesListView_KeyDown;

            // 创建右键菜单
            ContextMenuStrip contextMenu = new ContextMenuStrip();

            // 添加全选菜单项
            ToolStripMenuItem selectAllMenuItem = new ToolStripMenuItem("全选");
            selectAllMenuItem.Click += (sender, e) => SelectAllFiles();
            contextMenu.Items.Add(selectAllMenuItem);

            // 添加删除选中项菜单项
            ToolStripMenuItem deleteSelectedMenuItem = new ToolStripMenuItem("删除选中项");
            deleteSelectedMenuItem.Click += (sender, e) => DeleteSelectedFiles();
            contextMenu.Items.Add(deleteSelectedMenuItem);

            // 添加清空列表菜单项
            ToolStripMenuItem clearAllMenuItem = new ToolStripMenuItem("清空列表");
            clearAllMenuItem.Click += (sender, e) => ClearAllFiles();
            contextMenu.Items.Add(clearAllMenuItem);

            // 设置ListView的右键菜单
            dwgFilesListView.ContextMenuStrip = contextMenu;
        }

        // KeyDown事件处理
        private void DwgFilesListView_KeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl+A 全选
            if (e.Control && e.KeyCode == Keys.A)
            {
                SelectAllFiles();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            // Delete 删除选中项
            else if (e.KeyCode == Keys.Delete)
            {
                DeleteSelectedFiles();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        // 全选方法
        private void SelectAllFiles()
        {
            if (dwgFilesListView.Items.Count == 0)
            {
                return;
            }

            foreach (ListViewItem item in dwgFilesListView.Items)
            {
                item.Selected = true;
            }

            // 使第一个项目获取焦点
            if (dwgFilesListView.Items.Count > 0)
            {
                dwgFilesListView.Items[0].Focused = true;
            }

            AppendLog("已全选所有文件", LogLevel.Info);
        }

        // 删除选中项
        private void DeleteSelectedFiles()
        {
            if (dwgFilesListView.SelectedItems.Count == 0)
            {
                AppendLog("请先选择要删除的文件", LogLevel.Warning);
                return;
            }

            int removeCount = dwgFilesListView.SelectedItems.Count;

            // 从后向前删除，避免索引变化
            for (int i = dwgFilesListView.SelectedItems.Count - 1; i >= 0; i--)
            {
                ListViewItem item = dwgFilesListView.SelectedItems[i];
                string filePath = item.SubItems[2].Text;
                addedFiles.Remove(filePath);
                item.Remove();
            }

            // 更新文件计数
            UpdateFileCountLabel();

            AppendLog($"已删除 {removeCount} 个文件", LogLevel.Info);
        }

        // 清空列表
        private void ClearAllFiles()
        {
            if (dwgFilesListView.Items.Count == 0)
            {
                AppendLog("文件列表已为空", LogLevel.Info);
                return;
            }

            int count = dwgFilesListView.Items.Count;
            dwgFilesListView.Items.Clear();
            addedFiles.Clear();

            // 更新文件计数
            UpdateFileCountLabel();

            AppendLog($"已清空 {count} 个文件", LogLevel.Info);
        }

        // 修改原有的RemoveFileButton_Click方法来使用公共的DeleteSelectedFiles方法
        private void RemoveFileButton_Click(object sender, EventArgs e)
        {
            DeleteSelectedFiles();
        }

        // 修改原有的ClearFilesButton_Click方法来使用公共的ClearAllFiles方法
        private void ClearFilesButton_Click(object sender, EventArgs e)
        {
            ClearAllFiles();
        }

        #endregion

        #region 拖放功能实现

        private void DwgFilesListView_DragEnter(object sender, DragEventArgs e)
        {
            // 检查是否拖拽的是文件或文件夹
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void DwgFilesListView_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                // 获取拖拽的文件和文件夹项目
                string[] items = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (items == null || items.Length == 0)
                {
                    AppendLog("未接收到任何文件或文件夹", LogLevel.Warning);
                    return;
                }

                int addedCount = 0;

                // 处理拖放的每个项目
                foreach (string item in items)
                {
                    // 检查是否是文件夹
                    if (Directory.Exists(item))
                    {
                        try
                        {
                            // 递归获取文件夹中的所有DWG文件
                            string[] dwgFiles = Directory.GetFiles(item, "*.dwg", SearchOption.AllDirectories);

                            // 如果文件夹中没有DWG文件，输出警告
                            if (dwgFiles.Length == 0)
                            {
                                AppendLog($"文件夹 {Path.GetFileName(item)} 中未找到DWG文件", LogLevel.Warning);
                                continue;
                            }

                            // 添加文件夹中的每个DWG文件
                            int folderAddedCount = 0;
                            foreach (string file in dwgFiles)
                            {
                                // 检查文件是否已经添加过
                                if (!addedFiles.Contains(file))
                                {
                                    AddFileToListView(file);
                                    folderAddedCount++;
                                }
                            }

                            // 更新总添加计数
                            addedCount += folderAddedCount;

                            // 输出添加结果
                            if (folderAddedCount > 0)
                            {
                                AppendLog($"从文件夹 {Path.GetFileName(item)} 添加了 {folderAddedCount} 个DWG文件");
                            }
                            else
                            {
                                AppendLog($"文件夹 {Path.GetFileName(item)} 中的DWG文件已全部添加过", LogLevel.Info);
                            }
                        }
                        catch (Exception ex)
                        {
                            // 处理访问文件夹时可能发生的异常
                            AppendLog($"访问文件夹 {Path.GetFileName(item)} 时出错: {ex.Message}", LogLevel.Error);
                        }
                    }
                    // 检查是否是DWG文件
                    else if (File.Exists(item) && Path.GetExtension(item).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                    {
                        // 检查文件是否已经添加过
                        if (!addedFiles.Contains(item))
                        {
                            AddFileToListView(item);
                            addedCount++;
                            AppendLog($"添加文件: {Path.GetFileName(item)}");
                        }
                        else
                        {
                            AppendLog($"文件 {Path.GetFileName(item)} 已存在，跳过", LogLevel.Info);
                        }
                    }
                    // 既不是文件夹也不是DWG文件
                    else if (File.Exists(item))
                    {
                        AppendLog($"跳过不支持的文件类型: {Path.GetFileName(item)}", LogLevel.Warning);
                    }
                    else
                    {
                        AppendLog($"无法识别的项目: {item}", LogLevel.Warning);
                    }
                }

                // 更新文件计数显示
                UpdateFileCountLabel();
            }
            catch (Exception ex)
            {
                // 处理拖放操作中的任何异常
                AppendLog($"处理拖放操作时出错: {ex.Message}", LogLevel.Error);
            }
        }
        private void AddFileToListView(string filePath)
        {
            // 创建新的ListViewItem
            ListViewItem item = new ListViewItem(Path.GetFileName(filePath));
            item.SubItems.Add("待处理");
            item.SubItems.Add(filePath);

            // 添加到ListView和文件集合
            dwgFilesListView.Items.Add(item);
            addedFiles.Add(filePath);
        }

        private void UpdateFileCountLabel()
        {
            fileCountLabel.Text = $"文件数: {dwgFilesListView.Items.Count}";
        }

        // 添加Excel文本框的拖拽处理方法
        private void ExcelPathTextBox_DragEnter(object sender, DragEventArgs e)
        {
            // 检查是否拖拽的是文件
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // 检查是否只有一个文件且是Excel文件
                if (files.Length == 1 &&
                    (Path.GetExtension(files[0]).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                     Path.GetExtension(files[0]).Equals(".xls", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void ExcelPathTextBox_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                // 获取拖拽的文件
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files.Length == 1 &&
                    (Path.GetExtension(files[0]).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                     Path.GetExtension(files[0]).Equals(".xls", StringComparison.OrdinalIgnoreCase)))
                {
                    // 设置文本框值
                    excelPathTextBox.Text = files[0];

                    // 触发文本更改事件以运行相同的Excel加载逻辑
                    // 我们可以复用BrowseExcelButton_Click中的逻辑
                    ProcessExcelFile(files[0]);
                }
                else
                {
                    if (files.Length > 1)
                    {
                        AppendLog("只能拖放一个Excel文件", LogLevel.Warning);
                    }
                    else if (files.Length == 1)
                    {
                        AppendLog($"不支持的文件类型: {Path.GetFileName(files[0])}, 请选择Excel文件(.xlsx或.xls)", LogLevel.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"处理拖放的Excel文件时出错: {ex.Message}", LogLevel.Error);
            }
        }

        #endregion

        #region Excel操作
        // 提取Excel处理逻辑到单独的方法，以便复用
        private void ProcessExcelFile(string filePath)
        {
            AppendLog($"已选择Excel文件: {filePath}");

            try
            {
                // 清空之前的图纸特定标记
                drawingSpecificMarks.Clear();

                // 使用ExcelProcessor读取选中的Excel文件
                var excelProcessor = new MyOffice.ExcelProcessor();

                // 异步读取Write2Drawing工作表的数据
                Task.Run(async () =>
                {
                    try
                    {
                        var dataTable = await excelProcessor.ReadExcelSheet(filePath, "Write2Drawing");

                        if (dataTable == null || dataTable.Rows.Count == 0)
                        {
                            // 在UI线程中显示消息
                            this.Invoke(new Action(() =>
                            {
                                AppendLog("未找到'Write2Drawing'工作表或工作表为空");
                            }));
                            return;
                        }

                        // 在UI线程中处理数据
                        this.Invoke(new Action(() =>
                        {
                            AppendLog($"成功读取'Write2Drawing'工作表，共 {dataTable.Rows.Count} 行数据");

                            // 检查列名是否符合预期
                            bool hasRequiredColumns = dataTable.Columns.Contains("图纸名称") &&
                                                        dataTable.Columns.Contains("标记_名称") &&
                                                        dataTable.Columns.Contains("标记_内容");

                            if (!hasRequiredColumns)
                            {
                                AppendLog("警告: Excel工作表缺少必要的列 (图纸名称, 标记_名称, 标记_内容)", LogLevel.Warning);
                                return;
                            }

                            // 处理所有记录
                            foreach (DataRow row in dataTable.Rows)
                            {
                                string drawingName = row["图纸名称"]?.ToString()?.Trim() ?? string.Empty;
                                string markName = row["标记_名称"]?.ToString()?.Trim() ?? string.Empty;
                                string markContent = row["标记_内容"]?.ToString()?.Trim() ?? string.Empty;

                                if (string.IsNullOrEmpty(drawingName) || string.IsNullOrEmpty(markName))
                                    continue;

                                // 如果图纸名称是"ALL"，则立即应用为全局设置
                                if (drawingName.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                                {
                                    UpdateTextBoxByMarkName(markName, markContent);
                                }
                                else
                                {
                                    // 否则，存储为图纸特定的标记，以便后续使用
                                    if (!drawingSpecificMarks.ContainsKey(drawingName))
                                    {
                                        drawingSpecificMarks[drawingName] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                    }
                                    drawingSpecificMarks[drawingName][markName] = markContent;
                                }
                            }

                            // 更新图纸列表状态，显示哪些图纸有特定标记
                            UpdateDrawingListWithMarkInfo();

                            AppendLog("已应用全局图框信息，并记录特定图纸的标记信息", LogLevel.Success);

                            // 启用RadioButton                          
                            containsMatchRadioButton.Enabled = true;
                        }));
                    }
                    catch (Exception ex)
                    {
                        // 在UI线程中显示错误消息
                        this.Invoke(new Action(() =>
                        {
                            AppendLog($"读取Excel文件时出错: {ex.Message}", LogLevel.Error);
                        }));
                    }
                });
            }
            catch (Exception ex)
            {
                AppendLog($"处理Excel文件时出错: {ex.Message}", LogLevel.Error);
            }
        }

        // 修改BrowseExcelButton_Click方法以使用新的ProcessExcelFile方法
        private void BrowseExcelButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                openFileDialog.Title = "选择Excel文件";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelPathTextBox.Text = openFileDialog.FileName;
                    ProcessExcelFile(openFileDialog.FileName);
                }
            }
        }
       

        // 更新图纸列表，显示哪些图纸有特定标记
        private void UpdateDrawingListWithMarkInfo()
        {
            foreach (ListViewItem item in dwgFilesListView.Items)
            {
                string fileName = Path.GetFileNameWithoutExtension(item.Text);
                // 检查文件名是否包含任何标记集合中的关键字
                string matchedKey = drawingSpecificMarks.Keys.FirstOrDefault(s => fileName.Contains(s));
                if (matchedKey != null)
                {
                    item.SubItems[1].Text = $"有特定标记 ({drawingSpecificMarks[matchedKey].Count} 个)";
                    item.BackColor = Color.LightGreen;  // 可视化提示
                }
            }
        }      
        // 获取指定图纸的标记信息，结合全局(ALL)和图纸特定的标记
        public Dictionary<string, string> GetDrawingMarks(string drawingName)
        {
            // 结果字典，首先包含所有全局标记
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 从当前UI控件收集所有"ALL"标记的当前值
            CollectAllTextBoxValues(result);

            // 如果存在图纸特定标记，则覆盖全局标记
            if (drawingSpecificMarks.ContainsKey(drawingName))
            {
                foreach (var mark in drawingSpecificMarks[drawingName])
                {
                    result[mark.Key] = mark.Value;  // 覆盖或添加特定值
                }
            }

            return result;
        }

        // 从所有TextBox控件收集当前值
        private void CollectAllTextBoxValues(Dictionary<string, string> values)
        {
            // 收集第一个标签页的值
            foreach (Control control in tlpTitleBlock.Controls)
            {
                if (control is TextBox textBox && !string.IsNullOrEmpty(textBox.Name))
                {
                    values[textBox.Name] = textBox.Text;
                }
            }

            // 收集第二个标签页的值
            foreach (Control control in tlpRevisions.Controls)
            {
                if (control is TextBox textBox && !string.IsNullOrEmpty(textBox.Name))
                {
                    values[textBox.Name] = textBox.Text;
                }
            }
        }

        // 辅助方法：根据标记名称更新相应的TextBox
        private void UpdateTextBoxByMarkName(string markName, string markContent)
        {
            // 使用反射查找控件并更新
            foreach (Control control in this.tlpTitleBlock.Controls)
            {
                if (control is TextBox textBox && textBox.Name.Equals(markName, StringComparison.OrdinalIgnoreCase))
                {
                    textBox.Text = markContent;
                    AppendLog($"已更新全局标记 {markName}: {markContent}");
                    return;
                }
            }

            // 检查修订版本工作表中的控件
            foreach (Control control in this.tlpRevisions.Controls)
            {
                if (control is TextBox textBox && textBox.Name.Equals(markName, StringComparison.OrdinalIgnoreCase))
                {
                    textBox.Text = markContent;
                    AppendLog($"已更新全局标记 {markName}: {markContent}");
                    return;
                }
            }

            // 如果没有找到匹配的控件，记录一个提示
            AppendLog($"未找到匹配的控件: {markName}");
        }

        // 当选择图纸列表中的文件时，可以预览特定图纸的标记
        private void DwgFilesListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dwgFilesListView.SelectedItems.Count == 0)
                return;

            string fileName = Path.GetFileNameWithoutExtension(dwgFilesListView.SelectedItems[0].Text);

            // 查找文件名包含的关键字
            string matchedKey = drawingSpecificMarks.Keys.FirstOrDefault(s => fileName.Contains(s));

            // 如果找到匹配的关键字，显示特定标记预览
            if (matchedKey != null)
            {
                AppendLog($"===== 图纸 '{fileName}' 特定标记预览 =====");
                foreach (var mark in drawingSpecificMarks[matchedKey])
                {
                    AppendLog($"  {mark.Key}: {mark.Value}");
                }
                AppendLog("===============================");
            }
        }

        private async void DownloadTemplateButton_Click(object sender, EventArgs e)
        {
            using (System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "保存Excel模板";
                saveFileDialog.FileName = "批量改图依据模板2025.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var processor = new ExcelProcessor();
                        bool isTemplateDownloaded = await processor.DownloadTemplateAsync(saveFileDialog.FileName);
                        if (isTemplateDownloaded)
                        {
                            // 模拟操作成功
                            AppendLog($"已保存模板到: {saveFileDialog.FileName}", LogLevel.Success);
                            MessageBox.Show("模板已导出成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"导出模板时出错: {ex.Message}", LogLevel.Error);
                        MessageBox.Show($"导出模板时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion

        #region 处理操作

        private async void StartButton_Click(object sender, EventArgs e)
        {
            // 确保有要处理的文件
            if (dwgFilesListView.Items.Count == 0)
            {
                MessageBox.Show("请先添加需要处理的DWG文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 禁用开始按钮，防止重复点击
            startButton.Enabled = false;
            AppendLog("开始处理CAD图框更新...");

            try
            {
                // 收集全局标记（从所有TextBox控件）
                Dictionary<string, string> globalMarks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // 收集第一个标签页的值
                foreach (Control control in tlpTitleBlock.Controls)
                {
                    if (control is TextBox textBox && !string.IsNullOrEmpty(textBox.Name) && !string.IsNullOrEmpty(textBox.Text))
                    {
                        globalMarks[textBox.Name] = textBox.Text;
                        AppendLog($"收集全局标记: {textBox.Name} = {textBox.Text}");
                    }
                }

                // 收集第二个标签页的值
                foreach (Control control in tlpRevisions.Controls)
                {
                    if (control is TextBox textBox && !string.IsNullOrEmpty(textBox.Name) && !string.IsNullOrEmpty(textBox.Text))
                    {
                        globalMarks[textBox.Name] = textBox.Text;
                        AppendLog($"收集全局标记: {textBox.Name} = {textBox.Text}");
                    }
                }

                AppendLog($"共收集 {globalMarks.Count} 个全局标记");

                if (globalMarks.Count == 0 && drawingSpecificMarks.Count == 0)
                {
                    MessageBox.Show("没有发现任何需要更新的标记数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    startButton.Enabled = true;
                    return;
                }

                // 创建CAD图框更新器
                CADFrameUpdater frameUpdater = new CADFrameUpdater();

                // 订阅日志事件
                frameUpdater.LogMessageGenerated += (s, args) =>
                {
                    if (this.InvokeRequired)
                    {
                        this.Invoke(new Action(() =>
                        {
                            if (args.IsError)
                                AppendLog($"错误: {args.Message}", LogLevel.Error);
                            else
                                AppendLog(args.Message);
                        }));
                    }
                    else
                    {
                        if (args.IsError)
                            AppendLog($"错误: {args.Message}", LogLevel.Error);
                        else
                            AppendLog(args.Message);
                    }
                };

                // 处理计数
                int successCount = 0;
                int failureCount = 0;

                // 获取所有DWG文件路径
                List<string> filePaths = new List<string>();
                foreach (ListViewItem item in dwgFilesListView.Items)
                {
                    string filePath = item.SubItems[2].Text; // 假设第三列存储完整路径
                    filePaths.Add(filePath);

                    // 更新状态为"处理中"
                    item.SubItems[1].Text = "处理中...";
                    item.BackColor = Color.LightYellow;

                    // 更新界面
                    Application.DoEvents();
                }

                // 批量处理所有文件
                var result = await frameUpdater.BatchUpdateCADFrames(filePaths, globalMarks, drawingSpecificMarks);

                // 更新UI显示处理结果
                for (int i = 0; i < dwgFilesListView.Items.Count; i++)
                {
                    var item = dwgFilesListView.Items[i];
                    string fileName = Path.GetFileNameWithoutExtension(item.Text);

                    // 检查此文件是否有特定标记
                    bool hasSpecificMarks = drawingSpecificMarks.ContainsKey(fileName);

                    // 根据处理结果更新状态
                    if (i < result.Success)
                    {
                        item.SubItems[1].Text = hasSpecificMarks ?
                            $"更新完成 (全局+特定标记)" :
                            $"更新完成 (仅全局标记)";

                        item.BackColor = Color.LightGreen;
                    }
                    else
                    {
                        item.SubItems[1].Text = "处理失败";
                        item.BackColor = Color.LightPink;
                    }
                }

                AppendLog($"CAD图框更新完成: 成功 {result.Success} 个, 失败 {result.Failure} 个",
                    result.Success > 0 ? LogLevel.Success : LogLevel.Error);

                // 显示结果对话框
                if (result.Failure == 0)
                {
                    MessageBox.Show($"所有 {result.Success} 个文件处理成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (result.Success > 0)
                {
                    MessageBox.Show($"部分文件处理成功: 成功 {result.Success} 个, 失败 {result.Failure} 个",
                        "部分完成", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show($"所有 {result.Failure} 个文件处理失败!", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendLog($"处理过程中发生错误: {ex.Message}", LogLevel.Error);
                MessageBox.Show($"处理过程中发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复按钮状态
                startButton.Enabled = true;
            }
        }

        #endregion

        #region 日志功能

        // 日志级别枚举
        private enum LogLevel
        {
            Info,
            Warning,
            Error,
            Success
        }

        /// <summary>
        /// 添加日志到日志框
        /// </summary>
        /// <param name="message">日志消息</param>
        /// <param name="level">日志级别</param>
        private void AppendLog(string message, LogLevel level = LogLevel.Info)
        {
            // 设置文本框选择位置到末尾
            logRichTextBox.SelectionStart = logRichTextBox.TextLength;
            logRichTextBox.SelectionLength = 0;

            // 添加时间戳
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logEntry = $"[{timestamp}] {message}{Environment.NewLine}";

            // 根据日志级别设置颜色
            logRichTextBox.SelectionColor = GetLogLevelColor(level);

            // 添加文本
            logRichTextBox.AppendText(logEntry);

            // 滚动到最底部
            logRichTextBox.ScrollToCaret();
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

        #endregion

       
    }
}