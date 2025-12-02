using System;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyOffice;
using MyOffice.LogHelper;
using WinformUI.CADHelper;
using Ookii.Dialogs.WinForms;
using Microsoft.VisualBasic;

namespace WinformUI.BlockReplace
{
    public partial class BlockReplaceForm : Form
    {
        #region 私有字段

        private static readonly log4net.ILog log = LogManager.GetLogger<BlockReplaceForm>();
        private string _selectedDirectory = string.Empty;
        private string _sourceBlockFile = string.Empty; // 源块文件路径
        private BlockReplaceProcessor.DirectoryProcessResult _lastResult;
        private bool _isProcessing = false;

        #endregion

        #region 构造函数

        public BlockReplaceForm()
        {
            try
            {
                log.LogInfo("初始化块替换窗体");
                InitializeComponent();
                SetupEventHandlers();
                InitializeUI();
                log.LogInfo("块替换窗体初始化完成");
            }
            catch (Exception ex)
            {
                log.LogError("初始化块替换窗体失败", ex);
                MessageBox.Show($"初始化窗体失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 初始化方法

        private void SetupEventHandlers()
        {
            // 目录选择事件
            btnSelectDirectory.Click += BtnSelectDirectory_Click;
            btnClearDirectory.Click += BtnClearDirectory_Click;

            // 块配置事件
            btnBrowseSourceBlock.Click += BtnBrowseSourceBlock_Click;
            btnBrowseTargetBlock.Click += BtnBrowseTargetBlock_Click;

            // 执行控制事件
            btnPreview.Click += BtnPreview_Click;
            btnStartReplace.Click += BtnStartReplace_Click;
            btnStop.Click += BtnStop_Click;

            // 日志相关事件
            btnClearLog.Click += BtnClearLog_Click;
            btnSaveReport.Click += BtnSaveReport_Click;

            // 窗体关闭事件
            this.FormClosing += BlockReplaceForm_FormClosing;

            // 订阅进度事件
            BlockReplaceProcessor.ProgressUpdated += OnProgressUpdated;
        }

        private void InitializeUI()
        {
            // 初始化状态
            UpdateDirectoryLabel();
            
            // 初始化进度控件
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            lblProgress.Text = "就绪...";
            
            // 改进的用户提示，更友好和清晰
            AppendLog("欢迎使用块替换工具！", Color.Blue);
            AppendLog("这是一个功能强大的工具，用于替换DWG文件中的块定义。");
            AppendLog("请按照以下简单步骤操作：", Color.Green);
            AppendLog("  1. 首先选择包含要处理的DWG文件的目录");
            AppendLog("  2. 然后选择源块文件（.dwg格式）");
            AppendLog("  3. 设置目标块名称（默认与源块相同）");
            AppendLog("  4. 配置替换选项（如保留属性、位置等）");
            AppendLog("  5. 点击预览按钮查看预计替换结果");
            AppendLog("  6. 确认无误后点击开始替换按钮执行操作");
            AppendLog("\n提示：");
            AppendLog("  - 工具会递归处理所选目录及其所有子目录中的DWG文件");
            AppendLog("  - 建议在执行替换前先进行预览，以确保操作正确");
            AppendLog("  - 工具会自动创建备份文件，以防意外情况发生");
        }

        #endregion

        #region 目录选择

        private void BtnSelectDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                using (var folderDialog = new VistaFolderBrowserDialog())
                {
                    folderDialog.Description = "选择包含DWG文件的目录（将递归处理所有子目录）";
                    folderDialog.UseDescriptionForTitle = true;
                    folderDialog.ShowNewFolderButton = false;

                    if (folderDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        _selectedDirectory = folderDialog.SelectedPath;
                        
                        // 递归获取DWG文件数量
                        var dwgFiles = Directory.GetFiles(_selectedDirectory, "*.dwg", SearchOption.AllDirectories);
                        
                        if (dwgFiles.Length == 0)
                        {
                            AppendLog($"在选定目录及其子目录中未找到DWG文件", Color.Orange);
                            _selectedDirectory = string.Empty;
                            return;
                        }

                        UpdateDirectoryLabel();
                        AppendLog($"已选择目录: {_selectedDirectory}", Color.Green);
                        AppendLog($"找到 {dwgFiles.Length} 个DWG文件（包括子目录）", Color.Blue);
                        
                        // 显示部分文件列表
                        if (dwgFiles.Length <= 10)
                        {
                            AppendLog("文件列表：");
                            foreach (var file in dwgFiles)
                            {
                                AppendLog($"  - {GetRelativePath(_selectedDirectory, file)}");
                            }
                        }
                        else
                        {
                            AppendLog("部分文件列表（前10个）：");
                            for (int i = 0; i < 10; i++)
                            {
                                AppendLog($"  - {GetRelativePath(_selectedDirectory, dwgFiles[i])}");
                            }
                            AppendLog($"  ... 及其他 {dwgFiles.Length - 10} 个文件");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError("选择目录时发生错误", ex);
                AppendLog($"选择目录时发生错误: {ex.Message}", Color.Red);
            }
        }

        private void BtnClearDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                if (_isProcessing)
                {
                    MessageBox.Show("正在处理中，无法清空目录选择", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _selectedDirectory = string.Empty;
                UpdateDirectoryLabel();
                AppendLog("已清空目录选择", Color.Blue);
            }
            catch (Exception ex)
            {
                log.LogError("清空目录选择时发生错误", ex);
                AppendLog($"清空目录选择时发生错误: {ex.Message}", Color.Red);
            }
        }

        private void UpdateDirectoryLabel()
        {
            if (string.IsNullOrEmpty(_selectedDirectory))
            {
                lblDirectoryInfo.Text = "未选择目录";
                lblDirectoryInfo.ForeColor = Color.Gray;
            }
            else
            {
                var dwgCount = Directory.GetFiles(_selectedDirectory, "*.dwg", SearchOption.AllDirectories).Length;
                lblDirectoryInfo.Text = $"已选择: {_selectedDirectory} ({dwgCount} 个DWG文件)";
                lblDirectoryInfo.ForeColor = Color.DarkGreen;
            }
        }

        #endregion

        #region 块配置

        private void BtnBrowseSourceBlock_Click(object sender, EventArgs e)
        {
            try
            {
                using (var openDialog = new OpenFileDialog())
                {
                    openDialog.Filter = "AutoCAD 图形文件 (*.dwg)|*.dwg|所有文件 (*.*)|*.*";
                    openDialog.Title = "选择源块文件";
                    openDialog.CheckFileExists = true;
                    
                    if (openDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        _sourceBlockFile = openDialog.FileName;
                        AppendLog($"已选择源块文件: {_sourceBlockFile}", Color.Green);
                        
                        // 读取DWG文件中的块定义
                        AppendLog("正在读取文件中的块定义...", Color.Blue);
                        
                        List<string> blockNames = null;
                        
                        try
                        {
                            blockNames = CADHelper.DwgBlockReader.GetBlockDefinitions(_sourceBlockFile);
                        }
                        catch (Exception readEx)
                        {
                            AppendLog($"自动读取失败: {readEx.Message}", Color.Orange);
                            
                            // 根据错误类型提供不同的建议
                            string suggestion = GetReadFailureSuggestion(readEx);
                            AppendLog(suggestion, Color.Orange);
                            
                            // 提供手动输入备用方案
                            var result = MessageBox.Show(
                                $"自动读取块定义失败。\n\n错误信息：{readEx.Message}\n\n{suggestion}\n\n" +
                                "是否要手动输入块名称？\n\n" +
                                "是 - 手动输入块名称\n" +
                                "否 - 重新选择文件\n" +
                                "取消 - 取消操作",
                                "读取失败", 
                                MessageBoxButtons.YesNoCancel, 
                                MessageBoxIcon.Question);
                            
                            if (result == DialogResult.Yes)
                            {
                                // 手动输入块名称
                                var manualBlockName = ShowInputDialog("手动输入块名称", "请输入块名称（建议先在AutoCAD中打开文件查看）:", "");
                                if (!string.IsNullOrEmpty(manualBlockName))
                                {
                                    txtSourceBlockName.Text = manualBlockName;
                                    
                                    // 默认目标块名与源块相同
                                    if (string.IsNullOrEmpty(txtTargetBlockName.Text))
                                    {
                                        txtTargetBlockName.Text = manualBlockName;
                                    }
                                    
                                    AppendLog($"手动设置源块: {manualBlockName}", Color.Blue);
                                    AppendLog($"目标块名称: {txtTargetBlockName.Text}", Color.Blue);
                                    AppendLog("注意：请确保输入的块名称在源文件中确实存在", Color.Orange);
                                    AppendLog("建议：可以先在AutoCAD中打开源文件，使用BEDIT命令查看块定义", Color.Gray);
                                    return;
                                }
                            }
                            else if (result == DialogResult.No)
                            {
                                // 重新选择文件的提示
                                AppendLog("请重新选择源块文件，建议选择较新版本的AutoCAD保存的文件", Color.Blue);
                            }
                            
                            // 用户选择重新选择文件或取消
                            _sourceBlockFile = string.Empty;
                            txtSourceBlockName.Text = string.Empty;
                            AppendLog("已取消选择源块", Color.Gray);
                            return;
                        }
                        
                        if (blockNames == null || blockNames.Count == 0)
                        {
                            AppendLog("文件中没有找到可用的块定义", Color.Orange);
                            MessageBox.Show("该DWG文件中没有可用的块定义！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            _sourceBlockFile = string.Empty;
                            txtSourceBlockName.Text = string.Empty;
                            return;
                        }
                        
                        AppendLog($"找到 {blockNames.Count} 个块定义", Color.Green);
                        
                        // 显示块选择对话框
                        using (var blockSelectionForm = new BlockSelectionForm(blockNames, _sourceBlockFile))
                        {
                            if (blockSelectionForm.ShowDialog(this) == DialogResult.OK)
                            {
                                var selectedBlockName = blockSelectionForm.SelectedBlockName;
                                txtSourceBlockName.Text = selectedBlockName;
                                
                                // 默认目标块名与源块相同
                                if (string.IsNullOrEmpty(txtTargetBlockName.Text))
                                {
                                    txtTargetBlockName.Text = selectedBlockName;
                                }
                                
                                AppendLog($"已选择源块: {selectedBlockName}", Color.Blue);
                                
                                if (txtTargetBlockName.Text == selectedBlockName)
                                {
                                    AppendLog($"目标块名称已设置为: {selectedBlockName}（与源块相同）", Color.Blue);
                                }
                                
                                // 显示块信息
                                var blockInfo = CADHelper.DwgBlockReader.GetBlockInfo(_sourceBlockFile, selectedBlockName);
                                if (blockInfo.Exists)
                                {
                                    AppendLog($"块信息: 实体数量={blockInfo.EntityCount}, 属性数量={blockInfo.AttributeDefinitions.Count}", Color.Gray);
                                }
                            }
                            else
                            {
                                // 用户取消选择
                                _sourceBlockFile = string.Empty;
                                txtSourceBlockName.Text = string.Empty;
                                AppendLog("已取消选择源块", Color.Gray);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError("选择源块文件时发生错误", ex);
                AppendLog($"选择源块文件时发生错误: {ex.Message}", Color.Red);
                _sourceBlockFile = string.Empty;
                txtSourceBlockName.Text = string.Empty;
            }
        }

        private void BtnBrowseTargetBlock_Click(object sender, EventArgs e)
        {
            // 目标块可以手动输入修改
            var blockName = ShowInputDialog("输入目标块名称", "目标块名称:", txtTargetBlockName.Text);
            if (!string.IsNullOrEmpty(blockName))
            {
                txtTargetBlockName.Text = blockName;
                AppendLog($"设置目标块名称: {blockName}", Color.Blue);
            }
        }

        private string ShowInputDialog(string title, string prompt, string defaultValue = "")
        {
            // 简单的输入对话框实现
            var result = Microsoft.VisualBasic.Interaction.InputBox(prompt, title, defaultValue);
            return result;
        }

        /// <summary>
        /// 根据读取失败的异常获取建议
        /// </summary>
        /// <param name="exception">异常对象</param>
        /// <returns>建议信息</returns>
        private string GetReadFailureSuggestion(Exception exception)
        {
            string message = exception.Message.ToLower();
            
            if (message.Contains("enoInputFiler".ToLower()) || message.Contains("无输入文件"))
            {
                return "建议：\n1. 确保在AutoCAD环境中运行此工具\n2. 尝试用较新版本的AutoCAD重新保存文件\n3. 检查文件是否损坏";
            }
            
            if (message.Contains("fileaccesserr") || message.Contains("访问") || message.Contains("占用"))
            {
                return "建议：\n1. 关闭AutoCAD中可能打开的此文件\n2. 检查文件权限\n3. 确保文件未被其他程序锁定";
            }
            
            if (message.Contains("version") || message.Contains("版本"))
            {
                return "建议：\n1. 使用较新版本的AutoCAD保存文件\n2. 检查DWG文件格式是否支持";
            }
            
            if (message.Contains("corrupt") || message.Contains("损坏"))
            {
                return "建议：\n1. 尝试在AutoCAD中打开并修复文件\n2. 使用AutoCAD的RECOVER命令\n3. 从备份文件恢复";
            }
            
            return "建议：\n1. 确保文件格式正确\n2. 尝试在AutoCAD中打开文件验证\n3. 检查文件完整性";
        }

        #endregion

        #region 执行控制

        private async void BtnPreview_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateInput())
                    return;

                if (_isProcessing)
                {
                    AppendLog("已有处理任务在进行中", Color.Orange);
                    return;
                }

                _isProcessing = true;
                SetControlsEnabled(false);

                AppendLog("开始预览块替换操作...", Color.Blue);

                // 创建配置
                var config = CreateReplaceConfig();
                
                AppendLog($"预览配置:");
                AppendLog($"  源块文件: {_sourceBlockFile}");
                AppendLog($"  源块名称: {config.SourceBlockName}");
                AppendLog($"  目标块名称: {config.TargetBlockName}");
                AppendLog($"  保留属性: {(config.PreserveAttributes ? "是" : "否")}");
                AppendLog($"  保留位置: {(config.PreservePosition ? "是" : "否")}");
                AppendLog($"  保留旋转: {(config.PreserveRotation ? "是" : "否")}");
                AppendLog($"  保留缩放: {(config.PreserveScale ? "是" : "否")}");
                AppendLog($"  创建备份: {(config.CreateBackup ? "是" : "否")}");
                
                // 执行预览
                AppendLog($"正在分析目录: {_selectedDirectory}", Color.Blue);
                
                var previewResult = BlockReplaceProcessor.PreviewDirectory(_selectedDirectory, config);
                
                // 显示预览结果
                AppendLog("预览结果:", Color.Green);
                AppendLog($"  总文件数: {previewResult.TotalFiles}");
                
                var processableFiles = previewResult.FileInfos.Count(f => f.CanProcess);
                var unprocessableFiles = previewResult.FileInfos.Count(f => !f.CanProcess);
                var totalEstimatedReplacements = previewResult.FileInfos.Sum(f => f.EstimatedReplacements);
                
                AppendLog($"  可处理文件: {processableFiles}", Color.Green);
                AppendLog($"  无法处理文件: {unprocessableFiles}", unprocessableFiles > 0 ? Color.Orange : Color.Black);
                AppendLog($"  预计替换块数: {totalEstimatedReplacements}", Color.Blue);
                
                // 显示可处理的文件列表（前10个）
                if (processableFiles > 0)
                {
                    AppendLog("可处理的文件:");
                    var processableFileList = previewResult.FileInfos.Where(f => f.CanProcess).Take(10);
                    foreach (var fileInfo in processableFileList)
                    {
                        var relativePath = GetRelativePath(_selectedDirectory, fileInfo.FilePath);
                        var sizeKB = fileInfo.FileSizeBytes / 1024.0;
                        AppendLog($"  ✓ {relativePath} ({fileInfo.EstimatedReplacements}个块, {sizeKB:F1}KB)");
                    }
                    if (processableFiles > 10)
                    {
                        AppendLog($"  ... 及其他 {processableFiles - 10} 个文件");
                    }
                }
                
                // 显示无法处理的文件列表（前5个）
                if (unprocessableFiles > 0)
                {
                    AppendLog("无法处理的文件:", Color.Orange);
                    var unprocessableFileList = previewResult.FileInfos.Where(f => !f.CanProcess).Take(5);
                    foreach (var fileInfo in unprocessableFileList)
                    {
                        var relativePath = GetRelativePath(_selectedDirectory, fileInfo.FilePath);
                        AppendLog($"  ✗ {relativePath} - {fileInfo.ErrorMessage}", Color.Red);
                    }
                    if (unprocessableFiles > 5)
                    {
                        AppendLog($"  ... 及其他 {unprocessableFiles - 5} 个文件", Color.Orange);
                    }
                }
                
                // 显示预览图像（如果有）
                var firstPreviewImage = previewResult.FileInfos.FirstOrDefault(f => f.PreviewImage != null)?.PreviewImage;
                if (firstPreviewImage != null)
                {
                    pbPreviewImage.Image = firstPreviewImage;
                    if (pbBlockThumbnail != null)
                    {
                        pbBlockThumbnail.Image = firstPreviewImage; // 同时显示在块缩略图区域
                    }
                    AppendLog("已显示第一个文件的预览图像", Color.Blue);
                }
                else
                {
                    pbPreviewImage.Image = null;
                    if (pbBlockThumbnail != null)
                    {
                        pbBlockThumbnail.Image = null;
                    }
                    AppendLog("未找到可预览的图像", Color.Gray);
                }
                
                if (totalEstimatedReplacements > 0)
                {
                    AppendLog($"预览完成！预计将替换 {totalEstimatedReplacements} 个块实例。如无问题请点击'开始替换'按钮", Color.Green);
                }
                else
                {
                    AppendLog("预览完成！未找到需要替换的块实例。请检查目标块名称是否正确。", Color.Orange);
                }
            }
            catch (Exception ex)
            {
                log.LogError("预览操作失败", ex);
                AppendLog($"预览操作失败: {ex.Message}", Color.Red);
            }
            finally
            {
                _isProcessing = false;
                SetControlsEnabled(true);
            }
        }

        private async void BtnStartReplace_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateInput())
                    return;

                if (_isProcessing)
                {
                    AppendLog("已有处理任务在进行中", Color.Orange);
                    return;
                }

                _isProcessing = true;
                SetControlsEnabled(false);

                AppendLog("开始批量块定义替换...", Color.Blue);

                // 创建配置
                var config = CreateReplaceConfig();
                
                // 在UI线程中处理（AutoCAD操作必须在UI线程中）
                AppendLog($"正在递归处理目录: {_selectedDirectory}", Color.Blue);
                
                var result = BlockReplaceProcessor.ProcessDirectory(_selectedDirectory, config);
                _lastResult = result;

                // 显示结果
                AppendLog("批量块替换完成！", Color.Green);
                AppendLog($"总文件: {result.TotalFiles}, 成功: {result.SuccessfulFiles}, 失败: {result.FailedFiles}", Color.Blue);
                AppendLog($"耗时: {result.Duration.TotalSeconds:F1} 秒", Color.Blue);

                // 显示详细结果
                var totalBlocksReplaced = result.FileResults.Sum(f => f.TotalBlocksReplaced);
                AppendLog($"总共替换了 {totalBlocksReplaced} 个块实例", Color.Green);

                // 生成报告
                var reportPath = Path.Combine(_selectedDirectory, $"块替换报告_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                BlockReplaceProcessor.SaveReport(result, reportPath);
                AppendLog($"详细报告已保存: {reportPath}", Color.Green);
                
                // 显示成功和失败的文件
                if (result.SuccessfulFiles > 0)
                {
                    AppendLog("成功处理的文件:");
                    foreach (var fileResult in result.FileResults.Where(f => f.Success).Take(5))
                    {
                        AppendLog($"  ✓ {fileResult.FileName} - {fileResult.TotalBlocksReplaced} 个块", Color.Green);
                    }
                    if (result.SuccessfulFiles > 5)
                    {
                        AppendLog($"  ... 及其他 {result.SuccessfulFiles - 5} 个文件", Color.Green);
                    }
                }
                
                if (result.FailedFiles > 0)
                {
                    AppendLog("失败的文件:", Color.Red);
                    foreach (var fileResult in result.FileResults.Where(f => !f.Success).Take(5))
                    {
                        AppendLog($"  ✗ {fileResult.FileName} - {fileResult.ErrorMessage}", Color.Red);
                    }
                    if (result.FailedFiles > 5)
                    {
                        AppendLog($"  ... 及其他 {result.FailedFiles - 5} 个文件", Color.Red);
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError("批量替换失败", ex);
                AppendLog($"批量替换失败: {ex.Message}", Color.Red);
            }
            finally
            {
                _isProcessing = false;
                SetControlsEnabled(true);
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            try
            {
                // 请求取消处理
                BlockReplaceProcessor.RequestCancellation();
                AppendLog("用户请求停止操作，正在终止当前处理...", Color.Orange);
                
                // 更新进度显示
                lblProgress.Text = "正在停止...";
                
                // 禁用停止按钮防止重复点击
                btnStop.Enabled = false;
            }
            catch (Exception ex)
            {
                log.LogError("停止操作失败", ex);
                AppendLog($"停止操作失败: {ex.Message}", Color.Red);
            }
        }

        #endregion

        #region 日志管理

        private void BtnClearLog_Click(object sender, EventArgs e)
        {
            rtbLog.Clear();
        }

        private void BtnSaveReport_Click(object sender, EventArgs e)
        {
            try
            {
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
                    saveDialog.DefaultExt = "txt";
                    saveDialog.FileName = $"块替换日志_{DateTime.Now:yyyyMMdd_HHmmss}.txt";

                    if (saveDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        File.WriteAllText(saveDialog.FileName, rtbLog.Text, System.Text.Encoding.UTF8);
                        AppendLog($"日志已保存: {saveDialog.FileName}", Color.Green);
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError("保存日志失败", ex);
                AppendLog($"保存日志失败: {ex.Message}", Color.Red);
            }
        }

        private void AppendLog(string message, Color? color = null)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, Color?>(AppendLog), message, color);
                return;
            }

            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var logMessage = $"[{timestamp}] {message}\n";

            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.SelectionLength = 0;
            rtbLog.SelectionColor = color ?? Color.Black;
            rtbLog.AppendText(logMessage);
            rtbLog.SelectionColor = rtbLog.ForeColor;
            rtbLog.ScrollToCaret();
        }

        #endregion

        #region 事件处理

        private void BlockReplaceForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_isProcessing)
            {
                var result = MessageBox.Show("正在处理中，确定要关闭吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }
        }

        #endregion
    
        #region 辅助方法

        /// <summary>
        /// 获取相对路径（替代.NET Core的Path.GetRelativePath）
        /// </summary>
        private static string GetRelativePath(string relativeTo, string path)
        {
            try
            {
                var relativeToUri = new Uri(relativeTo.EndsWith(Path.DirectorySeparatorChar.ToString()) ? relativeTo : relativeTo + Path.DirectorySeparatorChar);
                var pathUri = new Uri(path);
                
                if (relativeToUri.Scheme != pathUri.Scheme)
                {
                    return path; // 不同的方案，返回绝对路径
                }
                
                var relativeUri = relativeToUri.MakeRelativeUri(pathUri);
                var relativePath = Uri.UnescapeDataString(relativeUri.ToString());
                
                return relativePath.Replace('/', Path.DirectorySeparatorChar);
            }
            catch
            {
                return Path.GetFileName(path); // 出错时返回文件名
            }
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrEmpty(_selectedDirectory))
            {
                AppendLog("请先选择包含DWG文件的目录", Color.Red);
                return false;
            }

            if (!Directory.Exists(_selectedDirectory))
            {
                AppendLog("选择的目录不存在", Color.Red);
                return false;
            }

            var dwgFiles = Directory.GetFiles(_selectedDirectory, "*.dwg", SearchOption.AllDirectories);
            if (dwgFiles.Length == 0)
            {
                AppendLog("选择的目录中没有DWG文件", Color.Red);
                return false;
            }

            if (string.IsNullOrWhiteSpace(_sourceBlockFile) || !File.Exists(_sourceBlockFile))
            {
                AppendLog("请选择源块文件", Color.Red);
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSourceBlockName.Text))
            {
                AppendLog("源块名称不能为空", Color.Red);
                txtSourceBlockName.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtTargetBlockName.Text))
            {
                AppendLog("请输入目标块名称", Color.Red);
                txtTargetBlockName.Focus();
                return false;
            }

            return true;
        }

        private BlockReplaceProcessor.BlockReplaceConfig CreateReplaceConfig()
        {
            return new BlockReplaceProcessor.BlockReplaceConfig
            {
                SourceBlockFile = _sourceBlockFile, // 源块文件路径
                SourceBlockName = txtSourceBlockName.Text.Trim(),
                TargetBlockName = txtTargetBlockName.Text.Trim(),
                PreserveAttributes = chkPreserveAttributes.Checked,
                PreservePosition = chkPreservePosition.Checked,
                PreserveRotation = chkPreserveRotation.Checked,
                PreserveScale = chkPreserveScale.Checked,
                CreateBackup = chkCreateBackup.Checked
            };
        }

        private void SetControlsEnabled(bool enabled)
        {
            btnSelectDirectory.Enabled = enabled;
            btnClearDirectory.Enabled = enabled;
            txtSourceBlockName.Enabled = enabled;
            txtTargetBlockName.Enabled = enabled;
            btnBrowseSourceBlock.Enabled = enabled;
            btnBrowseTargetBlock.Enabled = enabled;
            chkPreserveAttributes.Enabled = enabled;
            chkPreservePosition.Enabled = enabled;
            chkPreserveRotation.Enabled = enabled;
            chkPreserveScale.Enabled = enabled;
            chkCreateBackup.Enabled = enabled;
            btnPreview.Enabled = enabled;
            btnStartReplace.Enabled = enabled;
            btnStop.Enabled = !enabled;
        }

        #endregion

        #region 进度回调

        /// <summary>
        /// 处理进度更新事件
        /// </summary>
        private void OnProgressUpdated(object sender, BlockReplaceProcessor.ProgressEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<object, BlockReplaceProcessor.ProgressEventArgs>(OnProgressUpdated), sender, e);
                return;
            }

            try
            {
                // 更新进度条
                if (e.TotalFiles > 0)
                {
                    int progress = (int)((double)e.CurrentFile / e.TotalFiles * 100);
                    if (progress > 100) progress = 100;
                    progressBar.Value = progress;
                    
                    // 更新进度标签
                    if (e.CurrentFile > 0)
                    {
                        lblProgress.Text = $"处理中... {e.CurrentFile}/{e.TotalFiles} ({progress}%)"; 
                        if (!string.IsNullOrEmpty(e.CurrentFileName))
                        {
                            lblProgress.Text += $" - {e.CurrentFileName}";
                        }
                    }
                    else
                    {
                        lblProgress.Text = e.Message;
                    }
                }
                else
                {
                    lblProgress.Text = e.Message;
                }

                // 添加日志
                Color logColor = e.IsError ? Color.Red : Color.Black;
                AppendLog(e.Message, logColor);
                
                // 检查是否操作完成或取消
                if (e.Message.Contains("目录处理完成") || e.Message.Contains("操作已取消"))
                {
                    _isProcessing = false;
                    SetControlsEnabled(true);
                }
            }
            catch (Exception ex)
            {
                log.LogError("处理进度更新事件时发生错误", ex);
            }
        }

        #endregion
    }
}