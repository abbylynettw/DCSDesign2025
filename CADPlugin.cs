﻿﻿﻿using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Threading;
using System.Diagnostics;
using log4net;
using MyOffice.LogHelper;
using MyOffice;
using System.ComponentModel;
using WinformUI;
using System.Drawing;
using WinformUI.UpdateTable;
using Autodesk.AutoCAD.DataExtraction;
using Autodesk.AutoCAD.Runtime;
using WinformUI.UpdateFrameTag;
using WinformUI.BlockReplace;


namespace DCSDesign2025
{
    /// <summary>
    /// CAD插件主类
    /// </summary>
    public class CADPlugin : IExtensionApplication
    {
        // 获取日志记录器
        private static readonly ILog Log = MyOffice.LogHelper.LogManager.GetLogger<CADPlugin>();

        // 菜单管理器
        private MenuManager menuManager;

        // CAD应用程序启动时执行
        public void Initialize()
        {
            try
            {
                // 初始化日志系统
                MyOffice.LogHelper.LogManager.Initialize();

                Log.LogInfo("工程图纸批量更新工具插件初始化开始");

                // 设置全局异常处理器
                SetupGlobalExceptionHandlers();

                // 初始化菜单
                InitializeMenu();

                // 在CAD命令行输出初始化信息
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\n工程图纸批量更新工具已加载\n");
                ed.WriteMessage("使用命令 \"BATCHUPGRADE\" 启动工程图纸批量更新工具\n");
                ed.WriteMessage("使用命令 \"BATCHUPDATETABLE\" 启动批量更新表格工具\n");
                ed.WriteMessage("使用命令 \"BLOCKREPLACE\" 启动批量块替换工具\n");
                ed.WriteMessage("或者在菜单栏中的 \"DCS工程工具\" 菜单下选择相应的命令\n");

                // 记录插件初始化成功
                Log.LogInfo("工程图纸批量更新工具插件初始化完成");
            }
            catch (System.Exception ex)
            {
                // 记录初始化失败
                MyOffice.LogHelper.LogManager.GetLogger<CADPlugin>().LogError("插件初始化失败", ex);

                // 尝试在命令行显示错误
                try
                {
                    Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                    ed.WriteMessage("\n工程图纸批量更新工具初始化失败: " + ex.Message + "\n");
                }
                catch
                {
                    // 忽略此处的异常，因为日志已经记录了主要错误
                }
            }
        }

        // 初始化菜单
        private void InitializeMenu()
        {
            try
            {
                Log.LogInfo("开始初始化菜单");

                // 创建并初始化菜单管理器
                menuManager = new MenuManager();
                menuManager.Initialize();

                Log.LogInfo("菜单初始化完成");
            }
            catch (System.Exception ex)
            {
                Log.LogError("初始化菜单时出错", ex);
                throw; // 重新抛出异常，让上层处理
            }
        }

        // 设置全局异常处理器
        private void SetupGlobalExceptionHandlers()
        {
            // 处理UI线程未捕获的异常
            Application.Idle += (sender, e) =>
            {
                try
                {
                    // 为WinForms应用程序设置全局异常处理
                    if (System.Windows.Forms.Application.OpenForms.Count > 0)
                    {
                        System.Windows.Forms.Application.ThreadException += OnThreadException;
                    }
                }
                catch (System.Exception ex)
                {
                    Log.LogError("设置WinForms异常处理器时出错", ex);
                }
            };

            // 处理非UI线程未捕获的异常
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        }

        // 处理UI线程异常
        private void OnThreadException(object sender, ThreadExceptionEventArgs e)
        {
            HandleException(e.Exception, "UI线程未捕获的异常");
        }

        // 处理非UI线程异常
        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            System.Exception ex = e.ExceptionObject as System.Exception;
            HandleException(ex, "AppDomain未捕获的异常");
        }

        // 统一异常处理逻辑
        private void HandleException(System.Exception ex, string source)
        {
            try
            {
                // 记录异常
                Log.LogFatal($"{source}: {ex?.Message}", ex);

                // 尝试写入Windows事件日志
                try
                {
                    if (!EventLog.SourceExists("DCSDesign2025"))
                    {
                        EventLog.CreateEventSource("DCSDesign2025", "Application");
                    }
                    EventLog.WriteEntry("DCSDesign2025",
                        $"严重错误: {ex?.Message}\n{ex?.StackTrace}",
                        EventLogEntryType.Error);
                }
                catch
                {
                    // 忽略写入事件日志的错误
                }

                // 在UI中显示错误消息
                string errorMessage = $"程序发生严重错误:\n\n{ex?.Message}\n\n" +
                    "错误详情已记录到日志文件中。\n" +
                    "请联系技术支持并提供日志文件。";

                MessageBox.Show(errorMessage, "严重错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                // 这是最后的防线，如果连错误处理器都出错，只能静默失败
            }
        }

        // CAD应用程序关闭时执行
        public void Terminate()
        {
            try
            {
                // 检查是否需要清理
                if (Application.DocumentManager?.IsApplicationContext == true)
                {
                    menuManager.Remove();
                }

                // 移除全局异常处理器
                AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
                System.Windows.Forms.Application.ThreadException -= OnThreadException;

                // 记录插件终止
                Log.LogInfo("工程图纸批量更新工具插件终止");
            }
            catch (System.Exception ex)
            {
                // 记录错误但不抛出，避免影响AutoCAD关闭
                System.Diagnostics.Debug.WriteLine($"插件终止清理时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量更新命令 - 启动WinForm界面
        /// </summary>
        [CommandMethod("BATCHUPGRADE", CommandFlags.Session)]
        public void BatchUpgrade()
        {
            try
            {
                Log.LogInfo("用户执行BATCHUPGRADE命令");

                // 获取当前活动文档
                Document doc = Application.DocumentManager.MdiActiveDocument;
                // 获取编辑器引用以输出消息
                Editor ed = doc.Editor;
                // 通知用户工具启动中
                ed.WriteMessage("\n正在启动工程图纸批量更新工具...\n");

                Log.LogDebug("准备显示工程图纸批量更新工具窗体");

                // 显示表单
                Application.ShowModelessDialog(new BatchUpgradeProjectForm());

                Log.LogInfo("工程图纸批量更新工具窗体已显示");
            }
            catch (System.Exception ex)
            {
                // 记录异常
                Log.LogError("启动工程图纸批量更新工具失败", ex);

                // 出现异常时在命令行显示错误信息
                try
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n错误: " + ex.Message);
                }
                catch
                {
                    // 忽略此处可能的异常
                }

                // 在对话框中显示更详细的错误信息
                MessageBox.Show("启动工程图纸批量更新工具时出错:\n\n" + ex.Message,
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 批量更新命令 - 启动WinForm界面
        /// </summary>
        [CommandMethod("BATCHUPDATETABLE", CommandFlags.Session)]
        public void BATCHUPDATETABLE()
        {
            try
            {
                Log.LogInfo("用户执行BATCHUPDATETABLE命令");

                // 获取当前活动文档
                Document doc = Application.DocumentManager.MdiActiveDocument;
                // 获取编辑器引用以输出消息
                Editor ed = doc.Editor;
                // 通知用户工具启动中
                ed.WriteMessage("\n正在启动批量更新表格工具...\n");
                var form = new BatchUpdateTableForm();
                form.Size = new Size(1310, 872); // 设置宽度为800，高度为600
                // 显示表单
                Application.ShowModelessDialog(form);
            }
            catch (System.Exception ex)
            {
                // 记录异常
                Log.LogError("启动批量更新表格工具失败", ex);

                // 出现异常时在命令行显示错误信息
                try
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n错误: " + ex.Message);
                }
                catch
                {
                    // 忽略此处可能的异常
                }
                // 在对话框中显示更详细的错误信息
                MessageBox.Show("启动批量更新表格工具时出错:\n\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 批量更新命令 - 启动WinForm界面
        /// </summary>
        [CommandMethod("UPDATEFRAME", CommandFlags.Session)]
        public void UPDATEFRAME()
        {
            try
            {               
                // 获取当前活动文档
                Document doc = Application.DocumentManager.MdiActiveDocument;
                // 获取编辑器引用以输出消息
                Editor ed = doc.Editor;             
                var form = new ReadExcelModifyCAD();
                form.Size = new Size(1383, 932);
                // 显示表单
                Application.ShowModelessDialog(form);
            }
            catch (System.Exception ex)
            {
                // 记录异常
                Log.LogError("启动更新图框工具失败", ex);

                // 出现异常时在命令行显示错误信息
                try
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n错误: " + ex.Message);
                }
                catch
                {
                    // 忽略此处可能的异常
                }
                // 在对话框中显示更详细的错误信息
                MessageBox.Show("启动更新图框工具时出错:\n\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 批量块替换命令 - 启动WinForm界面
        /// </summary>
        [CommandMethod("BLOCKREPLACE", CommandFlags.Session)]
        public void BLOCKREPLACE()
        {
            try
            {
                Log.LogInfo("用户执行BLOCKREPLACE命令");

                // 获取当前活动文档
                Document doc = Application.DocumentManager.MdiActiveDocument;
                // 获取编辑器引用以输出消息
                Editor ed = doc.Editor;
                // 通知用户工具启动中
                ed.WriteMessage("\n正在启动批量块替换工具...\n");

                Log.LogDebug("准备显示批量块替换工具窗体");

                // 显示表单
                var form = new WinformUI.BlockReplace.BlockReplaceForm();
                Application.ShowModelessDialog(form);

                Log.LogInfo("批量块替换工具窗体已显示");
            }
            catch (System.Exception ex)
            {
                // 记录异常
                Log.LogError("启动批量块替换工具失败", ex);

                // 出现异常时在命令行显示错误信息
                try
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n错误: " + ex.Message);
                }
                catch
                {
                    // 忽略此处可能的异常
                }

                // 在对话框中显示更详细的错误信息
                MessageBox.Show("启动批量块替换工具时出错:\n\n" + ex.Message,
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}