﻿using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System;
using AutoCAD;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using MyOffice.LogHelper;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace DCSDesign2025
{
    /// <summary>
    /// 菜单管理类
    /// </summary>
    public class MenuManager
    {
        // 日志记录器
        private static readonly log4net.ILog log = MyOffice.LogHelper.LogManager.GetLogger<MenuManager>();

        // 菜单名称
        private const string MENU_NAME = "DCS设计2025";

        /// <summary>
        /// 初始化菜单
        /// </summary>
        public void Initialize()
        {
            try
            {               
                log.LogInfo("开始初始化菜单");

                // 添加菜单
                AddMenu();

                log.LogInfo("菜单初始化完成");
            }
            catch (Exception ex)
            {
                log.LogError("菜单初始化失败", ex);
                // 在命令行显示错误信息
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n菜单初始化失败: {ex.Message}\n");
            }
        }

        /// <summary>
        /// 添加菜单
        /// </summary>
        private void AddMenu()
        {
            try
            {
                // 获得CAD应用程序
                var app = (AcadApplication)Application.AcadApplication;

                // 检查菜单是否已经存在，如果存在则先移除
                RemoveExistingMenu(app);

                // 创建新菜单
                AcadPopupMenu pmParent = app.MenuGroups.Item(0).Menus.Add(MENU_NAME);

                // 添加菜单项
                pmParent.AddMenuItem(pmParent.Count + 1, "批量修改工程图纸(BATCHUPGRADE)", "BATCHUPGRADE\n");
                pmParent.AddMenuItem(pmParent.Count + 1, "批量更新表格(BATCHUPDATETABLE)", "BATCHUPDATETABLE\n");
                pmParent.AddMenuItem(pmParent.Count + 1, "更新图框标记值(UPDATEFRAME)", "UPDATEFRAME\n");
                pmParent.AddMenuItem(pmParent.Count + 1, "批量块替换(BLOCKREPLACE)", "BLOCKREPLACE\n");

                // 将创建的菜单加入到CAD的菜单栏中
                pmParent.InsertInMenuBar(app.MenuBar.Count + 1);

                log.LogInfo($"成功添加菜单: {MENU_NAME}");
            }
            catch (Exception ex)
            {
                log.LogError("添加菜单时出错", ex);
                throw; // 重新抛出异常，让上层处理
            }
        }

        /// <summary>
        /// 移除已存在的菜单
        /// </summary>
        private void RemoveExistingMenu(AcadApplication app)
        {
            try
            {
                // 遍历所有菜单项
                for (int i = 0; i < app.MenuBar.Count; i++)
                {
                    var menuItem = app.MenuBar.Item(i);
                    if (menuItem.Name == MENU_NAME)
                    {
                        // 移除找到的菜单
                        menuItem.RemoveFromMenuBar();
                        log.LogInfo($"已移除现有菜单: {MENU_NAME}");
                        break;
                    }
                }

                // 同时检查MenuGroups中的菜单
                for (int i = 0; i < app.MenuGroups.Item(0).Menus.Count; i++)
                {
                    if (app.MenuGroups.Item(0).Menus.Item(i).Name == MENU_NAME)
                    {
                        // 移除菜单组中的菜单
                        app.MenuGroups.Item(0).Menus.Item(i).RemoveFromMenuBar();
                        log.LogInfo($"已从菜单组中移除现有菜单: {MENU_NAME}");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError("移除现有菜单时出错", ex);
                // 这里吞掉异常，因为移除旧菜单失败不应阻止创建新菜单
            }
        }

        /// <summary>
        /// 移除菜单
        /// </summary>
        public void Remove()
        {
            try
            {
                // 获得CAD应用程序
                var app = (AcadApplication)Application.AcadApplication;

                // 移除菜单
                RemoveExistingMenu(app);

                log.LogInfo("菜单已移除");
            }
            catch (Exception ex)
            {
                log.LogError("移除菜单失败", ex);
            }
        }
    }
}