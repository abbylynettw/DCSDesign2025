using Microsoft.Win32;
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

namespace InstallerPlugin
{
    public partial class MainForm : Form
    {
        private string dllName = "DCSDesign2025.dll";

        // 定义支持的AutoCAD版本、产品代码和产品名称
        private Dictionary<string, (string Version, string Product)> acadInfo = new Dictionary<string, (string, string)>()
        {
            { "AutoCAD 2022", ("R24.1", "ACAD-5101:804") },
            { "AutoCAD 2023", ("R24.2", "ACAD-6101:804") }
        };

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 检测已安装的AutoCAD版本
            var installedVersions = DetectInstalledAcadVersions();

            if (installedVersions.Count == 0)
            {
                MessageBox.Show("未检测到支持的AutoCAD版本，请确保已安装AutoCAD 2022或2023。",
                                "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private List<string> DetectInstalledAcadVersions()
        {
            List<string> installedVersions = new List<string>();

            foreach (var item in acadInfo)
            {
                string acadName = item.Key;
                string version = item.Value.Version;
                string product = item.Value.Product;

                string regPath = $"SOFTWARE\\Autodesk\\AutoCAD\\{version}\\{product}";

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(regPath) ??
                                        Registry.LocalMachine.OpenSubKey(regPath))
                {
                    if (key != null)
                    {
                        installedVersions.Add(acadName);
                    }
                }
            }

            return installedVersions;
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void InstallButton_Click(object sender, EventArgs e)
        {
            string selectedVersion = GetSelectedVersion();
            if (string.IsNullOrEmpty(selectedVersion))
            {
                MessageBox.Show("请选择AutoCAD版本", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 获取当前应用程序目录
                string appPath = Path.GetDirectoryName(Application.ExecutablePath);
                string dllPath = Path.Combine(appPath, dllName);

                // 检查文件是否存在
                if (!File.Exists(dllPath))
                {
                    MessageBox.Show($"找不到插件文件: {dllPath}", "安装错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 检查选择的AutoCAD版本是否已安装
                if (!IsAcadVersionInstalled(selectedVersion))
                {
                    MessageBox.Show($"未检测到已安装的 {selectedVersion}，无法安装插件。请确保已安装此版本的AutoCAD。",
                                  "安装错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 获取版本信息
                string acadVersion = acadInfo[selectedVersion].Version;
                string acadProduct = acadInfo[selectedVersion].Product;

                // 安装注册表项
                bool registryUpdated = false;
                string regPath = $"SOFTWARE\\Autodesk\\AutoCAD\\{acadVersion}\\{acadProduct}\\Applications\\DCSDesign2025";

                try
                {
                    // 尝试写入HKCU (当前用户)
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(regPath))
                    {
                        if (key != null)
                        {
                            key.SetValue("LOADCTRLS", 2, RegistryValueKind.DWord);
                            key.SetValue("MANAGED", 1, RegistryValueKind.DWord);
                            key.SetValue("DESCRIPTION", "DCSDesign2025插件", RegistryValueKind.String);
                            key.SetValue("LOADER", dllPath, RegistryValueKind.String);
                            registryUpdated = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"注册表更新错误: {ex.Message}", "安装警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                if (registryUpdated)
                {
                    MessageBox.Show($"插件已成功安装到 {selectedVersion}。\n下次启动AutoCAD时，插件将自动加载。",
                                   "安装成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("安装过程中出现问题。可能无法自动加载插件。",
                                   "安装警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"安装过程中发生错误: {ex.Message}",
                               "安装错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UninstallButton_Click(object sender, EventArgs e)
        {
            string selectedVersion = GetSelectedVersion();
            if (string.IsNullOrEmpty(selectedVersion))
            {
                MessageBox.Show("请选择AutoCAD版本", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 获取版本信息
                string acadVersion = acadInfo[selectedVersion].Version;
                string acadProduct = acadInfo[selectedVersion].Product;

                // 卸载注册表项
                bool registryUpdated = false;
                string regPath = $"SOFTWARE\\Autodesk\\AutoCAD\\{acadVersion}\\{acadProduct}\\Applications\\DCSDesign2025";

                try
                {
                    // 尝试删除注册表项
                    Registry.CurrentUser.DeleteSubKey(regPath, false);
                    registryUpdated = true;
                }
                catch (Exception)
                {
                    // 注册表项可能不存在，忽略错误
                }

                if (registryUpdated)
                {
                    MessageBox.Show($"从 {selectedVersion} 成功卸载插件。",
                                   "卸载成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"在 {selectedVersion} 中未找到插件或卸载过程出现问题。",
                                   "卸载警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"卸载过程中发生错误: {ex.Message}",
                               "卸载错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsAcadVersionInstalled(string acadName)
        {
            if (!acadInfo.ContainsKey(acadName))
                return false;

            string version = acadInfo[acadName].Version;
            string product = acadInfo[acadName].Product;

            string regPath = $"SOFTWARE\\Autodesk\\AutoCAD\\{version}\\{product}";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(regPath) ??
                                    Registry.LocalMachine.OpenSubKey(regPath))
            {
                return key != null;
            }
        }

        private string GetSelectedVersion()
        {
            if (rb2022.Checked) return rb2022.Tag.ToString();
            if (rb2023.Checked) return rb2023.Tag.ToString();
            if (rb2024.Checked) return rb2024.Tag.ToString();
            return null;
        }
    }
}