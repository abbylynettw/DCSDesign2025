using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MyOffice
{
    public static class FileUtilHelper
    {
        /// <summary>
        /// 递归查找所有包含图纸和Excel文件的目录，返回目录路径到文件列表的映射
        /// </summary>
        /// <param name="rootDirectoryPath">要搜索的根目录</param>
        /// <returns>包含有效图集的字典，键为目录路径，值为该目录下的所有文件列表</returns>
        public static Dictionary<string, List<string>> GetDrawingSets(string rootDirectoryPath)
        {
            var drawingSets = new Dictionary<string, List<string>>();

            try
            {
                // 递归搜索所有子目录
                SearchDirectoryForDrawingSets(rootDirectoryPath, drawingSets);
            }
            catch (Exception)
            {
                // 静默处理异常，返回已找到的结果
            }

            return drawingSets;
        }

        /// <summary>
        /// 递归搜索文件夹，查找有效的图集
        /// </summary>
        /// <param name="directoryPath">当前搜索的目录路径</param>
        /// <param name="drawingSets">结果集合</param>
        private static void SearchDirectoryForDrawingSets(string directoryPath, Dictionary<string, List<string>> drawingSets)
        {
            try
            {
                // 检查当前目录是否包含有效的图集
                bool isValidSet = CheckIfDirectoryContainsValidDrawingSet(directoryPath, out List<string> filesList);

                if (isValidSet)
                {
                    drawingSets[directoryPath] = filesList;
                }

                // 递归处理所有子目录
                string[] subDirectories = Directory.GetDirectories(directoryPath);
                foreach (string subDir in subDirectories)
                {
                    SearchDirectoryForDrawingSets(subDir, drawingSets);
                }
            }
            catch
            {
                // 静默处理异常，继续处理其他目录
            }
        }

        /// <summary>
        /// 检查目录是否包含有效的图集（同时有DWG和Excel文件）
        /// </summary>
        /// <param name="directoryPath">要检查的目录路径</param>
        /// <param name="filesList">输出参数，包含目录中找到的所有文件</param>
        /// <returns>如果目录同时包含DWG和Excel文件，则返回true</returns>
        private static bool CheckIfDirectoryContainsValidDrawingSet(string directoryPath, out List<string> filesList)
        {
            filesList = new List<string>();

            try
            {
                // 获取当前目录中的所有DWG文件
                string[] dwgFiles = Directory.GetFiles(directoryPath, "*.dwg", SearchOption.TopDirectoryOnly);

                // 获取当前目录中的所有Excel文件
                string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .Concat(Directory.GetFiles(directoryPath, "*.xls", SearchOption.TopDirectoryOnly))
                    .ToArray();

                // 如果同时存在DWG和Excel文件，则添加到列表
                if (dwgFiles.Length > 0 && excelFiles.Length > 0)
                {
                    filesList.AddRange(dwgFiles);
                    filesList.AddRange(excelFiles);
                    return true;
                }

                return false;
            }
            catch
            {
                // 如果访问目录出错，返回false
                return false;
            }
        }



        /// <summary>
        /// 备份指定目录下的文件
        /// </summary>
        /// <param name="directoryPath">源目录路径</param>
        /// <param name="filePattern">文件匹配模式，如"*.dwg"</param>
        /// <param name="backupPrefix">备份文件夹前缀</param>
        /// <param name="searchOption">搜索选项，默认包含子目录</param>
        /// <param name="preserveStructure">是否保留目录结构</param>
        /// <param name="logAction">日志记录委托</param>
        /// <param name="errorLogAction">错误日志记录委托</param>
        /// <returns>备份目录的完整路径</returns>
        public static string BackupFiles(
            string directoryPath,
            string filePattern,
            string backupPrefix = "Backup",
            SearchOption searchOption = SearchOption.AllDirectories,
            bool preserveStructure = true,
            Action<string> logAction = null,
            Action<string> errorLogAction = null)
        {
            try
            {
                // 创建带时间戳的备份目录
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupDir = Path.Combine(directoryPath, $"{backupPrefix}_{timestamp}");
                Directory.CreateDirectory(backupDir);

                // 获取要备份的文件
                string[] files = Directory.GetFiles(directoryPath, filePattern, searchOption);

                foreach (string file in files)
                {
                    string destPath;

                    if (preserveStructure)
                    {
                        // 保持目录结构
                        string relativePath = file.Substring(directoryPath.Length).TrimStart('\\', '/');
                        destPath = Path.Combine(backupDir, relativePath);
                        Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                    }
                    else
                    {
                        // 仅备份文件，不保留目录结构
                        destPath = Path.Combine(backupDir, Path.GetFileName(file));
                    }

                    // 复制文件
                    File.Copy(file, destPath);

                    // 记录日志
                    logAction?.Invoke($"已备份: {(preserveStructure ? destPath.Substring(backupDir.Length).TrimStart('\\', '/') : Path.GetFileName(file))}");
                }

                return backupDir;
            }
            catch (Exception ex)
            {
                string errorMessage = $"创建{filePattern}备份时出错: {ex.Message}";
                errorLogAction?.Invoke(errorMessage);
                throw; // 重新抛出异常，由调用者处理
            }
        }

        /// <summary>
        /// 备份多种类型的文件
        /// </summary>
        /// <param name="directoryPath">源目录路径</param>
        /// <param name="filePatterns">多个文件匹配模式，如["*.xlsx", "*.xls"]</param>
        /// <param name="backupPrefix">备份文件夹前缀</param>
        /// <param name="searchOption">搜索选项</param>
        /// <param name="preserveStructure">是否保留目录结构</param>
        /// <param name="logAction">日志记录委托</param>
        /// <param name="errorLogAction">错误日志记录委托</param>
        /// <returns>备份目录的完整路径</returns>
        public static string BackupMultipleFileTypes(
            string directoryPath,
            string[] filePatterns,
            string backupPrefix = "Backup",
            SearchOption searchOption = SearchOption.TopDirectoryOnly,
            bool preserveStructure = true,
            Action<string> logAction = null,
            Action<string> errorLogAction = null)
        {
            try
            {
                // 创建带时间戳的备份目录
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupDir = Path.Combine(directoryPath, $"{backupPrefix}_{timestamp}");
                Directory.CreateDirectory(backupDir);

                // 获取所有匹配的文件
                List<string> allFiles = new List<string>();
                foreach (string pattern in filePatterns)
                {
                    allFiles.AddRange(Directory.GetFiles(directoryPath, pattern, searchOption));
                }

                foreach (string file in allFiles)
                {
                    string destPath;

                    if (preserveStructure)
                    {
                        // 保持目录结构
                        string relativePath = file.Substring(directoryPath.Length).TrimStart('\\', '/');
                        destPath = Path.Combine(backupDir, relativePath);
                        Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                    }
                    else
                    {
                        // 仅备份文件，不保留目录结构
                        destPath = Path.Combine(backupDir, Path.GetFileName(file));
                    }

                    // 复制文件
                    File.Copy(file, destPath);

                    // 记录日志
                    logAction?.Invoke($"已备份: {(preserveStructure ? destPath.Substring(backupDir.Length).TrimStart('\\', '/') : Path.GetFileName(file))}");
                }

                return backupDir;
            }
            catch (Exception ex)
            {
                string errorMessage = $"创建文件备份时出错: {ex.Message}";
                errorLogAction?.Invoke(errorMessage);
                throw; // 重新抛出异常，由调用者处理
            }
        }
    }


}
