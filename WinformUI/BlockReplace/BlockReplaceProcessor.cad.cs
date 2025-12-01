using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using MyOffice.LogHelper;
using log4net;
using Autodesk.AutoCAD.DatabaseServices;
using AcadException = Autodesk.AutoCAD.Runtime.Exception;
using AcadErrorStatus = Autodesk.AutoCAD.Runtime.ErrorStatus;
using Autodesk.AutoCAD.ApplicationServices;

namespace WinformUI.BlockReplace
{
    /// <summary>
    /// 块替换处理器
    /// 用于处理块定义的替换和重新插入，递归处理目录中的所有DWG文件
    /// </summary>
    public class BlockReplaceProcessor
    {
        private static readonly ILog _logger = MyOffice.LogHelper.LogManager.GetLogger(typeof(BlockReplaceProcessor));

        // 定义进度事件和事件参数类
        public class ProgressEventArgs : EventArgs
        {
            public string Message { get; set; }
            public bool IsError { get; set; }
            public int CurrentFile { get; set; }
            public int TotalFiles { get; set; }
            public string CurrentFileName { get; set; }

            public ProgressEventArgs(string message, bool isError = false, int currentFile = 0, int totalFiles = 0, string currentFileName = "")
            {
                Message = message;
                IsError = isError;
                CurrentFile = currentFile;
                TotalFiles = totalFiles;
                CurrentFileName = currentFileName;
            }
        }

        // 声明进度事件
        public static event EventHandler<ProgressEventArgs> ProgressUpdated;

        // 取消令牌
        private static volatile bool _cancellationRequested = false;

        /// <summary>
        /// 请求取消处理
        /// </summary>
        public static void RequestCancellation()
        {
            _cancellationRequested = true;
        }

        /// <summary>
        /// 重置取消状态
        /// </summary>
        public static void ResetCancellation()
        {
            _cancellationRequested = false;
        }

        /// <summary>
        /// 检查是否已请求取消
        /// </summary>
        public static bool IsCancellationRequested => _cancellationRequested;

        // 触发进度事件的方法
        private static void OnProgressUpdated(string message, bool isError = false, int currentFile = 0, int totalFiles = 0, string currentFileName = "")
        {
            ProgressUpdated?.Invoke(null, new ProgressEventArgs(message, isError, currentFile, totalFiles, currentFileName));
        }

        /// <summary>
        /// 预览结果类
        /// </summary>
        public class PreviewResult
        {
            public string DirectoryPath { get; set; }
            public int TotalFiles { get; set; }
            public List<FilePreviewInfo> FileInfos { get; set; } = new List<FilePreviewInfo>();
            public DateTime ScanTime { get; set; }
        }

        /// <summary>
        /// 文件预览信息
        /// </summary>
        public class FilePreviewInfo
        {
            public string FilePath { get; set; }
            public string FileName => Path.GetFileName(FilePath);
            public bool CanProcess { get; set; }
            public int EstimatedReplacements { get; set; }
            public string ErrorMessage { get; set; }
            public DateTime FileLastModified { get; set; }
            public long FileSizeBytes { get; set; }
        }

        /// <summary>
        /// 块定义替换结果
        /// </summary>
        public class BlockDefinitionReplaceResult
        {
            public string OriginalBlockName { get; set; }
            public string NewBlockName { get; set; }
            public int ReplacedInstanceCount { get; set; }
            public List<string> PreservedAttributes { get; set; } = new List<string>();
            public List<string> LostAttributes { get; set; } = new List<string>();
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
        }

        /// <summary>
        /// 目录处理结果
        /// </summary>
        public class DirectoryProcessResult
        {
            public string DirectoryPath { get; set; }
            public int TotalFiles { get; set; }
            public int ProcessedFiles { get; set; }
            public int SuccessfulFiles { get; set; }
            public int FailedFiles { get; set; }
            public List<FileProcessResult> FileResults { get; set; } = new List<FileProcessResult>();
            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }
            public TimeSpan Duration => EndTime - StartTime;
        }

        /// <summary>
        /// 文件处理结果
        /// </summary>
        public class FileProcessResult
        {
            public string FilePath { get; set; }
            public string FileName => Path.GetFileName(FilePath);
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public int TotalBlocksReplaced { get; set; }
            public List<BlockDefinitionReplaceResult> BlockResults { get; set; } = new List<BlockDefinitionReplaceResult>();
            public DateTime ProcessTime { get; set; }
        }

        /// <summary>
        /// 块替换配置
        /// </summary>
        public class BlockReplaceConfig
        {
            public string SourceBlockFile { get; set; } // 源块文件路径
            public string SourceBlockName { get; set; }
            public string TargetBlockName { get; set; }
            public bool PreserveAttributes { get; set; } = true;
            public bool PreservePosition { get; set; } = true;
            public bool PreserveRotation { get; set; } = true;
            public bool PreserveScale { get; set; } = true;
            public bool CreateBackup { get; set; } = true;
        }

        /// <summary>
        /// 预览目录中的DWG文件，不实际修改文件
        /// </summary>
        /// <param name="directoryPath">目录路径</param>
        /// <param name="config">块替换配置</param>
        /// <returns>预览结果</returns>
        public static PreviewResult PreviewDirectory(string directoryPath, BlockReplaceConfig config)
        {
            var result = new PreviewResult
            {
                DirectoryPath = directoryPath,
                ScanTime = DateTime.Now
            };

            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    throw new DirectoryNotFoundException($"目录不存在: {directoryPath}");
                }

                OnProgressUpdated($"开始预览目录: {directoryPath}");
                _logger.Info($"开始预览目录: {directoryPath}");
                
                // 递归获取所有DWG文件
                var dwgFiles = Directory.GetFiles(directoryPath, "*.dwg", SearchOption.AllDirectories);
                result.TotalFiles = dwgFiles.Length;
                
                OnProgressUpdated($"找到 {dwgFiles.Length} 个DWG文件，开始分析...", false, 0, dwgFiles.Length);
                _logger.Info($"找到 {dwgFiles.Length} 个DWG文件");

                for (int i = 0; i < dwgFiles.Length; i++)
                {
                    // 检查取消请求
                    if (IsCancellationRequested)
                    {
                        OnProgressUpdated("用户取消了预览操作", false, i, dwgFiles.Length);
                        break;
                    }
                    
                    var filePath = dwgFiles[i];
                    var fileName = Path.GetFileName(filePath);
                    
                    OnProgressUpdated($"分析文件: {fileName}", false, i + 1, dwgFiles.Length, fileName);
                    
                    var fileInfo = PreviewSingleFile(filePath, config);
                    result.FileInfos.Add(fileInfo);
                }
                
                var processableFiles = result.FileInfos.Count(f => f.CanProcess);
                var totalEstimatedReplacements = result.FileInfos.Sum(f => f.EstimatedReplacements);
                
                string summaryMessage = IsCancellationRequested 
                    ? $"预览已取消: 已分析 {result.FileInfos.Count}/{result.TotalFiles} 个文件"
                    : $"预览完成: 可处理 {processableFiles}/{result.TotalFiles} 个文件，预计替换 {totalEstimatedReplacements} 个块实例";
                
                OnProgressUpdated(summaryMessage, false, result.FileInfos.Count, result.TotalFiles);
                _logger.Info(summaryMessage);
            }
            catch (Exception ex)
            {
                OnProgressUpdated($"预览失败: {ex.Message}", true);
                _logger.Error($"预览目录失败: {directoryPath}", ex);
                throw;
            }

            return result;
        }
        /// <summary>
        /// 预览单个DWG文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="config">配置</param>
        /// <returns>文件预览信息</returns>
        private static FilePreviewInfo PreviewSingleFile(string filePath, BlockReplaceConfig config)
        {
            var fileInfo = new FilePreviewInfo
            {
                FilePath = filePath,
                CanProcess = false,
                EstimatedReplacements = 0
            };

            try
            {
                if (!File.Exists(filePath))
                {
                    fileInfo.ErrorMessage = "文件不存在";
                    return fileInfo;
                }

                // 获取文件信息
                var fileInfoDetails = new FileInfo(filePath);
                fileInfo.FileLastModified = fileInfoDetails.LastWriteTime;
                fileInfo.FileSizeBytes = fileInfoDetails.Length;

                // 使用只读模式预览文件
                fileInfo.EstimatedReplacements = CountBlockReferencesInFile(filePath, config.TargetBlockName);
                fileInfo.CanProcess = true;
                
                _logger.Debug($"预览文件: {Path.GetFileName(filePath)} - 预计替换 {fileInfo.EstimatedReplacements} 个块");
            }
            catch (Exception ex)
            {
                fileInfo.CanProcess = false;
                fileInfo.ErrorMessage = GetFriendlyErrorMessage(ex);
                _logger.Warn($"预览文件失败: {Path.GetFileName(filePath)} - {ex.Message}");
            }

            return fileInfo;
        }

        /// <summary>
        /// 统计文件中指定块的引用数量（只读模式）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="blockName">块名称</param>
        /// <returns>引用数量</returns>
        private static int CountBlockReferencesInFile(string filePath, string blockName)
        {
            int count = 0;
            Database db = null;

            try
            {
                // 使用只读模式打开文件
                db = new Database(false, true);
                db.ReadDwgFile(filePath, FileShare.Read, true, "");

                using (var trans = db.TransactionManager.StartTransaction())
                {
                    count = CountBlockReferences(db, trans, blockName);
                    trans.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger.Debug($"统计文件块引用失败: {Path.GetFileName(filePath)} - {ex.Message}");
                // 预览模式下不抛出异常，返回0
                count = 0;
            }
            finally
            {
                db?.Dispose();
            }

            return count;
        }

        /// <summary>
        /// 递归处理指定目录中的所有DWG文件
        /// </summary>
        /// <param name="directoryPath">目录路径</param>
        /// <param name="config">块替换配置</param>
        /// <returns>处理结果</returns>
        public static DirectoryProcessResult ProcessDirectory(string directoryPath, BlockReplaceConfig config)
        {
            var result = new DirectoryProcessResult
            {
                DirectoryPath = directoryPath,
                StartTime = DateTime.Now
            };

            try
            {
                // 重置取消状态
                ResetCancellation();
                
                if (!Directory.Exists(directoryPath))
                {
                    throw new DirectoryNotFoundException($"目录不存在: {directoryPath}");
                }

                OnProgressUpdated($"开始递归处理目录: {directoryPath}");
                _logger.Info($"开始递归处理目录: {directoryPath}");
                
                // 递归获取所有DWG文件
                var dwgFiles = Directory.GetFiles(directoryPath, "*.dwg", SearchOption.AllDirectories);
                result.TotalFiles = dwgFiles.Length;
                
                OnProgressUpdated($"找到 {dwgFiles.Length} 个DWG文件", false, 0, dwgFiles.Length);
                _logger.Info($"找到 {dwgFiles.Length} 个DWG文件");

                for (int i = 0; i < dwgFiles.Length; i++)
                {
                    // 检查取消请求
                    if (IsCancellationRequested)
                    {
                        OnProgressUpdated("用户取消了操作", false, i, dwgFiles.Length);
                        _logger.Info("用户取消了操作");
                        break;
                    }
                    
                    var filePath = dwgFiles[i];
                    var fileName = Path.GetFileName(filePath);
                    
                    try
                    {
                        OnProgressUpdated($"正在处理文件: {fileName}", false, i + 1, dwgFiles.Length, fileName);
                        _logger.Info($"正在处理文件: {fileName}");
                        
                        var fileResult = ProcessSingleDwgFile(filePath, config);
                        result.FileResults.Add(fileResult);
                        result.ProcessedFiles++;
                        
                        if (fileResult.Success)
                        {
                            result.SuccessfulFiles++;
                            OnProgressUpdated($"✓ {fileName} - 成功替换 {fileResult.TotalBlocksReplaced} 个块实例", false, i + 1, dwgFiles.Length, fileName);
                            _logger.Info($"✓ {fileName} - 成功替换 {fileResult.TotalBlocksReplaced} 个块实例");
                        }
                        else
                        {
                            result.FailedFiles++;
                            OnProgressUpdated($"✗ {fileName} - {fileResult.ErrorMessage}", true, i + 1, dwgFiles.Length, fileName);
                            _logger.Error($"✗ {fileName} - {fileResult.ErrorMessage}");
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedFiles++;
                        var fileResult = new FileProcessResult
                        {
                            FilePath = filePath,
                            Success = false,
                            ErrorMessage = ex.Message,
                            ProcessTime = DateTime.Now
                        };
                        result.FileResults.Add(fileResult);
                        OnProgressUpdated($"✗ {fileName} - 处理异常: {ex.Message}", true, i + 1, dwgFiles.Length, fileName);
                        _logger.Error($"处理文件异常: {fileName}", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                OnProgressUpdated($"处理目录失败: {ex.Message}", true);
                _logger.Error($"处理目录失败: {directoryPath}", ex);
                throw;
            }
            finally
            {
                result.EndTime = DateTime.Now;
                string summaryMessage = IsCancellationRequested 
                    ? $"操作已取消: 已处理 {result.ProcessedFiles}/{result.TotalFiles} 个文件，成功 {result.SuccessfulFiles} 个，失败 {result.FailedFiles} 个"
                    : $"目录处理完成: 总计 {result.TotalFiles} 个文件，成功 {result.SuccessfulFiles} 个，失败 {result.FailedFiles} 个，耗时 {result.Duration.TotalSeconds:F1} 秒";
                OnProgressUpdated(summaryMessage, false, result.ProcessedFiles, result.TotalFiles);
                _logger.Info(summaryMessage);
            }

            return result;
        }

        /// <summary>
        /// 处理单个DWG文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="config">配置</param>
        /// <returns>处理结果</returns>
        private static FileProcessResult ProcessSingleDwgFile(string filePath, BlockReplaceConfig config)
        {
            var result = new FileProcessResult
            {
                FilePath = filePath,
                ProcessTime = DateTime.Now
            };

            try
            {
                if (!File.Exists(filePath))
                {
                    result.ErrorMessage = "文件不存在";
                    return result;
                }

                // 创建备份
                if (config.CreateBackup)
                {
                    CreateBackup(filePath);
                }

                // 执行块定义替换
                var blockResult = ReplaceBlockDefinition(filePath, config);
                result.BlockResults.Add(blockResult);
                
                if (blockResult.Success)
                {
                    result.TotalBlocksReplaced = blockResult.ReplacedInstanceCount;
                    result.Success = true;
                }
                else
                {
                    result.ErrorMessage = blockResult.ErrorMessage;
                }

                _logger.Info($"文件 {Path.GetFileName(filePath)} 处理完成");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.Error($"处理DWG文件失败: {filePath}", ex);
            }

            return result;
        }

        /// <summary>
        /// 替换块定义并重新插入
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="config">配置</param>
        /// <returns>替换结果</returns>
        private static BlockDefinitionReplaceResult ReplaceBlockDefinition(string filePath, BlockReplaceConfig config)
        {
            var result = new BlockDefinitionReplaceResult
            {
                OriginalBlockName = config.SourceBlockName,
                NewBlockName = config.TargetBlockName
            };

            try
            {
                _logger.Info($"开始替换块定义: {filePath} - {config.SourceBlockName} -> {config.TargetBlockName}");
                
                // 验证输入参数
                if (string.IsNullOrEmpty(config.SourceBlockFile) || !File.Exists(config.SourceBlockFile))
                {
                    throw new FileNotFoundException($"源块文件不存在: {config.SourceBlockFile}");
                }

                // 使用多重读取策略处理目标DWG文件
                var replaceResult = ReplaceBlockDefinitionInDwg(filePath, config);
                
                result.ReplacedInstanceCount = replaceResult.ReplacedCount;
                result.PreservedAttributes = replaceResult.PreservedAttributes;
                result.LostAttributes = replaceResult.LostAttributes;
                result.Success = true;
                
                _logger.Info($"块定义替换完成: {filePath} - 替换了 {result.ReplacedInstanceCount} 个块实例");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = GetFriendlyErrorMessage(ex);
                _logger.Error($"块定义替换失败: {filePath} - {config.SourceBlockName} -> {config.TargetBlockName}", ex);
            }

            return result;
        }

        /// <summary>
        /// 在DWG文件中执行块定义替换（使用Database.ReadDwgFile和WblockCloneObjects）
        /// </summary>
        /// <param name="targetFilePath">目标文件路径</param>
        /// <param name="config">配置</param>
        /// <returns>替换结果</returns>
        private static (int ReplacedCount, List<string> PreservedAttributes, List<string> LostAttributes)
            ReplaceBlockDefinitionInDwg(string targetFilePath, BlockReplaceConfig config)
        {
            var preservedAttributes = new List<string>();
            var lostAttributes = new List<string>();
            int replacedCount = 0;

            try
            {
                _logger.Info($"开始替换块定义: {targetFilePath}");
                
                // 使用高效的Database.ReadDwgFile方法
                replacedCount = PerformBlockReplacementWithDatabase(targetFilePath, config, 
                    out preservedAttributes, out lostAttributes);
                
                _logger.Info($"块定义替换完成: {targetFilePath} - 替换了 {replacedCount} 个块实例");
            }
            catch (System.Exception ex)
            {
                _logger.Error($"块替换操作失败: {targetFilePath}", ex);
                throw;
            }

            return (replacedCount, preservedAttributes, lostAttributes);
        }

        /// <summary>
        /// 使用Database.ReadDwgFile和WblockCloneObjects执行高效块替换（推荐方法）
        /// </summary>
        private static int PerformBlockReplacementWithDatabase(
            string targetFilePath,
            BlockReplaceConfig config,
            out List<string> preservedAttributes,
            out List<string> lostAttributes)
        {
            preservedAttributes = new List<string>();
            lostAttributes = new List<string>();
            int replacedCount = 0;

            Database sourceDb = null;
            Database targetDb = null;

            try
            {
                _logger.Info($"使用高效方法执行块替换: {config.SourceBlockName} -> {config.TargetBlockName}");

                // 1. 打开源文件数据库
                sourceDb = new Database(false, true);
                sourceDb.ReadDwgFile(config.SourceBlockFile, FileShare.Read, true, "");
                _logger.Info($"成功打开源文件: {config.SourceBlockFile}");

                // 2. 打开目标文件数据库
                targetDb = new Database(false, true);
                targetDb.ReadDwgFile(targetFilePath, FileShare.ReadWrite, true, "");
                _logger.Info($"成功打开目标文件: {targetFilePath}");

                // 3. 在源文件中定位源块
                ObjectId sourceBlockId = ObjectId.Null;
                using (var sourceTrans = sourceDb.TransactionManager.StartTransaction())
                {
                    var sourceBlockTable = (BlockTable)sourceTrans.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                    
                    if (!sourceBlockTable.Has(config.SourceBlockName))
                    {
                        throw new System.Exception($"源文件中未找到块定义: {config.SourceBlockName}");
                    }
                    
                    sourceBlockId = sourceBlockTable[config.SourceBlockName];
                    _logger.Info($"在源文件中找到块定义: {config.SourceBlockName}");
                    sourceTrans.Commit();
                }

                // 4. 使用WblockCloneObjects复制块到目标文件
                using (var targetTrans = targetDb.TransactionManager.StartTransaction())
                {
                    // 准备克隆对象集合
                    var sourceIds = new ObjectIdCollection { sourceBlockId };
                    var idMap = new IdMapping();

                    // 使用WblockCloneObjects克隆块定义
                    // DuplicateRecordCloning.Replace 参数会自动替换同名块
                    targetDb.WblockCloneObjects(
                        sourceIds,           // 要克隆的对象ID集合
                        targetDb.BlockTableId, // 目标容器（BlockTable）
                        idMap,               // ID映射表
                        DuplicateRecordCloning.Replace, // 替换同名块
                        false                // 不保存源数据库
                    );

                    // 5. 统计被替换的块引用数量
                    replacedCount = CountBlockReferences(targetDb, targetTrans, config.TargetBlockName);
                    
                    targetTrans.Commit();
                    _logger.Info($"成功克隆块定义，替换了 {replacedCount} 个块引用");
                }

                // 6. 保存目标文件
                targetDb.SaveAs(targetFilePath, DwgVersion.Current);
                _logger.Info($"成功保存目标文件: {targetFilePath}");
            }
            catch (System.Exception ex)
            {
                _logger.Error($"高效块替换方法失败", ex);
                throw;
            }
            finally
            {
                // 释放资源
                sourceDb?.Dispose();
                targetDb?.Dispose();
            }

            return replacedCount;
        }

        /// <summary>
        /// 统计指定块的引用数量
        /// </summary>
        /// <param name="db">数据库</param>
        /// <param name="trans">事务</param>
        /// <param name="blockName">块名称</param>
        /// <returns>引用数量</returns>
        private static int CountBlockReferences(Database db, Transaction trans, string blockName)
        {
            int count = 0;
            
            try
            {
                // 检查块是否存在
                var blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!blockTable.Has(blockName))
                {
                    _logger.Warn($"块定义不存在: {blockName}");
                    return 0;
                }
                
                ObjectId blockId = blockTable[blockName];
                
                // 获取模型空间ID
                ObjectId modelSpaceId = blockTable[BlockTableRecord.ModelSpace];
                var spaceIds = new List<ObjectId> { modelSpaceId };
                
                // 添加所有布局空间
                foreach (ObjectId layoutId in blockTable)
                {
                    var btr = (BlockTableRecord)trans.GetObject(layoutId, OpenMode.ForRead);
                    if (btr.IsLayout && btr.Name != BlockTableRecord.ModelSpace)
                    {
                        spaceIds.Add(layoutId);
                    }
                }
                
                // 在每个空间中查找块引用
                foreach (ObjectId spaceId in spaceIds)
                {
                    var space = (BlockTableRecord)trans.GetObject(spaceId, OpenMode.ForRead);
                    foreach (ObjectId entityId in space)
                    {
                        if (entityId.ObjectClass.DxfName == "INSERT")
                        {
                            var blockRef = (BlockReference)trans.GetObject(entityId, OpenMode.ForRead);
                            if (blockRef.BlockTableRecord == blockId)
                            {
                                count++;
                            }
                        }
                    }
                }
                
                _logger.Debug($"统计块引用数量: {blockName} = {count}");
            }
            catch (System.Exception ex)
            {
                _logger.Error($"统计块引用数量时出错: {blockName}", ex);
            }
            
            return count;
        }

        /// <summary>
        /// 获取友好的错误消息
        /// </summary>
        /// <param name="exception">异常对象</param>
        /// <returns>友好的错误消息</returns>
        private static string GetFriendlyErrorMessage(System.Exception exception)
        {
            if (exception is AcadException acadEx)
            {
                string errorMessage = "读取DWG文件失败。";
                switch (acadEx.ErrorStatus)
                {
                    case AcadErrorStatus.NoInputFiler:
                        errorMessage += "\n\n可能的解决方案：\n1. 确保在AutoCAD环境中运行此工具\n2. 文件可能版本过新，请用较新版本的AutoCAD保存\n3. 文件可能已损坏";
                        break;
                    case AcadErrorStatus.FileAccessErr:
                        errorMessage += "\n\n文件访问错误，请检查：\n1. 文件是否被其他程序占用\n2. 是否有足够的文件访问权限";
                        break;
                    case AcadErrorStatus.SelfReference:
                        errorMessage = "无法从源文件导入块定义: 块定义存在循环引用";
                        errorMessage += "\n\n立即解决方案：\n1. 在AutoCAD中打开源块文件\n2. 使用BEDIT命令编辑块定义\n3. 检查并删除循环引用的对象\n4. 或使用WBLOCK命令重新导出块";
                        errorMessage += "\n\n技术建议：\n• 块A不能包含块A自身\n• 检查嵌套块的引用关系\n• 重新创建有问题的块定义";
                        break;
                    case AcadErrorStatus.DuplicateBlockName:
                        errorMessage = "块名称冲突: 目标文件中已存在同名块定义";
                        break;
                    case AcadErrorStatus.InvalidDwgVersion:
                        errorMessage += "\nDWG文件版本不兼容";
                        break;
                    default:
                        errorMessage += $"\n\n错误代码: {acadEx.ErrorStatus}";
                        break;
                }
                return $"{errorMessage}\n\n错误详情: {acadEx.ErrorStatus}";
            }
            
            if (exception is FileNotFoundException)
            {
                return $"文件不存在: {exception.Message}";
            }
            
            if (exception is IOException)
            {
                return $"文件访问错误: {exception.Message}";
            }
            
            return $"处理失败: {exception.Message}";
        }

        /// <summary>
        /// 创建文件备份
        /// </summary>
        /// <param name="filePath">文件路径</param>
        private static void CreateBackup(string filePath)
        {
            try
            {
                var backupPath = filePath + ".bak_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                File.Copy(filePath, backupPath);
                _logger.Info($"已创建备份文件: {backupPath}");
            }
            catch (Exception ex)
            {
                _logger.Error($"创建备份失败: {filePath}", ex);
                throw;
            }
        }

        /// <summary>
        /// 验证块替换配置
        /// </summary>
        /// <param name="config">配置</param>
        /// <returns>验证结果</returns>
        public static List<string> ValidateConfig(BlockReplaceConfig config)
        {
            var errors = new List<string>();

            if (config == null)
            {
                errors.Add("配置不能为空");
                return errors;
            }

            if (string.IsNullOrEmpty(config.SourceBlockName))
            {
                errors.Add("源块名称不能为空");
            }

            if (string.IsNullOrEmpty(config.TargetBlockName))
            {
                errors.Add("目标块名称不能为空");
            }

            if (config.SourceBlockName == config.TargetBlockName)
            {
                errors.Add("源块和目标块不能相同");
            }

            return errors;
        }

        /// <summary>
        /// 生成处理报告
        /// </summary>
        /// <param name="result">处理结果</param>
        /// <returns>报告内容</returns>
        public static string GenerateReport(DirectoryProcessResult result)
        {
            var report = new StringBuilder();
            
            report.AppendLine("=== 块替换处理报告 ===");
            report.AppendLine($"处理目录: {result.DirectoryPath}");
            report.AppendLine($"开始时间: {result.StartTime:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"结束时间: {result.EndTime:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"耗时: {result.Duration.TotalSeconds:F1} 秒");
            report.AppendLine();
            
            report.AppendLine("——— 总体统计 ———");
            report.AppendLine($"总文件数: {result.TotalFiles}");
            report.AppendLine($"已处理: {result.ProcessedFiles}");
            report.AppendLine($"成功: {result.SuccessfulFiles}");
            report.AppendLine($"失败: {result.FailedFiles}");
            report.AppendLine($"成功率: {(result.TotalFiles > 0 ? (double)result.SuccessfulFiles / result.TotalFiles * 100 : 0):F1}%");
            report.AppendLine();
            
            if (result.SuccessfulFiles > 0)
            {
                report.AppendLine("——— 成功文件 ———");
                foreach (var fileResult in result.FileResults.Where(f => f.Success))
                {
                    report.AppendLine($"✓ {fileResult.FileName} - 替换了 {fileResult.TotalBlocksReplaced} 个块实例");
                }
                report.AppendLine();
            }
            
            if (result.FailedFiles > 0)
            {
                report.AppendLine("——— 失败文件 ———");
                foreach (var fileResult in result.FileResults.Where(f => !f.Success))
                {
                    report.AppendLine($"✗ {fileResult.FileName} - {fileResult.ErrorMessage}");
                }
            }
            
            return report.ToString();
        }

        /// <summary>
        /// 保存报告到文件
        /// </summary>
        /// <param name="result">处理结果</param>
        /// <param name="reportPath">报告文件路径</param>
        public static void SaveReport(DirectoryProcessResult result, string reportPath)
        {
            try
            {
                var reportContent = GenerateReport(result);
                File.WriteAllText(reportPath, reportContent, Encoding.UTF8);
                _logger.Info($"报告已保存到: {reportPath}");
            }
            catch (Exception ex)
            {
                _logger.Error($"保存报告失败: {reportPath}", ex);
                throw;
            }
        }
    }
}