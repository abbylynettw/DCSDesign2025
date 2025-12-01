using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using MyOffice.LogHelper;

namespace WinformUI.CADHelper
{
    /// <summary>
    /// DWG文件块定义读取工具
    /// </summary>
    public static class DwgBlockReader
    {
        private static readonly log4net.ILog _logger = LogManager.GetLogger(typeof(DwgBlockReader));

        /// <summary>
        /// 读取DWG文件中的所有块定义名称（使用文档管理器）
        /// </summary>
        /// <param name="dwgFilePath">DWG文件路径</param>
        /// <returns>块定义名称列表</returns>
        public static List<string> GetBlockDefinitions(string dwgFilePath)
        {
            var blockNames = new List<string>();

            try
            {
                if (!File.Exists(dwgFilePath))
                {
                    _logger.Error($"DWG文件不存在: {dwgFilePath}");
                    return blockNames;
                }

                _logger.Info($"开始读取DWG文件中的块定义: {dwgFilePath}");

                // 尝试使用文档管理器打开文件
                try
                {
                    var docs = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                    var currentDoc = docs.Open(dwgFilePath, false);
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = currentDoc;
                    
                    using (currentDoc.LockDocument())
                    {
                        using (var util = new DrawingUtility(currentDoc.Database, false))
                        {
                            using (var trans = currentDoc.Database.TransactionManager.StartTransaction())
                            {
                                var blockTable = (BlockTable)trans.GetObject(currentDoc.Database.BlockTableId, OpenMode.ForRead);

                                foreach (ObjectId blockId in blockTable)
                                {
                                    var blockRecord = (BlockTableRecord)trans.GetObject(blockId, OpenMode.ForRead);
                                    
                                    // 跳过系统块（以*开头的块和纸空间、模型空间）
                                    if (!blockRecord.Name.StartsWith("*") && 
                                        blockRecord.Name != BlockTableRecord.ModelSpace &&
                                        blockRecord.Name != BlockTableRecord.PaperSpace)
                                    {
                                        // 检查块是否包含实体（不是空块）
                                        if (blockRecord.HasAttributeDefinitions || blockRecord.Cast<ObjectId>().Any())
                                        {
                                            blockNames.Add(blockRecord.Name);
                                            _logger.Debug($"找到块定义: {blockRecord.Name}");
                                        }
                                    }
                                }

                                trans.Commit();
                            }
                        }
                    }
                    
                    // 关闭文档
                    currentDoc.CloseAndDiscard();
                    
                    _logger.Info($"成功读取到 {blockNames.Count} 个块定义");
                }
                catch (System.Exception ex)
                {
                    _logger.Warn($"文档管理器方式失败，尝试直接读取方式: {ex.Message}");
                    
                    // 如果文档管理器方式失败，尝试直接读取方式
                    return GetBlockDefinitionsDirectRead(dwgFilePath);
                }
            }
            catch (System.Exception ex)
            {
                _logger.Error($"读取DWG文件时发生异常: {dwgFilePath}", ex);
                throw new System.Exception($"读取DWG文件失败: {ex.Message}", ex);
            }

            return blockNames.OrderBy(x => x).ToList();
        }
        
        /// <summary>
        /// 直接读取DWG文件的备用方法（当文档管理器方式失败时）
        /// </summary>
        /// <param name="dwgFilePath">DWG文件路径</param>
        /// <returns>块定义名称列表</returns>
        private static List<string> GetBlockDefinitionsDirectRead(string dwgFilePath)
        {
            var blockNames = new List<string>();
            
            // 检查文件是否被占用
            try
            {
                using (var fs = File.Open(dwgFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // 文件可以正常打开
                }
            }
            catch (IOException ex)
            {
                _logger.Error($"文件被占用或无法访问: {dwgFilePath}", ex);
                throw new System.Exception($"文件被占用或无法访问，请确保文件未在AutoCAD中打开: {ex.Message}");
            }

            // 尝试多种方式读取DWG文件
            Database db = null;
            try
            {
                // 方法1：使用基本参数读取
                db = new Database(false, false);
                db.ReadDwgFile(dwgFilePath, FileShare.Read, true, "");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex) when (ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NoInputFiler)
            {
                _logger.Warn($"方法1失败，尝试方法2: {ex.Message}");
                db?.Dispose();
                
                try
                {
                    // 方法2：不验证完整性
                    db = new Database(false, false);
                    db.ReadDwgFile(dwgFilePath, FileShare.Read, false, "");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex2)
                {
                    _logger.Warn($"方法2失败，尝试方法3: {ex2.Message}");
                    db?.Dispose();
                    
                    try
                    {
                        // 方法3：使用不同的共享模式
                        db = new Database(false, false);
                        db.ReadDwgFile(dwgFilePath, FileShare.ReadWrite, false, "");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex3)
                    {
                        _logger.Error($"所有方法都失败: {ex3.Message}");
                        db?.Dispose();
                        throw new System.Exception($"无法读取DWG文件。可能的原因：\n1. 文件版本不兼容\n2. 文件已损坏\n3. 需要在AutoCAD环境中运行\n\n错误详情: {ex3.Message}");
                    }
                }
            }

            using (db)
            {
                using (var trans = db.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                    foreach (ObjectId blockId in blockTable)
                    {
                        var blockRecord = (BlockTableRecord)trans.GetObject(blockId, OpenMode.ForRead);
                        
                        // 跳过系统块（以*开头的块和纸空间、模型空间）
                        if (!blockRecord.Name.StartsWith("*") && 
                            blockRecord.Name != BlockTableRecord.ModelSpace &&
                            blockRecord.Name != BlockTableRecord.PaperSpace)
                        {
                            // 检查块是否包含实体（不是空块）
                            if (blockRecord.HasAttributeDefinitions || blockRecord.Cast<ObjectId>().Any())
                            {
                                blockNames.Add(blockRecord.Name);
                                _logger.Debug($"找到块定义: {blockRecord.Name}");
                            }
                        }
                    }

                    trans.Commit();
                }
            }
            
            return blockNames;
        }

        /// <summary>
        /// 验证块定义是否存在于DWG文件中
        /// </summary>
        /// <param name="dwgFilePath">DWG文件路径</param>
        /// <param name="blockName">块名称</param>
        /// <returns>是否存在</returns>
        public static bool BlockExists(string dwgFilePath, string blockName)
        {
            try
            {
                var blockNames = GetBlockDefinitions(dwgFilePath);
                return blockNames.Contains(blockName, StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                _logger.Error($"验证块定义存在性时发生异常: {dwgFilePath}, 块名: {blockName}", ex);
                return false;
            }
        }

        /// <summary>
        /// 获取块定义的详细信息
        /// </summary>
        /// <param name="dwgFilePath">DWG文件路径</param>
        /// <param name="blockName">块名称</param>
        /// <returns>块信息</returns>
        public static BlockInfo GetBlockInfo(string dwgFilePath, string blockName)
        {
            var blockInfo = new BlockInfo { Name = blockName };

            try
            {
                Database db = null;
                try
                {
                    // 尝试多种方式读取DWG文件
                    db = new Database(false, false);
                    db.ReadDwgFile(dwgFilePath, FileShare.Read, true, "");
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex) when (ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NoInputFiler)
                {
                    db?.Dispose();
                    db = new Database(false, false);
                    db.ReadDwgFile(dwgFilePath, FileShare.Read, false, "");
                }

                using (db)
                {
                    using (var trans = db.TransactionManager.StartTransaction())
                    {
                        var blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                        if (blockTable.Has(blockName))
                        {
                            var blockId = blockTable[blockName];
                            var blockRecord = (BlockTableRecord)trans.GetObject(blockId, OpenMode.ForRead);

                            blockInfo.Exists = true;
                            blockInfo.HasAttributes = blockRecord.HasAttributeDefinitions;
                            blockInfo.EntityCount = blockRecord.Cast<ObjectId>().Count();

                            // 获取属性定义
                            if (blockRecord.HasAttributeDefinitions)
                            {
                                foreach (ObjectId objId in blockRecord)
                                {
                                    var entity = trans.GetObject(objId, OpenMode.ForRead);
                                    if (entity is AttributeDefinition attrDef)
                                    {
                                        blockInfo.AttributeDefinitions.Add(new AttributeInfo
                                        {
                                            Tag = attrDef.Tag,
                                            Prompt = attrDef.Prompt,
                                            DefaultValue = attrDef.TextString
                                        });
                                    }
                                }
                            }
                        }

                        trans.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"获取块信息时发生异常: {dwgFilePath}, 块名: {blockName}", ex);
                blockInfo.ErrorMessage = ex.Message;
            }

            return blockInfo;
        }
    }

    /// <summary>
    /// 块信息
    /// </summary>
    public class BlockInfo
    {
        public string Name { get; set; }
        public bool Exists { get; set; }
        public bool HasAttributes { get; set; }
        public int EntityCount { get; set; }
        public List<AttributeInfo> AttributeDefinitions { get; set; } = new List<AttributeInfo>();
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 属性信息
    /// </summary>
    public class AttributeInfo
    {
        public string Tag { get; set; }
        public string Prompt { get; set; }
        public string DefaultValue { get; set; }
    }
}