using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

using System.Data;
using DataTable = Autodesk.AutoCAD.DatabaseServices.DataTable;

using System.Runtime.InteropServices;

using Line = Autodesk.AutoCAD.DatabaseServices.Line;

using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

using Autodesk.AutoCAD.Runtime;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace WinformUI.CADHelper
{
    public class DrawingUtility : IDisposable
    {
        public Database db;
        Transaction transaction;
        OpenCloseTransaction OpenCloseTransaction;
        public Document doc;
        public Editor editor;
        DocumentLock docLock;
        public DrawingUtility()
        {
            doc = Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            transaction = db.TransactionManager.StartTransaction();
            editor = doc.Editor;
            //docLock = doc.LockDocument();
        }
        public DrawingUtility(Database db, bool hasDoc = true)
        {
            this.db = db;
            transaction = db.TransactionManager.StartTransaction();
            if (hasDoc)
            {
                doc = db.GetDocument();
                editor = doc.Editor;
                //docLock = doc.LockDocument();
            }
        }
        public DrawingUtility(bool openCloseTrans)
        {
            doc = Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            OpenCloseTransaction = db.TransactionManager.StartOpenCloseTransaction();
            editor = doc.Editor;
        }
        public void Dispose()
        {
            try
            {
                if (transaction != null && !transaction.IsDisposed)
                {
                    // 只有当事务还没被提交或中止时才提交
                    transaction.Dispose();
                }
                // 其他清理代码...
            }
            catch (Exception ex)
            {
                // 错误处理...
            }
        }
        /// <summary>
        /// 提交事物
        /// </summary>
        public void Flush()
        {
            transaction.Commit();
            transaction = db.TransactionManager.StartTransaction();
        }
        /// <summary>
        /// 提交事物
        /// </summary>
        public void Commit()
        {
            transaction.Commit();
        }
        public void Abort()
        {
            transaction.Abort();
        }
        /// <summary>
        /// 用户拾取实体
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public ObjectId SelectObjectId(string message)
        {
            PromptEntityResult per = this.editor.GetEntity(message);
            if (per.Status == PromptStatus.OK)
            {
                return per.ObjectId;
            }
            return ObjectId.Null;
        }
        /// <summary>
        /// 用户拾取实体
        /// </summary>
        /// <param name="message"></param>
        /// <param name="allowedType"></param>
        /// <returns></returns>
        public Entity SelectEntity(string message, Type allowedType, OpenMode open = OpenMode.ForRead)
        {
            PromptEntityOptions peo = new PromptEntityOptions(message);

            peo.SetRejectMessage("\n您选择的不对，请选择正确的实体！");
            peo.AddAllowedClass(allowedType, true);
            PromptEntityResult per = this.editor.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {
                var ent = GetEntityByObjectId(per.ObjectId, open);
                return ent;
            }
            return null;
        }
        /// <summary>
        /// 用户拾取实体
        /// </summary>
        /// <param name="message"></param>
        /// <param name="allowedType"></param>
        /// <returns></returns>
        public ObjectId SelectEntity(string message, Type allowedType)
        {
            PromptEntityOptions peo = new PromptEntityOptions(message);
            peo.SetRejectMessage("\n您选择的不对，请选择正确的实体！");
            peo.AddAllowedClass(allowedType, true);
            PromptEntityResult per = this.editor.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {
                return per.ObjectId;
            }
            return ObjectId.Null;
        }

        public Point3d? GetPoint(string message)
        {
            PromptPointOptions pso = new PromptPointOptions(message);
            PromptPointResult psr = this.editor.GetPoint(pso);
            if (psr.Status == PromptStatus.OK)
            {
                return psr.Value;
            }
            return null;
        }
        /// <summary> 通过ObjectID获取实体</summary> 
        /// <param name="oi"></param>
        /// <returns></returns>
        public Entity GetEntityByObjectId(ObjectId id, OpenMode mode = OpenMode.ForRead)
        {
            Entity entity = null;
            if (id != ObjectId.Null && !id.IsErased)
            {
                entity = transaction.GetObject(id, mode) as Entity;
            }
            return entity;
        }
        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId</returns>
        public ObjectId AddToModelSpace(Entity ent)
        {
            ObjectId entId;//用于返回添加到模型空间中的实体ObjectId                         
            //以读方式打开块表
            BlockTable bt = (BlockTable)this.transaction.GetObject(db.BlockTableId, OpenMode.ForRead);
            //以写方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            entId = btr.AppendEntity(ent);//将图形对象的信息添加到块表记录中
            this.transaction.AddNewlyCreatedDBObject(ent, true);//把对象添加到事务处理中                            
            return entId; //返回实体的ObjectId
        }
    
        /// <summary>
        /// 新增或更新进线拓扑
        /// </summary>
        /// <param name="ent"></param>
        /// <param name="topuIn"></param>
        public void CreateOrUpdateTopuIn(DBObject ent, string topuIn)
        {
            //var topuStr = GetExtensionAsEntityAttri(ent, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //topu.In.Add(topuIn);
            //string topuResult = JsonConvert.SerializeObject(topu);
            //UpDateExtensionAsAttriValueOne(ent, "Topu", topuResult);
        }
        /// <summary>
        /// 添加出线拓扑
        /// </summary>
        /// <param name="ent"></param>
        /// <param name="topuOut"></param>
        public void AddTopuOut(DBObject ent, string topuOut)
        {
            //var topuStr = GetExtensionAsEntityAttri(ent, "Topu");

            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //topu.Out.Add(topuOut);
            //string topuResult = JsonConvert.SerializeObject(topu);
            //UpDateExtensionAsAttriValueOne(ent, "Topu", topuResult);
        }

        /// <summary>
        /// 获得选中的实体
        /// </summary>
        /// <returns></returns>
        public List<Entity> GetSelected()
        {
            List<Entity> entities = new List<Entity>();
            PromptSelectionResult selects = this.editor.SelectImplied();
            if (selects.Status == PromptStatus.OK)
            {
                ObjectId[] objIds = selects.Value.GetObjectIds();
                foreach (var id in objIds)
                {
                    Entity entity = GetEntity(id) as Entity;
                    entities.Add(entity);
                }
            }
            return entities;
        }

        /// <summary>
        /// 获取选择集
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public List<Entity> GetSelection(string message = "")
        {
            List<Entity> entities = new List<Entity>();
            PromptSelectionOptions options = new PromptSelectionOptions();
            options.MessageForAdding = message;
            PromptSelectionResult result = doc.Editor.GetSelection(options);
            if (result.Status != PromptStatus.OK)
            {
                return entities;
            }
            else
            {
                foreach (var item in result.Value.GetObjectIds())
                {
                    Entity en = GetEntity(item);
                    entities.Add(en);
                }
            }
            return entities;

        }
        /// <summary>
        /// PickFirst
        /// </summary>
        /// <param name="msg"></param>
        /// <returns></returns>
        public List<Entity> SelectFirst(string msg)
        {
            List<Entity> entities = new List<Entity>();
            //获取当前已选择的实体
            PromptSelectionResult psr = editor.SelectImplied();
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet ssr = psr.Value;//获取选择集
                foreach (var item in ssr.GetObjectIds())
                {
                    Entity en = GetEntity(item);
                    entities.Add(en);
                }
            }
            else
            {
                entities = GetSelection(msg);
            }
            return entities;
        }
        public List<T> SelectFirst<T>(string msg) where T : Entity
        {
            List<T> entities = new List<T>();
            //获取当前已选择的实体
            PromptSelectionResult psr = editor.SelectImplied();
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet ssr = psr.Value;//获取选择集
                foreach (var item in ssr.GetObjectIds())
                {
                    var en = GetEntity(item);
                    if (en is T)
                    {
                        T t = en as T;
                        entities.Add(t);
                    }
                }
            }
            else
            {
                entities = GetSelectionT<T>(msg);
            }
            return entities;
        }
        public List<T> GetSelectionT<T>(string message) where T : Entity
        {
            List<T> list = new List<T>();
            TypedValueList values = new TypedValueList();
            values.Add(typeof(T));
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = message;
            PromptSelectionResult psr = this.editor.GetSelection(pso, filter);
            if (psr.Status == PromptStatus.OK)
            {
                foreach (var objectId in psr.Value.GetObjectIds())
                {
                    var ent = this.GetEntity(objectId);
                    list.Add((T)ent);
                }
            }
            return list;
        }

        public List<Entity> GetSelectionT<T, T1>(string message) where T : Entity
        {
            List<Entity> list = new List<Entity>();
            TypedValueList values = new TypedValueList();
            values.Add(typeof(T));
            values.Add(typeof(T1));
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionResult psr = this.editor.SelectAll(filter);
            if (psr.Status == PromptStatus.OK)
            {
                foreach (var objectId in psr.Value.GetObjectIds())
                {
                    var ent = this.GetEntity(objectId);
                    list.Add(ent);
                }
            }
            return list;
        }
        /// <summary> 通过ObjectID获取实体</summary> 
        /// <param name="oi"></param>
        /// <returns></returns>
        public Entity GetEntity(ObjectId oi)
        {
            return GetEntity(oi, OpenMode.ForRead);
        }

        public Entity GetEntity(ObjectId oi, OpenMode mode)
        {
            Entity entity = null;
            if (oi != ObjectId.Null && !oi.IsErased)
            {
                entity = transaction.GetObject(oi, mode) as Entity;
            }
            return entity;
        }


        public DBObject GetDBObject(ObjectId oi, OpenMode mode)
        {
            DBObject entity = null;
            if (oi != ObjectId.Null && !oi.IsErased)
            {
                entity = transaction.GetObject(oi, OpenMode.ForRead) as DBObject;
                try
                {
                    if (mode == OpenMode.ForWrite)
                    {
                        entity.UpgradeOpen();
                    }
                }
                catch (Exception ex)
                {
                    //图型可能不允许编辑，比如图型所在图层已锁定                   
                }
            }
            return entity;
        }
        public List<Polyline> GetChild(BlockReference ent)
        {
            //var topuStr = GetExtensionAsEntityAttri(ent, "Topu");           
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //var childs = topu.Out;//获取出线
            //List<Polyline> ents = new List<Polyline>();
            //foreach (var childId in childs)
            //{
            //    var childEnt = GetEntity(childId);
            //    var pl = childEnt as Polyline;
            //    if (pl != null)
            //    {
            //        ents.Add(pl);
            //    }
            //}
            //return ents;
            return null;
        }
        public BlockReference GetChild(Polyline pline)
        {
            //var topuStr = GetExtensionAsEntityAttri(pline, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //var childs = topu.Out;//获取出线
            //if (childs.Count==1)
            //{
            //    var block = (BlockReference)GetEntity(childs[0]);                
            //    return block;
            //}
            return null;
        }
        public MText GetLabel(Entity block)
        {
            //var topuStr = GetExtensionAsEntityAttri(block, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //var label = topu.Label;//获取出线
            //var dbText = (MText)GetEntity(label);
            //return dbText;
            return null;
        }
        public string GetLabelText(Entity block)
        {
            //var topuStr = GetExtensionAsEntityAttri(block, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //var label = topu.Label;//获取出线
            //var dbText = (MText)GetEntity(label);
            //if (dbText !=null)
            //{
            //    return dbText.Contents;
            //}
            return "";
        }

        public ObjectId GetMLeaderEnts(ObjectId objectId)
        {   
            var blockLableRec = (BlockTableRecord)transaction.GetObject(objectId, OpenMode.ForRead);
            foreach (ObjectId objId in blockLableRec)
            {
                Entity ent = objId.GetObject(OpenMode.ForRead) as Entity;
                if (ent is AttributeDefinition)
                {
                    return objId;
                }               
            }
            return ObjectId.Null; 
        }
        public  MLeader GetMLeaderByTagAndValue(List<MLeader> mLeaders,string tag,string value)
        {
            foreach (var mLeader in mLeaders)
            {              
                var bId = mLeader.BlockContentId;
                var attrId = GetMLeaderEnts(bId);
                var attr = mLeader.GetBlockAttribute(attrId);
                if (attr.Tag == tag && attr.TextString == value)
                {                    
                    return mLeader;
                }               
            }
            return null;
          
        }
        public bool SetMLeaderTagAndValue(MLeader mLeader,string value)
        {
            var bId = mLeader.BlockContentId;
            var attrId = GetMLeaderEnts(bId);
            var attr = mLeader.GetBlockAttribute(attrId);            
            attr.TextString = value;
            mLeader.SetBlockAttribute(attrId, attr);
            return true;
        }
        public List<string> GetChildId(DBObject ent)
        {
            //var topuStr = GetExtensionAsEntityAttri(ent, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //var childs = topu.Out;//获取进线           
            //return childs;
            return null;
        }

        public List<string> GetParentId(DBObject ent)
        {
            //var topuStr = GetExtensionAsEntityAttri(ent, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //var parent = topu.In; //获取进线           
            //return parent;
            return null;
        }

        /// <summary>
        /// 移除拓扑
        /// </summary>
        /// <param name="erasedEnt"></param>
        /// <param name="removedId"></param>
        public void removeTopu(DBObject erasedEnt)
        {
            //var eraseId = GetLegendId(erasedEnt);
            //var topuStr = GetExtensionAsEntityAttri(erasedEnt, "Topu");
            //Topu topu = JsonConvert.DeserializeObject<Topu>(topuStr);
            //foreach (var parentId in topu.In)
            //{
            //    var parent = GetEntityByLegendId(parentId);
            //    var parentTopuStr = GetExtensionAsEntityAttri(parent, "Topu");
            //    Topu parentTopu = JsonConvert.DeserializeObject<Topu>(parentTopuStr);
            //    if (parentTopu.Out.Contains(eraseId))
            //    {
            //        parentTopu.Out.Remove(eraseId);
            //    }
            //    var parentTopuResult = JsonConvert.SerializeObject(parentTopu);
            //    UpDateExtensionAsAttriValueOne(parent, "Topu", parentTopuResult);
            //}

            //foreach (var childId in topu.Out)
            //{
            //    var child = GetEntityByLegendId(childId);
            //    var childTopuStr = GetExtensionAsEntityAttri(child, "Topu");
            //    Topu childTopu = JsonConvert.DeserializeObject<Topu>(childTopuStr);
            //    if (childTopu.In.Contains(eraseId))
            //    {
            //        childTopu.In.Remove(eraseId);
            //    }
            //    var childTopuResult = JsonConvert.SerializeObject(childTopu);
            //    UpDateExtensionAsAttriValueOne(child, "Topu", childTopuResult);
            //}
        }

        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public List<Point3d> GetAttributePositionInBlockReference(string attributeValue)
        {
            List<Point3d> positions = new List<Point3d>();
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = GetEntity(id) as BlockReference;
                if (br == null) continue;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if (attRef.TextString.ToUpper() == attributeValue.ToUpper())
                    {
                        positions.Add(attRef.AlignmentPoint);
                    }
                }
            }
            return positions; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public Dictionary<Point3d,BlockReference> GetAttributePositionAndBr(string attributeValue,string tagKeyWord)
        {
            Dictionary<Point3d, BlockReference> dic = new Dictionary<Point3d, BlockReference>();
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = GetEntity(id) as BlockReference;
                if (br == null) continue;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if ((string.IsNullOrEmpty(tagKeyWord) || attRef.Tag.ToUpper().Contains(tagKeyWord.ToUpper())) && attRef.TextString.ToUpper() == attributeValue.ToUpper())
                    {
                        if (!dic.ContainsKey(attRef.AlignmentPoint))
                        {
                            dic.Add(attRef.AlignmentPoint, br);
                        }
                    }
                }
            }
            return dic; //返回块属性值
        }

        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public Dictionary<Point3d, BlockReference> GetAttributePositionAndBr(string attributeValue)
        {
            Dictionary<Point3d, BlockReference> dic = new Dictionary<Point3d, BlockReference>();
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = GetEntity(id) as BlockReference;
                if (br == null) continue;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if (attRef.Tag.ToUpper() == attributeValue.ToUpper())
                    {
                        if (!dic.ContainsKey(attRef.AlignmentPoint))
                        {
                            dic.Add(attRef.AlignmentPoint, br);
                        }
                    }
                }
            }
            return dic; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public BlockReference? GetBrPositionByAttributeValue(string attributeValue)
        {            
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = GetEntity(id) as BlockReference;
                if (br == null) continue;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if (attRef.TextString.ToUpper() == attributeValue.ToUpper())
                    {
                        return br;
                    }
                }
            }
            return null; //返回块属性值
        }

        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public BlockReference? GetBrPositionByAttributeValueContains(string attributeValue)
        {
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = GetEntity(id) as BlockReference;
                if (br == null) continue;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if (attRef.TextString.ToUpper().Contains(attributeValue.ToUpper()))
                    {
                        return br;
                    }
                }
            }
            return null; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public List<Point3d> GetAttributePositionInBlockReferenceByTag(ObjectId blockReferenceId, string attributeName)
        {
            List<Point3d> positions = new List<Point3d>();
            // 获取块参照
            BlockReference bref = (BlockReference)this.transaction.GetObject(blockReferenceId, OpenMode.ForRead);
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                {
                    positions.Add(attRef.Position);
                }
            }
            return positions; //返回块属性值
        }

        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public Point3d? GetOneAttributePositionInBlockReferenceByTag(ObjectId blockReferenceId, string attributeName)
        {          
            // 获取块参照
            BlockReference bref = (BlockReference)this.transaction.GetObject(blockReferenceId, OpenMode.ForRead);
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                {
                    return attRef.AlignmentPoint;
                }
            }
            return null; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public List<Point3d> GetAttributePositionInBlockReferenceByTag(string attributeName)
        {
            List<Point3d> positions = new List<Point3d>();
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = GetEntity(id) as BlockReference;
                if (br == null) continue;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                    {
                        positions.Add(attRef.Position);
                    }
                }
            }                       
            return positions; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="bref"></param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public Point3d GetAttributePositionInBlockReference(BlockReference bref, string attributeName)
        {
            Point3d position = new Point3d();
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                {
                    position = attRef.Position;
                    break;
                }
            }
            return position; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="bref"></param>
        /// <returns>返回指定名称的块属性值</returns>
        public List<Point3d> GetAttributePositionsInBlockReference(BlockReference bref, string attributeValue)
        {
            List<Point3d> positions = new List<Point3d>();
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.TextString.ToUpper() == attributeValue.ToUpper())
                {
                    positions.Add(attRef.Position);
                    
                }
            }
            return positions; //返回块属性值
        }
        /// <summary>
        /// 获取指定名称的块属性参照位置
        /// </summary>
        /// <param name="bref"></param>
        /// <returns>返回指定名称的块属性值</returns>
        public string GetAttributeValueInBlockReferenceByTag(BlockReference bref, string tag)
        {           
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)transaction.GetObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.Tag.ToUpper() == tag.ToUpper())
                {
                    return attRef.TextString;

                }
            }
            return ""; //返回块属性值
        }
      
        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="rotateAngle">旋转角度</param>
        /// <param name="attNameValues">属性的名称与取值</param>
        /// <returns>返回块参照的Id</returns>
        public ObjectId InsertBlockReference(string blockName, Point3d position, Dictionary<string, string> attNameValues, string layerName = "")
        {

            var clayer = this.db.Clayer;
            if (layerName != "")
            {
                db.SetCurrentLayer(layerName);
            }
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues != null && attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            db.Clayer = clayer;
            return br.ObjectId;//返回添加的块参照的Id
        }

        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="attNameValues">属性的名称与取值</param>
        /// <returns>返回块参照的Id</returns>
        public ObjectId InsertBlockReference(string blockName, Point3d position)
        {

            var clayer = this.db.Clayer;
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //设置属性值
                        attribute.TextString = "";
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            db.Clayer = clayer;
            return br.ObjectId;//返回添加的块参照的Id
        }

        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="rotateAngle">旋转角度</param>
        /// <param name="attNameValues">属性的名称与取值</param>
        /// <returns>返回块参照的Id</returns>
        public ObjectId InsertBlockReferenceInModelSpace(string blockName, Point3d position, Dictionary<string, string> attNameValues, string layerName, double scale = 1)
        {

            var clayer = this.db.Clayer;
            if (layerName != "")
            {
                db.SetCurrentLayer(layerName);
            }
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）

            BlockTableRecord space = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            br.TransformBy(Matrix3d.Scaling(scale, position));
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            db.Clayer = clayer;
            return br.ObjectId;//返回添加的块参照的Id
        }

        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="rotateAngle">旋转角度</param>
        /// <param name="attNameValues">属性的名称与取值</param>
        /// <returns>返回块参照的Id</returns>
        public ObjectId InsertBlockReference(string blockName, Point3d position, Dictionary<string, string> attNameValues, ObjectId layerId, int colorIndex)
        {

            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.db.CurrentSpaceId.GetObject(OpenMode.ForWrite);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            br.LayerId = layerId;
            br.ColorIndex = colorIndex;
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            return br.ObjectId;//返回添加的块参照的Id
        }
        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="rotateAngle">旋转角度</param>
        /// <param name="attNameValues">属性的名称与取值</param>
        /// <returns>返回块参照的Id</returns>
        public ObjectId InsertBlockReference(BlockReference br, Dictionary<string, string> attNameValues)
        {

            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.db.CurrentSpaceId.GetObject(OpenMode.ForWrite);
            ObjectId btrId = bt[br.Name];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (attNameValues!=null&& record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            return br.ObjectId;//返回添加的块参照的Id
        }

        public ObjectId InsertBlockReferenceAndSetTag(BlockReference br, Dictionary<string, string> attNameValues)
        {
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.db.CurrentSpaceId.GetObject(OpenMode.ForWrite);
            ObjectId btrId = bt[br.Name];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        else
                        {
                            attribute.TextString = attDef.Tag;
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            return br.ObjectId;//返回添加的块参照的Id
        }
  
        /// <summary>
        /// 绘制标注
        /// </summary>
        /// <param name="ent"></param>
        /// <param name="text"></param>
        public void DrawLabel(ObjectId objectId, string text)
        {
            var ent = this.GetEntity(objectId);
            //获取实体的包围框
            var extents = ent.GeometricExtents;
            var p = new Point3d((extents.MinPoint.X + extents.MaxPoint.X) / 2, (extents.MinPoint.Y + extents.MaxPoint.Y) / 2, 0);
            p += Vector3d.YAxis * ((extents.MaxPoint.Y - extents.MinPoint.Y) / 2 + 4);
            var dbText = new DBText();
            dbText.TextString = text;
            dbText.Height = 3.5;
            dbText.Position = p;
            dbText.HorizontalMode = TextHorizontalMode.TextCenter;
            dbText.VerticalMode = TextVerticalMode.TextVerticalMid;
            dbText.AlignmentPoint = dbText.Position;
            this.AddToModelSpace(dbText);
            //this.Flush();
            //var guid = Guid.NewGuid().ToString("N");
            //this.CreateExtensionDefault(dbText, null, guid);
            //return guid;            
        }
        /// <summary>
        /// 绘制标注
        /// </summary>
        /// <param name="ent"></param>
        /// <param name="text"></param>
        public void DrawPlineLabel(string text, Point3d p)
        {
            //获取实体的包围框         
            var dbText = new DBText();
            dbText.TextString = text;
            dbText.Height = 3;
            dbText.Position = p;
            this.AddToModelSpace(dbText);
        }
        /// 创建新图层
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <returns>返回新建图层的ObjectId</returns>
        public ObjectId AddLayer(string layerName, Color color)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            if (!lt.Has(layerName))//如果不存在名为layerName的图层，则新建一个图层
            {
                //定义一个新的层表记录
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = layerName;//设置图层名
                ltr.Color = color;
                lt.UpgradeOpen();//切换层表的状态为写以添加新的图层
                //将层表记录的信息添加到层表中
                lt.Add(ltr);
                //把层表记录添加到事务处理中
                db.TransactionManager.AddNewlyCreatedDBObject(ltr, true);
                lt.DowngradeOpen();//为了安全，将层表的状态切换为读
            }
            return lt[layerName];//返回新添加的层表记录的ObjectId
        }


        /// <summary>
        /// 获取块参照的属性名和属性值
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="bref"></param>
        /// <returns>返回块参照的属性名和属性值</returns>
        public void GetAttributesInBlockReference(BlockReference bref, Dictionary<string, string> dic)
        {
            // 遍历块参照的属性，并将其属性名和属性值添加到字典中
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                if (!dic.ContainsKey(attRef.Tag))
                {
                    dic.Add(attRef.Tag, attRef.TextString);
                }
            }
        }
        /// <summary>
        /// 获取块参照的属性名和属性值
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <returns>返回块参照的属性名和属性值</returns>
        public Dictionary<string, string> GetAttributesInBlockReference(BlockReference block)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            // 遍历块参照的属性，并将其属性名和属性值添加到字典中
            foreach (ObjectId attId in block.AttributeCollection)
            {
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                if (!dic.ContainsKey(attRef.Tag))
                {
                    dic.Add(attRef.Tag, attRef.TextString);
                }
            }
            return dic;
        }
        /// <summary>
        /// 获取块参照的属性名和
        /// </summary>
        /// <returns>返回块参照的属性名和属性值</returns>
        public List<string> GetAttributesTagInBlockReference(BlockReference bref)
        {
            List<string> tags = new List<string>();
            // 遍历块参照的属性，并将其属性名和属性值添加到字典中
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                if (attRef != null)
                {
                    tags.Add(attRef.Tag);
                }
            }
            return tags;
        }
        public List<AttributeReference> GetAttributesReferenceInBlockReference(BlockReference bref)
        {
            List<AttributeReference> tags = new List<AttributeReference>();
            // 遍历块参照的属性，并将其属性名和属性值添加到字典中
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                if (attRef != null)
                {
                    tags.Add(attRef);
                }
            }
            return tags;
        }
        /// <summary>
        /// 获取指定块名的块参照
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="blockName">块名</param>
        /// <returns>返回指定块名的块参照</returns>
        public List<BlockReference> GetAllBlockReferences(string blockName)
        {
            List<BlockReference> blocks = new List<BlockReference>();

            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //打开指定块名的块表记录
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[blockName], OpenMode.ForRead);
            //获取指定块名的块参照集合的Id
            ObjectIdCollection blockIds = btr.GetBlockReferenceIds(true, true);
            foreach (ObjectId id in blockIds) // 遍历块参照的Id
            {
                //获取块参照
                BlockReference block = (BlockReference)GetDBObject(id, OpenMode.ForRead);
                blocks.Add(block); // 将块参照添加到返回列表 
            }
            return blocks; //返回块参照列表
        }
        /// <summary>
        /// 获取指定块名的块参照
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="blockName">块名</param>
        /// <returns>返回指定块名的块参照</returns>
        public BlockReference GetOneBlockReference(string blockName)
        {

            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //打开指定块名的块表记录
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[blockName], OpenMode.ForRead);
            //获取指定块名的块参照集合的Id
            ObjectIdCollection blockIds = btr.GetBlockReferenceIds(true, true);
            foreach (ObjectId id in blockIds) // 遍历块参照的Id
            {
                //获取块参照
                BlockReference block = (BlockReference)GetDBObject(id, OpenMode.ForRead);
                return block;
            }
            return null; //返回块参照列表
        }
        /// <summary>
        /// 获取指定块名的块参照
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="blockName">块名</param>
        /// <returns>返回指定块名的块参照</returns>
        public BlockReference GetOneBlockReference(string blockName, string layerName)
        {

            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //打开指定块名的块表记录
            if (bt.Has(blockName)) return null;
            BlockTableRecord btr = (BlockTableRecord)GetDBObject(bt[blockName], OpenMode.ForRead);
            if (btr == null) return null;
            //获取指定块名的块参照集合的Id
            ObjectIdCollection blockIds = btr.GetBlockReferenceIds(true, true);
            foreach (ObjectId id in blockIds) // 遍历块参照的Id
            {
                //获取块参照
                BlockReference block = (BlockReference)GetDBObject(id, OpenMode.ForRead);
                if (block.Layer == layerName)
                {
                    return block;
                }
            }
            return null; //返回块参照列表
        }
        public List<Entity> GetAllEntity(Database db, bool isIgnoreREVLayer = false)
        {
            var entities = new List<Entity>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased)
                {
                    //获取实体
                    Entity ent = GetEntity(id, OpenMode.ForRead);
                    if (isIgnoreREVLayer)
                    {
                        if (ent.Layer.Contains("REV")) continue; ;
                    }
                    entities.Add(ent);

                }
            }
            return entities;
        }
        public List<BlockReference> GetAllBlockReference(Database db, string blockName)
        {
            var entities = new List<BlockReference>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased)
                {
                    //获取实体
                    BlockReference ent = GetEntity(id, OpenMode.ForRead) as BlockReference;
                    if (ent != null && ent.Name == blockName)
                    {
                        entities.Add(ent);
                    }
                }
            }
            return entities;
        }
        public List<T> GetAllEntity<T>() where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased && !btr.IsAnonymous)
                {
                    //获取实体
                    T ent = GetEntity(id, OpenMode.ForRead) as T;
                    if (ent != null)
                    {
                        entities.Add(ent);
                    }
                }
            }
            return entities;
        }

        public List<T> GetAllEntityInPaperSpace<T>() where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased && !btr.IsAnonymous)
                {
                    //获取实体
                    T ent = GetEntity(id, OpenMode.ForRead) as T;
                    if (ent != null)
                    {
                        entities.Add(ent);
                    }
                }
            }
            return entities;
        }
        public List<T> GetAllEntityInCurrentSpace<T>() where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased)
                {
                    //获取实体
                    T ent = GetEntity(id, OpenMode.ForRead) as T;
                    if (ent != null)
                    {
                        entities.Add(ent);
                    }
                }
            }
            return entities;
        }
        public List<T> GetAllEntityInModelSpace<T>() where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased)
                {
                    //获取实体
                    T ent = GetEntity(id, OpenMode.ForRead) as T;
                    if (ent != null)
                    {
                        entities.Add(ent);
                    }
                }
            }
            return entities;
        }
        public List<T> GetAllEntityIModelSpace<T>(bool isIgnoreREVLayer = true) where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                //if (!btr.IsAnonymous && !btr.IsLayout)
                //{
                if (!btr.IsErased)
                {
                    //获取实体
                    T ent = GetEntity(id, OpenMode.ForRead) as T;
                    if (isIgnoreREVLayer)
                    {
                        if (ent.Layer.Contains("REV")) continue; ;
                        var layer = GetDBObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        if (layer == null || !layer.IsPlottable)
                        {
                            continue;
                        }
                    }

                    if (ent != null)
                    {
                        entities.Add(ent);
                    }
                }
                //}

            }
            return entities;
        }
        public List<T> GetAllEntityInCurrentLayer<T>() where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            var layerId = db.Clayer;
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (btr.LayoutId != layerId) continue;
                if (!btr.IsAnonymous && !btr.IsLayout)
                {
                    if (!btr.IsErased)
                    {
                        //获取实体
                        T ent = GetEntity(id, OpenMode.ForRead) as T;
                        if (ent != null)
                        {
                            entities.Add(ent);
                        }
                    }
                }
               
            }
            return entities;
        }
        public List<T> GetAllEntityInLayer<T>(string layerName) where T : Entity
        {
            var entities = new List<T>();
            //打开块表
            BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
            //以读方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in btr) // 遍历块参照的Id
            {
                if (!btr.IsErased)
                {
                    //获取实体
                    T ent = GetEntity(id, OpenMode.ForRead) as T;
                    if (ent != null && ent.Layer == layerName)
                    {
                        entities.Add(ent);
                    }
                }
            }
            return entities;
        }
        /// <summary>
        /// 获取指定块名的块参照
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="blockName">块名</param>
        /// <returns>返回指定块名的块参照</returns>
        public void EraseAllBlockReferences(List<BlockReference> blocks)
        {
            foreach (var bref in blocks)
            {
                if (bref != null && !bref.IsErased)
                {
                    bref.UpgradeOpen();
                    bref.Erase();
                    bref.DowngradeOpen();
                }
            }
        }
        /// <summary>
        /// 为对象添加扩展数据
        /// </summary>
        /// <param name="id">对象的Id</param>
        /// <param name="regAppName">注册应用程序名</param>
        /// <param name="values">要添加的扩展数据项列表</param>
        public void AddXData(ObjectId id, string regAppName)
        {
            //获取数据库的注册应用程序表
            //RegAppTable regTable = (RegAppTable)db.RegAppTableId.GetObject(OpenMode.ForWrite);
            RegAppTable regTable = GetDBObject(db.RegAppTableId, OpenMode.ForWrite) as RegAppTable;
            //如里不存在名为regAppName的记录，则创建新的注册应用程序表记录
            if (!regTable.Has(regAppName))
            {
                //创建一个注册应用程序表记录用来表示扩展数据
                RegAppTableRecord regRecord = new RegAppTableRecord();
                regRecord.Name = regAppName;//设置扩展数据的名字
                //在注册应用程序表加入扩展数据，并通知事务处理
                regTable.Add(regRecord);
                db.TransactionManager.AddNewlyCreatedDBObject(regRecord, true);
            }
            //以写的方式打开要添加扩展数据的实体
            DBObject obj = GetDBObject(id, OpenMode.ForWrite);
            TypedValue value = new TypedValue((int)DxfCode.ExtendedDataRegAppName, regAppName);
            //将扩展数据的应用程序名添加到扩展数据项列表的第一项    
            obj.XData = new ResultBuffer();
            obj.XData.Add(value);//将新建的扩展数据附加到实体中             
            obj.DowngradeOpen();
            regTable.DowngradeOpen();//为了安全，将应用程序注册表切换为读的状态

        }
        /// <summary>
        /// 获取对象的扩展数据
        /// </summary>
        /// <param name="id">对象的Id</param>
        /// <param name="regAppName">注册应用程序名</param>
        /// <returns>返回获得的扩展数据</returns>
        public TypedValueList GetXData(ObjectId id, string regAppName)
        {
            TypedValueList values = new TypedValueList();
            //打开对象
            DBObject obj = id.GetObject(OpenMode.ForRead);
            //获取对象中名为regAppName的扩展数据
            values = obj.GetXDataForApplication(regAppName);
            return values;//返回获得的扩展数据
        }
        /// <summary>
        ///  获取许多块的块属性
        /// </summary>
        /// <param name="db"></param>
        /// <param name="util"></param>
        /// <param name="blockNames"></param>
        /// <returns></returns>
        public string GetBlockTagValue(DrawingUtility util, string tagName)
        {
            BlockTable bt = (BlockTable)util.GetDBObject(util.db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)util.GetDBObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = util.GetEntity(id) as BlockReference;
                if (br == null) continue;
                if (IsHasTag(br, tagName))
                {
                    var tagValue = GetAttributeInBlockReference(br, tagName);
                    return tagValue;
                }
            }
            return "";
        }
        /// <summary>
        ///  获取许多块的块属性
        /// </summary>
        /// <param name="db"></param>
        /// <param name="util"></param>
        /// <param name="blockNames"></param>
        /// <returns></returns>
        public List<string> GetBlockTagValues(DrawingUtility util, string tagName)
        {
            List<string> tagValues = new List<string>();
            BlockTable bt = (BlockTable)util.GetDBObject(util.db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)util.GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (var id in btr)
            {
                var br = util.GetEntity(id) as BlockReference;
                if (br == null) continue;
                if (IsHasTag(br, tagName))
                {
                    var tagValue = GetAttributeInBlockReference(br, tagName);
                    tagValues.Add(tagValue);
                }
            }
            return tagValues;
        }
        /// <summary>
        /// 获取指定名称的块属性值
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public bool IsHasTag(BlockReference bref, string attributeName)
        {
            string attributeValue = string.Empty; // 属性值
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                {
                    return true;
                }
            }
            return false; //返回块属性值
        }
        public bool IsHasTag(string blockName, string tagName)
        {
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.db.CurrentSpaceId.GetObject(OpenMode.ForRead);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null && attDef.Tag == tagName)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// 获取指定名称的块属性值
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public string GetAttributeInBlockReference(BlockReference bref, string attributeName)
        {
            string attributeValue = string.Empty; // 属性值
            // 遍历块参照的属性
            foreach (ObjectId attId in bref.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                //判断属性名是否为指定的属性名
                if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                {
                    attributeValue = attRef.TextString;//获取属性值
                    break;
                }
            }
            return attributeValue; //返回块属性值
        }
        

        public List<string> GetAttributeInBlockReference(string blockName)
        {
            List<string> tags = new List<string>();
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)this.db.CurrentSpaceId.GetObject(OpenMode.ForRead);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        tags.Add(attDef.Tag);
                    }
                }
            }
            return tags;
        }
        /// <summary>
        /// 获得块属性
        /// </summary>
        /// <param name="util"></param>
        /// <param name="tagName"></param>
        /// <returns></returns>
        public List<AttributeReference> GetAttributeReference(DrawingUtility util, List<BlockReference> brefs, List<MLeader> mLeaders, string tagName)
        {
            List<AttributeReference> attributes = new List<AttributeReference>();

            List<ObjectId> toErased = new List<ObjectId>();
            foreach (var bref in brefs)
            {
                try
                {
                    // 遍历块参照的属性
                    foreach (ObjectId attId in bref.AttributeCollection)
                    {
                        // 获取块参照属性对象
                        AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                        //判断属性名是否为指定的属性名
                        if (attRef.Tag.ToUpper() == tagName.ToUpper())
                        {
                            var value = attRef.TextString;
                            var p = attRef.Position;
                            attributes.Add(attRef);
                        }
                    }

                }
                catch (Exception)
                {


                }
            }
            foreach (var mleader in mLeaders)
            {
                if (mleader != null && mleader.ContentType == ContentType.BlockContent)
                {
                    DBObjectCollection dBObjectCollection = new DBObjectCollection();
                    mleader.Explode(dBObjectCollection);
                    foreach (var dbObject in dBObjectCollection)
                    {
                        if (dbObject is BlockReference)
                        {
                            var br = dbObject as Entity;
                            var objectId = util.AddToModelSpace(br);

                            var br1 = GetDBObject(objectId, OpenMode.ForRead) as BlockReference;
                            // 遍历块参照的属性
                            foreach (ObjectId attId in br1.AttributeCollection)
                            {
                                // 获取块参照属性对象
                                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                                //判断属性名是否为指定的属性名
                                if (attRef.Tag.ToUpper() == tagName.ToUpper())
                                {
                                    var value = attRef.TextString;
                                    var p = attRef.Position;
                                    attributes.Add(attRef);
                                }
                            }
                            toErased.Add(objectId);
                        }
                    }
                }
            }
            toErased.ForEach(e => e.Erase());
            return attributes;
        }
        /// <summary>
        /// 判断IN点 在上还是在左在右
        /// </summary>
        /// <param name="bref"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public ChildOnParent GetPDirection(BlockReference bref, Point3d p)
        {
            var extents = bref.GeometricExtents;
            var pLeft = new Point2d(extents.MinPoint.X, (extents.MaxPoint.Y + extents.MinPoint.Y) / 2);
            var pTop = new Point2d((extents.MinPoint.X + extents.MaxPoint.X) / 2, extents.MaxPoint.Y);
            var pRight = new Point2d(extents.MaxPoint.X, (extents.MaxPoint.Y + extents.MinPoint.Y) / 2);
            var pBottom = new Point2d((extents.MinPoint.X + extents.MaxPoint.X) / 2, extents.MinPoint.Y);
            var pts = new List<Point2d> { pLeft, pTop, pRight, pBottom };
            var pClosed = pts.MinBy(pt => pt.GetDistanceTo(p.toPoint2d()));
            if (pClosed == pLeft) return ChildOnParent.left;
            if (pClosed == pTop) return ChildOnParent.top;
            if (pClosed == pRight) return ChildOnParent.right;
            if (pClosed == pBottom) return ChildOnParent.bottom;
            return ChildOnParent.bottom;
        }
        /// <summary>
        /// 判断IN点 在上还是在左在右
        /// </summary>
        /// <param name="bref"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        public ChildOnParent GetPDirectionOnlyLR(BlockReference bref, Point3d p)
        {
            var extents = bref.GeometricExtents;
            var pLeft = new Point2d(extents.MinPoint.X, (extents.MaxPoint.Y + extents.MinPoint.Y) / 2);
            var pRight = new Point2d(extents.MaxPoint.X, (extents.MaxPoint.Y + extents.MinPoint.Y) / 2);
            var pts = new List<Point2d> { pLeft, pRight };
            var pClosed = pts.MinBy(pt => pt.GetDistanceTo(p.toPoint2d()));
            if (pClosed == pLeft) return ChildOnParent.left;
            if (pClosed == pRight) return ChildOnParent.right;
            return ChildOnParent.bottom;
        }
        public enum ChildOnParent
        {
            left,
            top,
            right,
            bottom,
            none,
        }

        public BlockReference GetBlockReference(string tagValue)
        {
            List<BlockReference> brs = db.GetEntsInDatabase<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                    if (attRef.TextString == tagValue)
                    {
                        return br;
                    }
                }
            }
            return null;
        }
        public List<BlockReference> GetBlockReferenceByTagValue(string tagValue)
        {
            List<BlockReference> blocks = new List<BlockReference>();
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                    if (attRef.TextString == tagValue)
                    {
                        blocks.Add(br);
                    }
                }
            }
            return blocks;
        }
        public List<BlockReference> GetBlockReferenceByTag(string tag)
        {
            List<BlockReference> blocks = new List<BlockReference>();
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                    if (attRef.Tag.ToUpper() == tag.ToUpper())
                    {
                        blocks.Add(br);
                    }
                }
            }
            return blocks;
        }
        public BlockReference GetBlockReferenceByTagAndValue(string tag, string value)
        {
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                    if (attRef.Tag.ToUpper() == tag.ToUpper() && attRef.TextString == value)
                    {
                        return br;
                    }
                }
            }
            return null;
        }
       

        public void SetBlockReferenceByTagAndValueWithConnectLine(string tag, string old, string newValue, Action<string> action)
        {
            bool isMod = false;
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if (attRef.Tag == tag && attRef.TextString == old)
                    {
                        attRef.TextString = newValue;
                        isMod = true;
                        action.Invoke(old);
                    }
                    attRef.DowngradeOpen();

                }
                if (isMod)
                {
                    BrefChanged(old, br);
                    isMod = false;
                }
            }
        }
        public void BrefChanged(string oldnameplateStr, BlockReference bref,List<ObjectId> exceptedIds=null)
        {
            var curNpName = GetAttributeInBlockReference(bref, "NAMEPLATE");
            //nameplate标记变化
            //找到和它相交的线
            var plines = GetAllEntity<Polyline>();
            if (exceptedIds!=null)
            {
                plines = plines.Where(p => !exceptedIds.Contains(p.ObjectId)).ToList();
            }
            var mtexts = GetAllEntity<MText>();
            if (exceptedIds != null)
            {
                mtexts = mtexts.Where(p => !exceptedIds.Contains(p.ObjectId)).ToList();
            }
            var blocks = GetAllEntity<BlockReference>();
            if (exceptedIds != null)
            {
                blocks = blocks.Where(p => !exceptedIds.Contains(p.ObjectId)).ToList();
            }
            List<BlockReference> duanziBrefs = new List<BlockReference>();
            if (bref.Name.Contains("终端固定件"))
            {
                var currentBr = bref;
                while (currentBr != null)
                {
                    currentBr = GetCrossBlockRef(blocks, currentBr);
                    if (currentBr != null)
                    {
                        duanziBrefs.Add(currentBr);
                    }
                }
                foreach (var br in duanziBrefs)
                {
                    foreach (var pline in plines)
                    {
                        Point3dCollection pts = new Point3dCollection();
                        br.IntersectWith(pline, Intersect.OnBothOperands, pts, 0, 0);
                        if (pts.Count > 0)
                        {
                            var pCross = pts[0];
                            Point3d pClosed = Math.Abs(pCross.DistanceTo(pline.StartPoint)) < Math.Abs(pCross.DistanceTo(pline.EndPoint)) ? pline.StartPoint : pline.EndPoint;
                            Point3d pAnother = Math.Abs(pCross.DistanceTo(pline.StartPoint)) < Math.Abs(pCross.DistanceTo(pline.EndPoint)) ? pline.EndPoint : pline.StartPoint;

                            //求线上的标注
                            var mtextOnLines = mtexts.Where(m => Math.Abs(m.Location.DistanceTo(pline.GetClosestPointTo(m.Location, false))) < 0.2);
                            if (mtextOnLines == null || mtextOnLines.Count() == 0) continue;
                            foreach (MText mtext in mtextOnLines)
                            {
                                if (!mtext.Contents.Contains(oldnameplateStr)) continue;
                                string[] arr = mtext.Text.Split('/');
                                if (arr.Length != 2) continue;
                                //求mtext的起点，mtext起点偏移1
                                Point3d p0 = mtext.Location;
                                Point3d p1 = mtext.Location + mtext.Direction;
                                //距离块附近的起点或者终点更近
                                if (Math.Abs(mtext.Location.DistanceTo(pClosed)) < Math.Abs(mtext.Location.DistanceTo(pAnother)))
                                {
                                    if (Math.Abs(p0.DistanceTo(pClosed)) < Math.Abs(p1.DistanceTo(pClosed)))
                                    {
                                        arr[0] = arr[0].Replace(oldnameplateStr, curNpName);
                                        var resultStr = $"{arr[0]}/{arr[1]}";
                                        mtext.UpgradeOpen();
                                        mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                        mtext.DowngradeOpen();
                                    }
                                    else
                                    {
                                        arr[1] = arr[1].Replace(oldnameplateStr, curNpName);
                                        var resultStr = $"{arr[0]}/{arr[1]}";
                                        mtext.UpgradeOpen();
                                        mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                        mtext.DowngradeOpen();
                                    }
                                }
                                else //距离另一个点更近
                                {
                                    if (Math.Abs(p0.DistanceTo(pAnother)) < Math.Abs(p1.DistanceTo(pAnother)))
                                    {
                                        arr[1] = arr[1].Replace(oldnameplateStr, curNpName);
                                        var resultStr = $"{arr[0]}/{arr[1]}";
                                        mtext.UpgradeOpen();
                                        mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                        mtext.DowngradeOpen();
                                    }
                                    else
                                    {
                                        arr[0] = arr[0].Replace(oldnameplateStr, curNpName);
                                        var resultStr = $"{arr[0]}/{arr[1]}";
                                        mtext.UpgradeOpen();
                                        mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                        mtext.DowngradeOpen();
                                    }
                                }


                            }
                        }
                    }
                }
            }
            else
            {

                foreach (var pline in plines)
                {
                    Point3dCollection pts = new Point3dCollection();
                    bref.IntersectWith(pline, Intersect.OnBothOperands, pts, 0, 0);
                    if (pts.Count > 0)
                    {
                        var pCross = pts[0];
                        Point3d pClosed = Math.Abs(pCross.DistanceTo(pline.StartPoint)) < Math.Abs(pCross.DistanceTo(pline.EndPoint)) ? pline.StartPoint : pline.EndPoint;
                        Point3d pAnother = Math.Abs(pCross.DistanceTo(pline.StartPoint)) < Math.Abs(pCross.DistanceTo(pline.EndPoint)) ? pline.EndPoint : pline.StartPoint;

                        //求线上的标注
                        var mtextOnLines = mtexts.Where(m => Math.Abs(m.Location.DistanceTo(pline.GetClosestPointTo(m.Location, false))) < 0.2);
                        if (mtextOnLines == null || mtextOnLines.Count() == 0) continue;
                        foreach (MText mtext in mtextOnLines)
                        {
                            if (!mtext.Contents.Contains(oldnameplateStr)) continue;
                            string[] arr = mtext.Text.Split('/');
                            if (arr.Length != 2) continue;
                            //求mtext的起点，mtext起点偏移1
                            Point3d p0 = mtext.Location;
                            Point3d p1 = mtext.Location + mtext.Direction;
                            //距离块附近的起点或者终点更近
                            if (Math.Abs(mtext.Location.DistanceTo(pClosed)) < Math.Abs(mtext.Location.DistanceTo(pAnother)))
                            {
                                if (Math.Abs(p0.DistanceTo(pClosed)) < Math.Abs(p1.DistanceTo(pClosed)))
                                {
                                    arr[0] = arr[0].Replace(oldnameplateStr, curNpName);
                                    var resultStr = $"{arr[0]}/{arr[1]}";
                                    mtext.UpgradeOpen();
                                    mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                    mtext.DowngradeOpen();
                                }
                                else
                                {
                                    arr[1] = arr[1].Replace(oldnameplateStr, curNpName);
                                    var resultStr = $"{arr[0]}/{arr[1]}";
                                    mtext.UpgradeOpen();
                                    mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                    mtext.DowngradeOpen();
                                }
                            }
                            else //距离另一个点更近
                            {
                                if (Math.Abs(p0.DistanceTo(pAnother)) < Math.Abs(p1.DistanceTo(pAnother)))
                                {
                                    arr[1] = arr[1].Replace(oldnameplateStr, curNpName);
                                    var resultStr = $"{arr[0]}/{arr[1]}";
                                    mtext.UpgradeOpen();
                                    mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                    mtext.DowngradeOpen();
                                }
                                else
                                {
                                    arr[0] = arr[0].Replace(oldnameplateStr, curNpName);
                                    var resultStr = $"{arr[0]}/{arr[1]}";
                                    mtext.UpgradeOpen();
                                    mtext.Contents = mtext.Contents.Replace(mtext.Text, resultStr);
                                    mtext.DowngradeOpen();
                                }
                            }
                        }
                    }
                }
            }
        }

        BlockReference GetCrossBlockRef(List<BlockReference> blocks, BlockReference br)
        {

            foreach (BlockReference brDown in blocks.Where(b => b != br && b.Position.Y < br.Position.Y))
            {
                Point3dCollection pts = new Point3dCollection();
                br.IntersectWith(brDown, Intersect.OnBothOperands, pts, 0, 0);
                var a = Math.Abs(br.GeometricExtents.MinPoint.Y - brDown.GeometricExtents.MaxPoint.Y);
                var b = Math.Abs((br.GeometricExtents.MaxPoint.X + br.GeometricExtents.MinPoint.X) / 2) - ((brDown.GeometricExtents.MaxPoint.X + brDown.GeometricExtents.MinPoint.X) / 2);
                if (pts.Count > 0 || (a < 0.001) && (b < 0.001))
                {
                    return brDown;
                }
            }
            return null;
        }
        public void SetBlockReferenceByTagAndValue(string tag, string old, string newValue)
        {
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs)
            {
                bool hasTagValue = false;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if (attRef.Tag == tag && attRef.TextString == old)
                    {
                        attRef.TextString = newValue == null ? "" : newValue;
                        hasTagValue = true;
                    }
                    attRef.DowngradeOpen();
                }
                if (hasTagValue) BrefChanged(old, br);
            }
        }        
        public void SetBlockReferenceByTagAndValueExceptGroup(string tag, string old, string newValue)
        {
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs)
            {
                if (IsDBObjectInGroup(br.ObjectId))
                {
                    continue;
                }
                bool hasTagValue = false;
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if (attRef.Tag == tag && attRef.TextString == old)
                    {
                        attRef.TextString = newValue;
                        hasTagValue = true;
                    }
                    attRef.DowngradeOpen();
                }
                if (hasTagValue) BrefChanged(old, br);
            }
        }
        public void SetBlockReferenceByTagAndValue(List<BlockReference> brs, string tag, string old, string newValue)
        {            
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if (attRef.Tag == tag && attRef.TextString == old)
                    {
                        attRef.TextString = newValue;
                    }
                    attRef.DowngradeOpen();
                }
            }
        }
        public void SetBlockReferenceByTagAndValue(string tag, string old, string newValue, List<ObjectId> exceptIds)
        {
            List<BlockReference> brs = GetAllEntity<BlockReference>();
            foreach (var br in brs.Where(b => !exceptIds.Contains(b.ObjectId)))
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if (attRef.Tag == tag && attRef.TextString == old)
                    {
                        if (newValue!=null)
                        {
                            attRef.TextString = newValue;
                        }
                       
                    }
                    attRef.DowngradeOpen();
                }
            }
        }
        public void SetBlockReferenceByTagAndValue(BlockReference br, string tag,  string newValue)
        {
            // 遍历块参照的属性
            foreach (ObjectId attId in br.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                if (attRef.Tag == tag)
                {
                    attRef.TextString = newValue;
                }
                attRef.DowngradeOpen();
            }
        }

        public void SetBlockReferenceByTagAndValue(BlockReference br, string tag, string oldValue, string newValue, Action<string> action = null, List<ObjectId> exceptIds=null)
        {
            // 遍历块参照的属性
            foreach (ObjectId attId in br.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                if (attRef.Tag.ToUpper() == tag.ToUpper() && attRef.TextString == oldValue)
                {
                    attRef.TextString = newValue;
                    action?.Invoke(oldValue);
                }
                attRef.DowngradeOpen();

            }
            BrefChanged(oldValue, br, exceptIds);
        }
        public void SetBlockReferenceByTagAndValueUnChange(BlockReference br, string tag, string oldValue, string newValue, Action<string> action = null, List<ObjectId> exceptIds = null)
        {
            // 遍历块参照的属性
            foreach (ObjectId attId in br.AttributeCollection)
            {
                // 获取块参照属性对象
                AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                if (attRef.Tag.ToUpper() == tag.ToUpper() && attRef.TextString == oldValue)
                {
                    attRef.TextString = newValue;
                    action?.Invoke(oldValue);
                }
                attRef.DowngradeOpen();

            }           
        }
        public void SetBlockReferenceByTagAndValue(string old, string newValue, ref int times, bool isContains = true)
        {
            List<BlockReference> brs = db.GetEntsInDatabase<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if ((isContains && attRef.TextString.Contains(old)) || attRef.TextString == old)
                    {
                        attRef.TextString = attRef.TextString.Replace(old, newValue);
                        times++;
                    }
                    attRef.DowngradeOpen();
                }
            }
        }
        public void SetBlockReferenceByTagAndValue(string old, string newValue, bool isContains = true)
        {
            List<BlockReference> brs = db.GetEntsInDatabase<BlockReference>();
            foreach (var br in brs)
            {
                // 遍历块参照的属性
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForWrite);
                    if ((isContains && attRef.TextString.Contains(old)) || attRef.TextString == old)
                    {
                        attRef.TextString = attRef.TextString.Replace(old, newValue);                       
                    }
                    attRef.DowngradeOpen();
                }
            }
        }

        public void DrawPline(Point3d pStart,double x,Vector3d vec)
        {

            var line = new Polyline();
            line.AddVertexAt(0, pStart.toPoint2d(), 0, 0, 0);
            line.AddVertexAt(1, (pStart + vec * x).toPoint2d(), 0, 0, 0);                 
            AddToModelSpace(line);
            MoveEntityToBottom(line.ObjectId);
        }
        public void DrawPline(Point3d pStart, Point3d pEnd, Vector3d vec)
        {

            var line = new Polyline();
            line.AddVertexAt(0, pStart.toPoint2d(), 0, 0, 0);
            line.AddVertexAt(1, pEnd.toPoint2d(), 0, 0, 0);
            AddToModelSpace(line);
            MoveEntityToBottom(line.ObjectId);
        }
        public ObjectId DrawPline(params Point3d[] pts)
        {
            var line = new Polyline();
            for (int i = 0; i < pts.Length; i++)
            {
                line.AddVertexAt(0, pts[i].toPoint2d(), 0, 0, 0);
            }
            var id = AddToModelSpace(line);
            MoveEntityToBottom(line.ObjectId);
            return line.ObjectId;
        }
        public void DrawCircle(Point3d pCenter, double r)
        {
            var cir = new Circle();
            cir.Center = pCenter;
            cir.Radius = r;           
            AddToModelSpace(cir);
        }
        public bool BatchUpdateText(string sourceText, string targetText)
        {
            if (sourceText == targetText) return false;
            if (string.IsNullOrEmpty(sourceText) || targetText == null) return false;
            try
            {
                List<BlockReference> brs = db.GetEntsInDatabase<BlockReference>();
                foreach (var br in brs)
                {
                    // 遍历块参照的属性
                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        // 获取块参照属性对象
                        AttributeReference attRef = (AttributeReference)GetDBObject(attId, OpenMode.ForRead);
                        if (attRef.TextString == sourceText)
                        {
                            attRef.UpgradeOpen();
                            attRef.TextString = targetText;
                            //attRef.ColorIndex = 1;
                            attRef.DowngradeOpen();
                        }
                    }
                }
                List<MText> mtexts = db.GetEntsInDatabase<MText>();
                foreach (var mtext in mtexts)
                {
                    if (mtext.Text.Contains(sourceText))
                    {
                        mtext.UpgradeOpen();
                        mtext.Contents = mtext.Contents.Replace(sourceText, targetText);
                        //mtext.ColorIndex = 1;
                        mtext.DowngradeOpen();
                    }
                }
                List<DBText> dBTexts = db.GetEntsInDatabase<DBText>();
                foreach (var dbText in dBTexts)
                {
                    if (dbText.TextString.Contains(sourceText))
                    {
                        dbText.UpgradeOpen();
                        dbText.TextString = dbText.TextString.Replace(sourceText, targetText);
                        //dbText.ColorIndex = 1;
                        dbText.DowngradeOpen();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }
        public bool CopyEntitiesToCurrent(string dwgPath, Point3d basePoint, Matrix3d? matrix3d = null)
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //读取指定目录的数据到新数据库
            using (Database newDB = new Database(false, true))
            {
                // 读取图纸数据库
                newDB.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
                newDB.CloseInput(true);
                var mtx = Matrix3d.Displacement(Point3d.Origin.GetVectorTo(basePoint));
                if (matrix3d != null)
                {
                    mtx *= matrix3d.Value;
                }
                this.db.Insert(mtx, newDB, false);
                return true;
            }
        }       
        public void CopyEntitiesToCurrent(string dwgPath, string blockName)
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //读取指定目录的数据到新数据库
            using (Database newDB = new Database(false, true))
            {
                // 读取图纸数据库
                newDB.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
                newDB.CloseInput(true);
                var objectid = this.db.Insert(blockName, newDB, false);
                //objectid.Erase();
            }
        }
        public Matrix3d GetMatrixInDatabase(string dwgPath, Point3d pBase, double Y)
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //读取指定目录的数据到新数据库
            using (Database newDB = new Database(false, true))
            {
                // 读取图纸数据库
                newDB.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
                Plane plane = new Plane();
                Point3d pt1 = new Point3d(plane, newDB.Limmin);
                Point3d pt2 = new Point3d(plane, newDB.Limmax);
                newDB.CloseInput(true);
                var scaleY = Y / Math.Abs(pt1.Y - pt2.Y);
                var mtx = Matrix3d.Scaling(scaleY, pBase);
                return mtx;
            }
        }
        public void UpdateTable<T>(Table table, List<T> list)
        {
            table.RemoveDataLink();
            for (int i = 0; i < table.Columns.Count(); i++)
            {
                //当前列标题
                var curTitle = table.Cells[0, i].GetTextString(FormatOption.IgnoreMtextFormat);
                if (curTitle.Contains("/"))
                {
                    curTitle = curTitle.Replace("/", "_");
                }

                for (int j = 0; j < table.Rows.Count && j < list.Count; j++)
                {
                    T currentT = list[j];
                    PropertyInfo prop = currentT.GetType().GetProperty(curTitle);
                    if (prop != null)
                    {
                        var value = prop.GetValue(currentT, null);
                        if (value == null || value.GetType() == typeof(int) && value.ToString() == "0")
                        {
                            value = "";
                        }
                        if (table.Cells[j + 1, i].GetTextString(FormatOption.IgnoreMtextFormat) == value.ToString()) continue;
                        table.Cells[j + 1, i].SetValue(value.ToString(), ParseOption.ParseOptionNone);
                    }
                }
            }
        }

        public void UpdateTable(Table table, string oldValue, string newValue, bool isContains)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            string str = "";
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    string val = table.Cells[i, j].GetTextString(FormatOption.IgnoreMtextFormat);
                    if (!string.IsNullOrEmpty(val) && ((val.Contains(oldValue) && isContains) || !isContains && val == oldValue))
                    {
                        table.Cells[i, j].TextString = table.Cells[i, j].TextString.Replace(oldValue, newValue);
                    }
                }
            }
            table.DowngradeOpen();
        }

        public void UpdateTableVersion(Table table)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            string str = "";
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string val = table.Cells[i, 4].GetTextString(FormatOption.IgnoreMtextFormat);
                if (!string.IsNullOrEmpty(val))
                {
                    str += val + "\n";
                    table.Cells[i, 4].SetValue(table.Cells[i, 4].TextString.Replace(val, "A"), ParseOption.ParseOptionNone);
                }
            }
            table.DowngradeOpen();
        }
        public void UpdateTableMLCount(Table table, int row, int count)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            string val = table.Cells[row, 4].GetTextString(FormatOption.IgnoreMtextFormat);
            int res = 0;
            int.TryParse(val, out res);
            table.Cells[row, 4].SetValue((res + count).ToString(), ParseOption.ParseOptionNone);
            table.DowngradeOpen();
        }
        public void UpdateTableNP(Table table, int BNCount, int XKcount)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            for (int i = 0; i < BNCount; i += 1)
            {
                table.Cells[130 + 2 * i, 1].SetValue((500 + i) + "BN", ParseOption.ParseOptionNone);
                table.Cells[130 + 2 * i, 2].SetValue("Ⅳ", ParseOption.ParseOptionNone);
                table.Cells[130 + 2 * i, 3].SetValue("1", ParseOption.ParseOptionNone);
            }
            for (int i = 0; i < XKcount; i += 1)
            {
                table.Cells[131 + 2 * i, 1].SetValue((500 + i) + "XK", ParseOption.ParseOptionNone);
                table.Cells[131 + 2 * i, 2].SetValue("Ⅲ", ParseOption.ParseOptionNone);
                table.Cells[131 + 2 * i, 3].SetValue("1", ParseOption.ParseOptionNone);
            }
            table.DowngradeOpen();
        }
        public void UpdateTableNP600(Table table, int BNCount, int XKcount)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            for (int i = 0; i < BNCount; i += 1)
            {
                table.Cells[180 + 2 * i, 1].SetValue((600 + i) + "BN", ParseOption.ParseOptionNone);
                table.Cells[180 + 2 * i, 3].SetValue("1", ParseOption.ParseOptionNone);
                table.Cells[180 + 2 * i, 2].SetValue("Ⅳ", ParseOption.ParseOptionNone);
            }
            for (int i = 0; i < XKcount; i += 1)
            {
                table.Cells[180 + 2 * i + 1, 1].SetValue((600 + i) + "XK", ParseOption.ParseOptionNone);
                table.Cells[180 + 2 * i + 1, 3].SetValue("1", ParseOption.ParseOptionNone);
                table.Cells[180 + 2 * i + 1, 2].SetValue("Ⅲ", ParseOption.ParseOptionNone);
            }
            table.DowngradeOpen();
        }
        public void UpdateTableCabName(Table table)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            string str = "";
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    string val = table.Cells[i, j].GetTextString(FormatOption.IgnoreMtextFormat);
                    if (!string.IsNullOrEmpty(val) && val.Contains("3I"))
                    {
                        table.Cells[i, j].TextString = table.Cells[i, j].TextString.Replace("3I", "4I");
                    }
                }
            }
            table.DowngradeOpen();
        }
      
        public void UpdateTableCabNameForNP(Table table)
        {
            table.UpgradeOpen();
            table.RemoveDataLink();
            table.Cells[1, 1].TextString = table.Cells[1, 1].TextString.Replace("3I", "4I");
            table.DowngradeOpen();
        }
        public void RemoveDBObjectByGroup(string pGroupName)
        {
            try
            {
                var gp = GetGroup(pGroupName);
                if (gp != null)
                {
                    ObjectId[] ObjectIds = gp.GetAllEntityIds();
                    foreach (var itemId in ObjectIds)
                    {
                        var ent = GetDBObject(itemId, OpenMode.ForWrite);
                        ent.Erase();
                        ent.DowngradeOpen();
                    }
                    DelGroup(pGroupName);
                }
            }
            catch (Exception ex)
            {
               
            }
        }

        public void RemoveDBObjectByAllGroup()
        {
            try
            {
                DBDictionary dict = transaction.GetObject(db.GroupDictionaryId, OpenMode.ForWrite, true) as DBDictionary;
                foreach (var item in dict)
                {
                    var gp = GetGroup(item.Key);
                    if (gp != null)
                    {
                        ObjectId[] ObjectIds = gp.GetAllEntityIds();
                        foreach (var itemId in ObjectIds)
                        {
                            var ent = GetDBObject(itemId, OpenMode.ForWrite);
                            ent.Erase();
                            ent.DowngradeOpen();
                        }
                        DelGroup(item.Key);
                    }                    
                }
                dict.DowngradeOpen();
                
            }
            catch (Exception ex)
            {
              
            }
        }
        public List<Entity> GetDBObjectByGroup(string pGroupName)
        {
            List<Entity> entities = new List<Entity>();
            try
            {
                var gp = GetGroup(pGroupName);
                if (gp != null)
                {
                    ObjectId[] ObjectIds = gp.GetAllEntityIds();
                    foreach (var itemId in ObjectIds)
                    {
                        var ent = GetEntity(itemId, OpenMode.ForWrite);
                        entities.Add(ent);                        
                    }                    
                }
            }
            catch (Exception ex)
            {
                
            }
            return entities;
        }
        public bool IsDBObjectInGroup(ObjectId objectId)
        {           
            try
            {
                DBDictionary dict = transaction.GetObject(db.GroupDictionaryId, OpenMode.ForRead, true) as DBDictionary;
                foreach (var item in dict)
                {
                    var gp = GetGroup(item.Key);
                    if (gp != null)
                    {
                        ObjectId[] ObjectIds = gp.GetAllEntityIds();
                        if (ObjectIds.Contains(objectId))
                        {
                            return true;
                        }
                    }
                }                    
            }
            catch
            {
                return false;
            }
            return false;
        }
        public Point3d GetGroupLeftUpPoint(string pGroupName)
        {
            Group gp = GetGroup(pGroupName);
            List<Extents3d> extents3Ds = new List<Extents3d>();
            if (gp != null)
            {
                ObjectId[] ObjectIds = gp.GetAllEntityIds();
                foreach (var itemId in ObjectIds)
                {
                    var ent = GetDBObject(itemId, OpenMode.ForRead) as Entity;
                    Extents3d extents3D = ent.GeometricExtents;
                    extents3Ds.Add(extents3D);
                }
            }
            var maxX = extents3Ds.Min(e => e.MinPoint.X);
            var maxY = extents3Ds.Max(e => e.MaxPoint.Y);
            return new Point3d(maxX, maxY, 0);
        }
        public Point3d GetGroupMidPoint(string pGroupName,bool isLeftOrRight)
        {
            Group gp = GetGroup(pGroupName);
            List<Extents3d> extents3Ds = new List<Extents3d>();
            if (gp != null)
            {
                ObjectId[] ObjectIds = gp.GetAllEntityIds();
                foreach (var itemId in ObjectIds)
                {
                    var ent = GetDBObject(itemId, OpenMode.ForRead) as Entity;
                    Extents3d extents3D = ent.GeometricExtents;
                    extents3Ds.Add(extents3D);
                }
            }
            var maxX = extents3Ds.Min(e => e.MaxPoint.X);
            var minX = extents3Ds.Min(e => e.MinPoint.X);
            var maxY = extents3Ds.Max(e => e.MaxPoint.Y);
            var minY = extents3Ds.Max(e => e.MinPoint.Y);
            if (isLeftOrRight)
            {
                return new Point3d(minX, (maxY + minY) / 2, 0);
            }
            return new Point3d(maxX, (maxY + minY) / 2, 0);
        }
        public Point3d GetGroupMinPoint(string pGroupName)
        {
            Group gp = GetGroup(pGroupName);
            List<Extents3d> extents3Ds = new List<Extents3d>();
            if (gp != null)
            {
                ObjectId[] ObjectIds = gp.GetAllEntityIds();
                foreach (var itemId in ObjectIds)
                {
                    var ent = GetDBObject(itemId, OpenMode.ForRead) as Entity;
                    Extents3d extents3D = ent.GeometricExtents;
                    extents3Ds.Add(extents3D);
                }
            }
            var maxX = extents3Ds.Min(e => e.MinPoint.X);
            var maxY = extents3Ds.Min(e => e.MinPoint.Y);
            return new Point3d(maxX, maxY, 0);
        }
        public Point3d GetGroupMaxPoint(string pGroupName)
        {
            Group gp = GetGroup(pGroupName);
            List<Extents3d> extents3Ds = new List<Extents3d>();
            if (gp != null)
            {
                ObjectId[] ObjectIds = gp.GetAllEntityIds();
                foreach (var itemId in ObjectIds)
                {
                    var ent = GetDBObject(itemId, OpenMode.ForRead) as Entity;
                    Extents3d extents3D = ent.GeometricExtents;
                    extents3Ds.Add(extents3D);
                }
            }
            var maxX = extents3Ds.Max(e => e.MaxPoint.X);
            var maxY = extents3Ds.Max(e => e.MaxPoint.Y);
            return new Point3d(maxX, maxY, 0);
        }
        public void MoveDBObjectByGroup(string pGroupName, Point3d pOrigin, Point3d pEnd)
        {
            try
            {
                Group gp = GetGroup(pGroupName);
                if (gp != null)
                {
                    ObjectId[] ObjectIds = gp.GetAllEntityIds();
                    foreach (var itemId in ObjectIds)
                    {
                        Move(itemId, pOrigin, pEnd);
                    }
                }
            }
            catch (Exception ex)
            {
                
            }
        }
        public void CopyDBObjectByGroup(string pGroupName, Point3d pOrigin, Point3d pEnd, string keywords, string newName, string keyword1 = null, string newName1 = null)
        {
            try
            {
                Group gp = GetGroup(pGroupName);
                if (gp != null)
                {
                    ObjectId[] ObjectIds = gp.GetAllEntityIds();
                    foreach (var itemId in ObjectIds)
                    {
                        var objectid = Copy(itemId, pOrigin, pEnd);
                        var ent = GetDBObject(objectid, OpenMode.ForRead) as MText;
                        if (ent != null && ent.Contents.Contains(keywords))
                        {
                            ent.UpgradeOpen();
                            ent.Contents = ent.Contents.Replace(keywords, newName);
                            ent.DowngradeOpen();
                        }
                        if (ent != null && keyword1 != null && newName1 != null && ent.Contents.Contains(keyword1))
                        {
                            ent.UpgradeOpen();
                            ent.Contents = ent.Contents.Replace(keyword1, newName1);
                            ent.DowngradeOpen();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               
            }
        }

        public List<Entity> CopyDBObjectByGroup(string pGroupName, Point3d pOrigin, Point3d pEnd)
        {
            List<Entity> list = new List<Entity>();
            try
            {
                Group gp = GetGroup(pGroupName);
                if (gp != null)
                {
                    ObjectId[] ObjectIds = gp.GetAllEntityIds();
                    foreach (var itemId in ObjectIds)
                    {
                        var objectid = Copy(itemId, pOrigin, pEnd);
                        var ent = GetEntity(objectid);
                        list.Add(ent);
                    }
                }                
            }
            catch (Exception ex)
            {
              
            }
            return list;
        }
        /// <summary>
        /// 移动实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="sourcePt">移动的源点</param>
        /// <param name="targetPt">移动的目标点</param>
        public void Move(ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于移动实体的矩阵
            Vector3d vector = sourcePt.GetVectorTo(targetPt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            //以写的方式打开id表示的实体对象
            Entity ent = GetDBObject(id, OpenMode.ForWrite) as Entity;
            ent.TransformBy(mt);//对实体实施移动
            ent.DowngradeOpen();//为防止错误，切换实体为读的状态
        }
        /// <summary>
        /// 复制实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="sourcePt">复制的源点</param>
        /// <param name="targetPt">复制的目标点</param>
        /// <returns>返回复制实体的ObjectId</returns>
        public ObjectId Copy(ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于复制实体的矩阵
            Vector3d vector = sourcePt.GetVectorTo(targetPt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            //获取id表示的实体对象
            Entity ent = GetDBObject(id, OpenMode.ForRead) as Entity;
            //获取实体的拷贝
            Entity entCopy = ent.GetTransformedCopy(mt);
            //将复制的实体对象添加到模型空间
            ObjectId copyId = id.Database.AddToModelSpace(entCopy);
            return copyId; //返回复制实体的ObjectId
        }
        public Group GetGroup(string pGroupName)
        {
            try
            {
                Group gp;
                DBDictionary dict = transaction.GetObject(db.GroupDictionaryId, OpenMode.ForWrite, true) as DBDictionary;
                if (dict.Contains(pGroupName))
                {
                    gp = dict.GetAt(pGroupName).GetObject(OpenMode.ForWrite) as Group;
                    return gp;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                editor.WriteMessage("\nerror on creategroup.");
              
                return null;
            }
        }
        public void DelGroup(string pGroupName)
        {
            DBDictionary dict = transaction.GetObject(db.GroupDictionaryId, OpenMode.ForWrite, true) as DBDictionary;
            if (dict.Contains(pGroupName))
            {
                dict.Remove(pGroupName);
            }
            dict.DowngradeOpen();
        }
        public void DelAllGroup()
        {
            DBDictionary dict = transaction.GetObject(db.GroupDictionaryId, OpenMode.ForWrite, true) as DBDictionary;
            foreach (var item in dict)
            {
                dict.Remove(item.Key);
            }           
            dict.DowngradeOpen();
        }
        /// <summary>
        /// 创建组
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="groupName">组名</param>
        /// <param name="ids">要加入实体的ObjectId集合</param>
        /// <returns>返回组的Id</returns>
        public ObjectId CreateGroup(string groupName, ObjectIdList ids)
        {
            //打开当前数据库的组字典对象
            DBDictionary groupDict = (DBDictionary)GetDBObject(db.GroupDictionaryId, OpenMode.ForRead);
            //如果已经存在指定名称的组，则返回
            if (groupDict.Contains(groupName)) return ObjectId.Null;
            //新建一个组对象
            Group group = new Group(groupName, true);
            groupDict.UpgradeOpen(); //切换组字典为写的状态
            //在组字典中加入新创建的组对象，并指定它的搜索关键字为groupName
            groupDict.SetAt(groupName, group);
            //通知事务处理完成组对象的加入
            db.TransactionManager.AddNewlyCreatedDBObject(group, true);
            group.Append(ids); // 在组对象中加入实体对象
            groupDict.DowngradeOpen(); //为了安全，将组字典切换成写
            return group.ObjectId; //返回组的Id
        }
        public string GetPageTypeSummary()
        {
            if (db.HasSummaryInfo())
            {
                if (db.HasCustomProperty("PAGETYPE"))
                {
                    var info = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
                    return info.CustomPropertyTable["PAGETYPE"].ToString();
                }
            }            
            return "未找到PAGETYPE属性！";
        }
        public string? GetDwgSummary(string tag)
        {
            if (db.HasSummaryInfo())
            {
                if (db.HasCustomProperty(tag))
                {
                    var info = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
                    return info.CustomPropertyTable[tag].ToString();
                }
            }
            return null;
        }
        public int GetPageTypeIndexSummary()
        {
            if (db.HasSummaryInfo())
            {
                if (db.HasCustomProperty("PAGETYPEINDEX"))
                {
                    var info = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
                    return int.Parse(info.CustomPropertyTable["PAGETYPEINDEX"].ToString());
                }
            }
            return 0;
        }

        public List<Entity> GetBlockRefEntities(BlockReference blockRef)
        {
            List<Entity> curList = new List<Entity>();
            DBObjectCollection colls = new DBObjectCollection();
            blockRef.Explode(colls);
            foreach (DBObject obj in colls)
            {
                Entity ent = obj as Entity;
                if (ent is AttributeDefinition)
                {
                    curList.Add(ent.Clone() as Entity);
                }
                else
                {
                    curList.Add(ent);
                }
            }
            return curList;
        }
        public ObjectId AttachXref(string fileName, string blockName, ObjectId layerId, Point3d insertionPoint, Scale3d scaleFactors, double rotation, bool isOverlay)
        {
            ObjectId xrefId = ObjectId.Null;//外部参照的Id
            //选择以覆盖的方式插入外部参照
            if (isOverlay) xrefId = db.OverlayXref(fileName, blockName);
            //选择以附着的方式插入外部参照
            else xrefId = db.AttachXref(fileName, blockName);
            //根据外部参照创建一个块参照，并指定其插入点
            BlockReference bref = new BlockReference(insertionPoint, xrefId);
            bref.ScaleFactors = scaleFactors;//外部参照块的缩放因子
            bref.Rotation = rotation;//外部参照块的旋转角度
            bref.SetLayerId(layerId, true);
            var brId = AddToModelSpace(bref);
            return brId;//返回外部参照的Id
        }

        public void UnloadExternalReference()
        {

            BlockTable bt = (BlockTable)transaction.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId id in bt)//遍历块表
            {
                BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(id, OpenMode.ForRead);
                if (btr.IsFromExternalReference && btr.Name == System.IO.Path.GetFileNameWithoutExtension(doc.Name) + "(1)")
                {
                    db.UnloadXrefs(new ObjectIdList(id));
                    break;
                }
            }
        }
        public void BindXref(ObjectId objectId)
        {
            BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(objectId, OpenMode.ForRead);
            db.BindXrefs(new ObjectIdList(objectId), true);
        }
        public void MoveEntityToBottom(params ObjectId[] id)
        {
            BlockTable bt = (BlockTable)this.transaction.GetObject(db.BlockTableId, OpenMode.ForRead);
            //以写方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            DrawOrderTable drawOrderTable = transaction.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
            drawOrderTable.MoveToBottom(new ObjectIdList(id));
          
        }
        public void MoveEntityToTop(params ObjectId[] id)
        {
            BlockTable bt = (BlockTable)this.transaction.GetObject(db.BlockTableId, OpenMode.ForRead);
            //以写方式打开模型空间块表记录.
            BlockTableRecord btr = (BlockTableRecord)this.transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            DrawOrderTable drawOrderTable = transaction.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
            drawOrderTable.MoveToTop(new ObjectIdList(id));
            
        }
        public void DelLayer(string layerName)
        {
            //打开层表
            LayerTable lt = (LayerTable)GetDBObject(db.LayerTableId, OpenMode.ForRead);
            //关闭层表
            foreach (ObjectId id in lt)//遍历层表
            {
                //打开层表记录
                LayerTableRecord ltr = (LayerTableRecord)GetDBObject(id, OpenMode.ForRead);
                if (ltr.Name == layerName)
                {
                    ltr.UpgradeOpen();
                    ltr.Erase();
                    ltr.DowngradeOpen();
                }
            }
        }
        public PromptEntityResult SelectBlockRef(string msg)
        {
            PromptEntityOptions promptEntityOptions = new PromptEntityOptions(msg);
            promptEntityOptions.SetRejectMessage("请选择块参照\n");
            promptEntityOptions.AddAllowedClass(typeof(BlockReference), true);
            var result = editor.GetEntity(promptEntityOptions);
            return result;
        }

        /// <summary>
        /// 获得块中pline和其他实体的交点
        /// </summary>
        /// <returns></returns>
        public Point3d GetFarCrossPInBlockRef(BlockReference br, Line line)
        {
            List<Point3d> point3Ds = new List<Point3d>();
            DBObjectCollection colls = new DBObjectCollection();
            br.Explode(colls);//炸开块           
            foreach (var dbobject in colls)
            {
                if (dbobject is Entity)
                {
                    var en = dbobject as Entity;
                    Point3dCollection pts = new Point3dCollection();
                    try
                    {
                        en.IntersectWith(line, Intersect.OnBothOperands, pts, 0, 0);
                    }
                    catch (Exception e)
                    {

                    }

                    if (pts.Count > 0)
                    {
                        foreach (Point3d p in pts)
                        {
                            point3Ds.Add(p);
                        }
                    }
                }
            }

            if (point3Ds.Count > 0)
            {
                return point3Ds.MinBy(p => p.DistanceTo(line.EndPoint));
            }
            return line.StartPoint;
        }

        public ObjectId AddDataLink(string dlName, string connectString)
        {
            //数据链接管理对象
            DataLinkManager dlm = db.DataLinkManager;
            //判断数据链接是否存在 如果存在移除
            ObjectId dlId = dlm.GetDataLink(dlName);
            if (dlId != ObjectId.Null)
            {
                dlm.RemoveDataLink(dlId);
            }
            //创建并添加新的数据链接
            DataLink dl = new DataLink();
            dl.DataAdapterId = "AcExcel";
            dl.Name = dlName;
            dl.Description = "新加的链接";
            dl.ConnectionString = connectString;
            dl.UpdateOption |= (int)UpdateOption.AllowSourceUpdate;
            dlId = dlm.AddDataLink(dl);
            return dlId;
        }

        public void SetDataLink(Table table, ObjectId dlId)
        {
            table.SetDataLink(table.Cells, dlId, false);
        }
        public void UpdateTableByDatalink(Table table)
        {
            ObjectId dlId = table.GetDataLink(0, 0);
            DataLink dl = GetDBObject(dlId, OpenMode.ForWrite) as DataLink;
            //更新数据
            dl.Update(UpdateDirection.SourceToData, UpdateOption.None);
            //从链接更新到表格
            table.UpdateDataLink(UpdateDirection.SourceToData, UpdateOption.None);
            dl.UpgradeOpen();
            table.DowngradeOpen();
        }
        public void UpdateExcelByDataLink(Table table, ObjectId dlId)
        {
            //从CAD表格数据更新链接
            table.UpdateDataLink(UpdateDirection.DataToSource, UpdateOption.ForceFullSourceUpdate);
            //由链接数据更新Excel表格           
            DataLink dl = GetDBObject(dlId, OpenMode.ForWrite) as DataLink;
            dl.Update(UpdateDirection.DataToSource, UpdateOption.ForceFullSourceUpdate);
            dl.DowngradeOpen();
        }
        /// <summary>
        /// 求交点
        /// </summary>
        /// <param name="lineOrPline"></param>
        /// <returns></returns>
        public List<Point3d> GetCrossPoint(Line line, ObjectId selfObjectId)
        {

            //对于平行于X轴的直线 求交点 并且把交点保存到xdata
            var linesy = GetAllEntity<Line>();
            //求与X轴平行的线 断横           
            if (linesy.Count > 0)
            {
                var vec = line.StartPoint.GetVectorTo(line.EndPoint);
                linesy = linesy.Where(l => !l.StartPoint.GetVectorTo(l.EndPoint).IsParallelTo(vec)).ToList();
            }
            var plines = GetAllEntity<Polyline>();

            //求和其他直线或者多段线的交点
            List<Point3d> ptsAll = new List<Point3d>();
            var curves = linesy.Cast<Curve>().Union(plines.Cast<Curve>()).Where(l => l.ObjectId != selfObjectId);
            foreach (var curve in curves)
            {
                Point3dCollection pts = new Point3dCollection();
                var p1 = line.StartPoint;
                var p2 = line.EndPoint;
                var p3 = curve.StartPoint;
                var p4 = curve.EndPoint;
                line.IntersectWith(curve, Intersect.OnBothOperands, pts, 0, 0);//求交点
                if (pts.Count > 0)
                {
                    foreach (Point3d pt in pts)
                    {
                        ptsAll.Add(pt);
                    }
                }
            }
            ptsAll = ptsAll.OrderBy(p => p.X).ToList();
            return ptsAll;
        }

        public List<Line> GetSplitCurves(Polyline pline)
        {
            List<Line> lines = new List<Line>();
            var pts = new Point3dCollection();
            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                pts.Add(pline.GetPoint3dAt(i));
            }
            DBObjectCollection dBObjects = pline.GetSplitCurves(pts);
            foreach (Polyline line in dBObjects)
            {
                lines.Add(new Line(line.StartPoint, line.EndPoint));
            }
            return lines;
        }
        /// <summary>
        /// 炸开块
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="brs"></param>
        /// <returns></returns>
        public List<T> ExplodeBlockRef<T>(List<BlockReference> brs, string keyWord) where T : DBObject
        {

            List<T> ts = new List<T>();
            var handleBrs = brs.Where(b => b.ObjectId != ObjectId.Null && b.Name.Contains(keyWord));
            if (handleBrs == null || handleBrs.Count() == 0) return ts;
            foreach (var br in handleBrs)
            {
                DBObjectCollection dBObjectCollection = new DBObjectCollection();
                br.Explode(dBObjectCollection);
                foreach (DBObject item in dBObjectCollection)
                {
                    if (item is T)
                    {
                        ts.Add((T)item);
                    }
                }
            }
            return ts;
        }
        /// <summary>
        /// 炸开块
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="brs"></param>
        /// <returns></returns>
        public List<T> ExplodeBlockRef<T>(BlockReference br) where T : DBObject
        {

            List<T> ts = new List<T>();
            DBObjectCollection dBObjectCollection = new DBObjectCollection();
            br.Explode(dBObjectCollection);
            foreach (DBObject item in dBObjectCollection)
            {
                if (item is T)
                {
                    ts.Add((T)item);
                }
            }
            return ts;
        }
        public ObjectId InsertAttributeReference(Point3d p, string tagName)
        {
            var currentStyle = db.GetTextStyle("CNCS");
            currentStyle.SetTextStyleProp(3.5, 0.7, 0, false, false, false, AnnotativeStates.False, false);
            AttributeDefinition attr = new AttributeDefinition(p, "", tagName, "", currentStyle);
            attr.HorizontalMode = TextHorizontalMode.TextCenter;
            attr.VerticalMode = TextVerticalMode.TextBottom;
            attr.AlignmentPoint = p;
            attr.WidthFactor = 0.7;
            attr.Height = 3.5;
            var objectId = AddToModelSpace(attr);
            return objectId;
        }
        public  AttributeDefinition CreateAttributeDefination(Point3d p, string tagName)
        {
            bool Invisible = true;
            if (tagName.Contains("LEFT_NO") || tagName.Contains("RIGHT_NO"))
            {
                Invisible = false;
            }
            var currentStyle = db.GetTextStyle("CNCS");
            currentStyle.SetTextStyleProp(2.5, 0.7, 0, false, false, false, AnnotativeStates.False, false);
            AttributeDefinition attr = new AttributeDefinition
            {
                Invisible = Invisible,
                Tag = tagName,
                HorizontalMode = TextHorizontalMode.TextMid,
                VerticalMode = TextVerticalMode.TextBase,
                AlignmentPoint = p,
                WidthFactor = 0.7,
                //TextStyleId = currentStyle
            };                   
            return attr;
        }
        public Extents3d GetAllEntsExtent()
        {
            Extents3d totalExt = new Extents3d();
            try
            {
                BlockTable bt = (BlockTable)GetDBObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btRcd = (BlockTableRecord)GetDBObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId entId in btRcd)
                {
                    Entity ent = GetEntity(entId, OpenMode.ForRead) as Entity;
                    totalExt.AddExtents(ent.GeometricExtents);
                }
               
            }
            catch (Exception e)
            {

            }
            return totalExt;
        }
        /// <summary>
        /// 确保布局中的图纸显示在布局中的中间，而不需要使用缩放命令来显示
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layoutName">布局的名称</param>
        public  void CenterLayoutViewport(ObjectIdCollection eraseId, string layoutName)
        {
            Extents3d ext = GetAllEntsExtent();
            BlockTable bt = this.db.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
            foreach (ObjectId btrId in bt)
            {
                BlockTableRecord btr = btrId.GetObject(OpenMode.ForRead) as BlockTableRecord;
                if (btr.IsLayout)
                {
                    Layout layout = btr.LayoutId.GetObject(OpenMode.ForRead) as Layout;
                    if (layout.LayoutName.CompareTo(layoutName) == 0)
                    {
                        int vpIndex = 0;
                        ObjectId firstViewportId = new ObjectId();
                        ObjectId secondViewportId = new ObjectId();
                        foreach (ObjectId entId in btr)
                        {
                            if (eraseId.Contains(entId)) continue;
                            Entity ent = entId.GetObject(OpenMode.ForWrite) as Entity;
                            if (ent is Viewport)
                            {
                                Viewport vp = ent as Viewport;
                                if (vpIndex == 0)
                                {
                                    firstViewportId = entId;
                                    vpIndex++;
                                }
                                else if (vpIndex == 1)
                                {
                                    secondViewportId = entId;
                                }
                            }
                        }
                        //布局复制过来之后得到了两个视口，第一个视口与模型空间关联，第二个视口则是在正确的位置上
                        if (secondViewportId.IsValid && firstViewportId.IsValid)
                        {
                            Viewport secondVP = secondViewportId.GetObject(OpenMode.ForWrite) as Viewport;
                            secondVP.ColorIndex = 1;
                            secondVP.Erase();
                            Viewport firstVP = firstViewportId.GetObject(OpenMode.ForWrite) as Viewport;
                            firstVP.CenterPoint = secondVP.CenterPoint;
                            firstVP.Height = secondVP.Height;
                            firstVP.Width = secondVP.Width;
                            firstVP.ColorIndex = 5;
                            Point3d midPt = GeTools.MidPoint(ext.MinPoint, ext.MaxPoint);
                            firstVP.ViewCenter = new Point2d(midPt.X, midPt.Y); ;
                            double xScale = secondVP.Width / ((ext.MaxPoint.X - ext.MinPoint.X) * 1.1);
                            double yScale = secondVP.Height / ((ext.MaxPoint.Y - ext.MinPoint.Y) * 1.1);
                            firstVP.CustomScale = Math.Min(xScale, yScale);
                            firstVP.StandardScale = StandardScaleType.ScaleToFit;
                            secondVP.DowngradeOpen();
                            firstVP.DowngradeOpen();                      
                        }
                    }
                }
            }


        }

        public Point3d GetMidPoint(Entity entity)
        {
            var p1 = entity.GeometricExtents.MaxPoint;
            var p2 = entity.GeometricExtents.MinPoint;
            var p = new Point3d((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2, 0);
            return p;
        }
        public Point3d GetMidLeftPoint(Entity entity)
        {
            var p1 = entity.GeometricExtents.MaxPoint;
            var p2 = entity.GeometricExtents.MinPoint;
            var p = new Point3d(p2.X, (p1.Y + p2.Y) / 2, 0);
            return p;
        }
        public Point3d GetMidRightPoint(Entity entity)
        {
            var p1 = entity.GeometricExtents.MaxPoint;
            var p2 = entity.GeometricExtents.MinPoint;
            var p = new Point3d(p1.X, (p1.Y + p2.Y) / 2, 0);
            return p;
        }
        public  void DelAllLayouts(string layoutName)
        {
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            foreach (ObjectId id in bt)
            {
                BlockTableRecord btr = (BlockTableRecord)id.GetObject(OpenMode.ForRead);
                if (btr.IsLayout && btr.Name.ToUpper() != BlockTableRecord.ModelSpace.ToUpper())
                {
                    Layout layout = (Layout)btr.LayoutId.GetObject(OpenMode.ForRead);
                    if (layout.LayoutName != layoutName)
                    {
                        LayoutManager lm = LayoutManager.Current;                    
                        if (btr.LayoutId != ObjectId.Null) lm.DeleteLayout(layout.LayoutName);
                    }

                }
            }
        }
        public string GetTextStyle(string styleName)
        {
            TextStyleTable st = (TextStyleTable)db.TextStyleTableId.GetObject(OpenMode.ForRead);
            foreach (ObjectId id in st)
            {
                TextStyleTableRecord btr = (TextStyleTableRecord)id.GetObject(OpenMode.ForRead);
                if (btr.Name == styleName)
                {
                    return btr.FileName;
                }
            }
            return "";
        }
        public string GetTextStyleName(ObjectId textId)
        {
            TextStyleTable st = (TextStyleTable)db.TextStyleTableId.GetObject(OpenMode.ForRead);
            foreach (ObjectId id in st)
            {
                
                if (id == textId)
                {
                    TextStyleTableRecord btr = (TextStyleTableRecord)id.GetObject(OpenMode.ForRead);
                    return btr.Name;
                }
            }
            return "";
        }

        public ObjectId InsertImage(string file)
        {
            ObjectId dictId = RasterImageDef.GetImageDictionary(db);
            if (dictId.IsNull) dictId = RasterImageDef.CreateImageDictionary(db);
            DBDictionary dict = (DBDictionary)GetDBObject(dictId, OpenMode.ForRead);
            RasterImageDef def = new RasterImageDef();
            def.SourceFileName = file;
            def.Load();
            dict.UpgradeOpen();
            ObjectId defid = dict.SetAt(Guid.NewGuid().ToString("N"), def);
            transaction.AddNewlyCreatedDBObject(def, true);
            RasterImage image = new RasterImage();
            image.ImageDefId = defid;
            image.ShowImage = true;
            //image.Orientation = new CoordinateSystem3d(Point3d.Origin, new Vector3d(image.Width, 0, 0), new Vector3d(0, image.Height, 0));
            var ObjectId = db.AddToModelSpace(image);
            RasterImage.EnableReactors(true);
            image.AssociateRasterDef(def);
            return ObjectId;
        }
        public Point3d GetGroupOriginPoint(string groupName)
        {
            //找到组
            var entities = GetDBObjectByGroup(groupName);
            var br = entities.FirstOrDefault(e => e is BlockReference) as BlockReference;
            return br.Position;
        }
        public void ReplaceText(string oldValue, string newValue)
        {
            var mtexts = GetAllEntity<MText>();
            foreach (var mtext in mtexts)
            {               
                mtext.UpgradeOpen();
                mtext.Contents = mtext.Contents.Replace(oldValue, newValue);
                mtext.DowngradeOpen();
            }           
        }
        public void ReplaceTextExceptGroup(string oldValue, string newValue)
        {
            var mtexts = GetAllEntity<MText>();
            foreach (var mtext in mtexts)
            {
                if (IsDBObjectInGroup(mtext.ObjectId))
                {
                    continue;
                }
                mtext.UpgradeOpen();
                mtext.Contents = mtext.Contents.Replace(oldValue, newValue);
                mtext.DowngradeOpen();
            }
        }
        public void ReplaceText(MText mText, string oldValue, string newValue)
        {           
            mText.UpgradeOpen();
            mText.Contents = mText.Contents.Replace(oldValue, newValue);
            mText.DowngradeOpen();
        }

        public Point3d GetGroupLeftTopEntityPoint(string groupName)
        {
            var results = new List<Point3d>();
            var ents = GetDBObjectByGroup(groupName);
            foreach (var item in ents)
            {
                if (item is Line)
                {
                    var l = item as Line;
                    results.Add(l.StartPoint);
                    results.Add(l.EndPoint);
                }
                if (item is Polyline)
                {
                    var l = item as Polyline;
                    results.Add(l.StartPoint);
                    results.Add(l.EndPoint);
                }
            }
            return results.OrderBy(r => r.X).MaxBy(r => r.Y);
        }
   

        /// <summary>
        /// 根据布局名称和块名称获取当前布局中所有匹配块名的块参照
        /// </summary>
        /// <param name="layoutName">布局名称</param>
        /// <param name="blockName">块名称</param>
        /// <returns>返回指定布局中所有匹配块名的块参照列表</returns>
        public List<BlockReference> GetBlocksInLayout(string layoutName, string blockName)
        {
            List<BlockReference> blocks = new List<BlockReference>();

            // 获取布局空间的ID
            ObjectId layoutSpaceId = ObjectId.Null;

            // 打开布局
            DBDictionary layoutDict = (DBDictionary)transaction.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            if (layoutDict.Contains(layoutName))
            {
                Layout layout = (Layout)transaction.GetObject(layoutDict.GetAt(layoutName), OpenMode.ForRead);
                layoutSpaceId = layout.BlockTableRecordId;
            }
            else
            {
                // 如果没有找到指定的布局，返回空列表
                return blocks;
            }

            // 以读方式打开布局空间块表记录
            BlockTableRecord layoutSpace = (BlockTableRecord)transaction.GetObject(layoutSpaceId, OpenMode.ForRead);

            // 获取所有实体
            foreach (ObjectId id in layoutSpace)
            {
                if (id.IsNull || id.IsErased)
                    continue;

                // 尝试获取块参照
                Entity ent = transaction.GetObject(id, OpenMode.ForRead) as Entity;
                BlockReference blockRef = ent as BlockReference;

                if (blockRef != null && blockRef.Name == blockName)
                {
                    blocks.Add(blockRef);
                }
            }
            return blocks;
        }
    }

   

}
