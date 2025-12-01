using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace WinformUI.CADHelper
{
    /// <summary>
    /// 视口操作类
    /// </summary>
    public static class ViewportTools
    {
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl, EntryPoint = "?acedSetCurrentVPort@@YA?AW4ErrorStatus@Acad@@PBVAcDbViewport@@@Z")]
        extern static private int acedSetCurrentVPort(IntPtr AcDbVport);
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl, EntryPoint = "?acedSetCurrentVPort@@YA?AW4ErrorStatus@Acad@@H@Z")]
        extern static private int acedSetCurrentVPort(int vpnumber);
        /// <summary>
        /// 将视口置为当前
        /// </summary>
        /// <param name="editor">命令行对象</param>
        /// <param name="vport">视口对象</param>
        public static void SetCurrentVPort(this Editor editor, Viewport vport)
        {
            
            acedSetCurrentVPort(vport.UnmanagedObject);
        }

        /// <summary>
        /// 将视口置为当前
        /// </summary>
        /// <param name="editor">命令行对象</param>
        /// <param name="vportNumber">视口编号</param>
        public static void SetCurrentVPort(this Editor editor, int vportNumber)
        {
            acedSetCurrentVPort(vportNumber);
        }

        /// <summary>
        /// 获取当前活动视口
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回当前活动视口的Id</returns>
        public static ObjectId CurrentViewportTableRecordId(this Database db,DrawingUtility util)
        {
            ObjectId vtrId=ObjectId.Null;
            ViewportTable vt=(ViewportTable)db.ViewportTableId.GetObject(OpenMode.ForRead);
            foreach (ObjectId id in vt)
            {
                if (!id.IsErased)
                {
                    var ent = util.GetDBObject(id, OpenMode.ForRead) as Viewport;
                    if (ent == null) continue;
                    ent.UpgradeOpen();
                    ent.StandardScale = StandardScaleType.ScaleToFit;
                    ent.DowngradeOpen();
                    vtrId = id;                  
                }
            }
            return vtrId;
        }
        public static void SetCurrentViewPortFit(this Database db)
        {
            var viewPorts = db.GetEntsInPaperSpace<Viewport>();
            foreach (var viewport in viewPorts)
            {
               
                if (viewport.Locked)
                {
                    viewport.UpgradeOpen();
                    viewport.Locked = false;
                    viewport.StandardScale = StandardScaleType.ScaleToFit;
                    viewport.Locked = true;
                    viewport.DowngradeOpen();
                    viewport.UpdateDisplay();
                }
                else
                {
                    viewport.UpgradeOpen();
                    viewport.StandardScale = StandardScaleType.ScaleToFit;
                    viewport.UpdateDisplay();
                }                   
                           
            }
        }
        /// <summary>
        /// 获取当前活动视口
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回当前活动视口的Id</returns>
       public static ObjectId CurrentViewportTableRecordId(this Database db)
        {
            ObjectId vtrId = ObjectId.Null;
            ViewportTable vt = (ViewportTable)db.ViewportTableId.GetObject(OpenMode.ForRead);
            foreach (ObjectId id in vt)
            {
                if (!id.IsErased)
                {                  
                    vtrId = id;
                    break;
                }
            }
            return vtrId;
        }

        /// <summary>
        /// 获取当前活动视口
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回当前活动视口的Id</returns>
        public static ObjectId GetCurrentViewportTableRecord(this Database db)
        {
            ObjectId vtrId = ObjectId.Null;
            ViewportTable vt = (ViewportTable)db.ViewportTableId.GetObject(OpenMode.ForRead);
            foreach (ObjectId id in vt)
            {
                if (!id.IsErased)
                {
                    vtrId = id;
                    break;
                }
            }
            return vtrId;
        }
    }
}
