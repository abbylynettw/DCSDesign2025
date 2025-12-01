using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

namespace WinformUI.CADHelper
{
    public class PlotSettingsEx : PlotSettings
    {
        private PlotSettingsValidator validator = PlotSettingsValidator.Current;

        public PlotSettingsEx(PlotSettings ps):base(ps.ModelType)
        {
            this.CopyFrom(ps);//从已有设置中获取打印设置
            validator.RefreshLists(this);//更新打印设备、图纸尺寸和打印样式表信息            
        }

       public List<string> DeviceList
        {
            get
            {
                StringCollection deviceCollection = validator.GetPlotDeviceList();
                List<string> deviceList = new List<string>();
                foreach (var device in deviceCollection)
                {
                    deviceList.Add(device);
                }
                return deviceList;
            }
        }

        public string PlotPaperUnitsLocal
        {
            get
            {
                string plotUnitLocal = "";
                switch (base.PlotPaperUnits)
                {
                    case PlotPaperUnit.Pixels: plotUnitLocal = "英寸"; break;
                    case PlotPaperUnit.Millimeters: plotUnitLocal = "毫米"; break;
                    case PlotPaperUnit.Inches: plotUnitLocal = "像素"; break;
                    default:
                        break;
                }
                return plotUnitLocal;
            }
            set
            {
                if (value != PlotPaperUnitsLocal)
                {
                    switch (value)
                    {
                        case "英寸": validator.SetPlotPaperUnits(this, PlotPaperUnit.Inches); break;
                        case "毫米": validator.SetPlotPaperUnits(this, PlotPaperUnit.Millimeters); break;
                        case "像素": validator.SetPlotPaperUnits(this, PlotPaperUnit.Pixels); break;
                        default:
                            break;
                    }
                }
            }
        }

        public bool IsPlotToFile
        {
            get
            {
                //如果未选择打印设备，则直接返回
                if (PlotConfigurationName == "无") return false;
                //获取当前打印配置，并返回其是否打印到文件属性
                PlotConfig config = PlotConfigManager.CurrentConfig;
                return config.IsPlotToFile;
            }
        }
        public bool NoPlotToFile
        {
            get
            {
                //如果未选择打印设备，则直接返回
                if (PlotConfigurationName == "无") return false;
                //获取当前打印配置，并返回其是否必须打印到文件
                PlotConfig config = PlotConfigManager.CurrentConfig;
                if (config.PlotToFileCapability == PlotToFileCapability.MustPlotToFile)
                    return false;
                else return true;
            }
        }

        public void UpdatePlotSettings(ObjectId psId)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (doc.LockDocument())
            using (Transaction trans=doc.TransactionManager.StartTransaction())
            {
                PlotSettings ps = psId.GetObject(OpenMode.ForWrite) as PlotSettings;
                if (ps != null)
                {
                    ps.CopyFrom(this);//复制当前打印设置
                }
                trans.Commit();
            }
        }
    }
}
