
using System;
using System.ComponentModel;
using WinformUI;
using WinformUI.UpdateFrameTag;
using WinformUI.UpdateTable;
using static System.Net.Mime.MediaTypeNames;


namespace TestUI
{
    internal class Program
    {            
        [STAThread]
        static void Main(string[] args)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            // 创建窗体实例
            BatchUpgradeProjectForm form = new BatchUpgradeProjectForm();
            form.ShowDialog();

            BatchUpdateTableForm form1 = new BatchUpdateTableForm();

            //ReadExcelModifyCAD form2 = new ReadExcelModifyCAD();

            // 显示窗体 - 对于控制台应用程序，使用ShowDialog()会更合适
            form1.ShowDialog();

            // 如果需要在窗体关闭后执行其他控制台操作
            //Console.WriteLine("窗体已关闭");            
            // Console.ReadLine(); // 可选：如果需要防止控制台立即关闭
        }
    }
}