using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace WinformUI.BlockReplace
{
    /// <summary>
    /// 块选择对话框
    /// </summary>
    public partial class BlockSelectionForm : Form
    {
        #region 私有字段

        private List<string> _blockNames;
        private string _selectedBlockName;

        #endregion

        #region 属性

        /// <summary>
        /// 选择的块名称
        /// </summary>
        public string SelectedBlockName => _selectedBlockName;

        #endregion

        #region 构造函数

        public BlockSelectionForm(List<string> blockNames, string dwgFilePath)
        {
            InitializeComponent();
            _blockNames = blockNames ?? new List<string>();
            InitializeUI(dwgFilePath);
        }

        #endregion

        #region 初始化方法

        private void InitializeUI(string dwgFilePath)
        {
            this.Text = "选择块定义";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // 创建控件
            var lblInfo = new Label
            {
                Text = $"在文件中找到以下块定义，请选择一个:\n{dwgFilePath}",
                Location = new Point(12, 12),
                Size = new Size(460, 40),
                AutoSize = false
            };

            var listBoxBlocks = new ListBox
            {
                Name = "listBoxBlocks",
                Location = new Point(12, 60),
                Size = new Size(460, 250),
                SelectionMode = SelectionMode.One
            };

            // 添加块名称到列表
            foreach (var blockName in _blockNames.OrderBy(x => x))
            {
                listBoxBlocks.Items.Add(blockName);
            }

            // 默认选择第一个
            if (listBoxBlocks.Items.Count > 0)
            {
                listBoxBlocks.SelectedIndex = 0;
            }

            var btnOK = new Button
            {
                Text = "确定",
                Location = new Point(316, 325),
                Size = new Size(75, 25),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(397, 325),
                Size = new Size(75, 25),
                DialogResult = DialogResult.Cancel
            };

            // 双击选择
            listBoxBlocks.MouseDoubleClick += (s, e) =>
            {
                if (listBoxBlocks.SelectedItem != null)
                {
                    _selectedBlockName = listBoxBlocks.SelectedItem.ToString();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            };

            // 确定按钮事件
            btnOK.Click += (s, e) =>
            {
                if (listBoxBlocks.SelectedItem != null)
                {
                    _selectedBlockName = listBoxBlocks.SelectedItem.ToString();
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    MessageBox.Show("请选择一个块定义！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            // 添加控件到窗体
            this.Controls.AddRange(new Control[] { lblInfo, listBoxBlocks, btnOK, btnCancel });

            // 设置默认按钮
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        #endregion
    }
}