using System.Drawing;
using System.Windows.Forms;

namespace WinformUI.UpdateTable
{
    partial class BatchUpdateTableForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BatchUpdateTableForm));
            this.mainSplitContainer = new System.Windows.Forms.SplitContainer();
            this.configPanel = new System.Windows.Forms.Panel();
            this.mappingGroupBox = new System.Windows.Forms.GroupBox();
            this.gridMappings = new System.Windows.Forms.DataGridView();
            this.EnabledColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.DwgKeywordColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TableTitleColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExcelSheetColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mappingButtonPanel = new System.Windows.Forms.Panel();
            this.checkSelectAll = new System.Windows.Forms.CheckBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.excelGroupBox = new System.Windows.Forms.GroupBox();
            this.excelTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this.txtFolderPath = new System.Windows.Forms.TextBox();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.controlPanel = new System.Windows.Forms.Panel();
            this.logGroupBox = new System.Windows.Forms.GroupBox();
            this.txtLog = new System.Windows.Forms.RichTextBox();
            this.updateGroupBox = new System.Windows.Forms.GroupBox();
            this.updateTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this.radioExcelToCAD = new System.Windows.Forms.RadioButton();
            this.radioCADToExcel = new System.Windows.Forms.RadioButton();
            this.checkBackup = new System.Windows.Forms.CheckBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.checkMergeCells = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.mainSplitContainer)).BeginInit();
            this.mainSplitContainer.Panel1.SuspendLayout();
            this.mainSplitContainer.Panel2.SuspendLayout();
            this.mainSplitContainer.SuspendLayout();
            this.configPanel.SuspendLayout();
            this.mappingGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridMappings)).BeginInit();
            this.mappingButtonPanel.SuspendLayout();
            this.excelGroupBox.SuspendLayout();
            this.excelTableLayout.SuspendLayout();
            this.controlPanel.SuspendLayout();
            this.logGroupBox.SuspendLayout();
            this.updateGroupBox.SuspendLayout();
            this.updateTableLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainSplitContainer
            // 
            resources.ApplyResources(this.mainSplitContainer, "mainSplitContainer");
            this.mainSplitContainer.Name = "mainSplitContainer";
            // 
            // mainSplitContainer.Panel1
            // 
            this.mainSplitContainer.Panel1.Controls.Add(this.configPanel);
            // 
            // mainSplitContainer.Panel2
            // 
            this.mainSplitContainer.Panel2.Controls.Add(this.controlPanel);
            // 
            // configPanel
            // 
            this.configPanel.Controls.Add(this.mappingGroupBox);
            this.configPanel.Controls.Add(this.excelGroupBox);
            resources.ApplyResources(this.configPanel, "configPanel");
            this.configPanel.Name = "configPanel";
            // 
            // mappingGroupBox
            // 
            this.mappingGroupBox.Controls.Add(this.gridMappings);
            this.mappingGroupBox.Controls.Add(this.mappingButtonPanel);
            resources.ApplyResources(this.mappingGroupBox, "mappingGroupBox");
            this.mappingGroupBox.Name = "mappingGroupBox";
            this.mappingGroupBox.TabStop = false;
            // 
            // gridMappings
            // 
            this.gridMappings.AllowUserToAddRows = false;
            this.gridMappings.AllowUserToDeleteRows = false;
            this.gridMappings.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridMappings.BackgroundColor = System.Drawing.SystemColors.Window;
            resources.ApplyResources(this.gridMappings, "gridMappings");
            this.gridMappings.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.EnabledColumn,
            this.DwgKeywordColumn,
            this.TableTitleColumn,
            this.ExcelSheetColumn});
            this.gridMappings.MultiSelect = false;
            this.gridMappings.Name = "gridMappings";
            this.gridMappings.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            // 
            // EnabledColumn
            // 
            this.EnabledColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            resources.ApplyResources(this.EnabledColumn, "EnabledColumn");
            this.EnabledColumn.Name = "EnabledColumn";
            // 
            // DwgKeywordColumn
            // 
            resources.ApplyResources(this.DwgKeywordColumn, "DwgKeywordColumn");
            this.DwgKeywordColumn.Name = "DwgKeywordColumn";
            // 
            // TableTitleColumn
            // 
            resources.ApplyResources(this.TableTitleColumn, "TableTitleColumn");
            this.TableTitleColumn.Name = "TableTitleColumn";
            // 
            // ExcelSheetColumn
            // 
            resources.ApplyResources(this.ExcelSheetColumn, "ExcelSheetColumn");
            this.ExcelSheetColumn.Name = "ExcelSheetColumn";
            // 
            // mappingButtonPanel
            // 
            this.mappingButtonPanel.Controls.Add(this.checkSelectAll);
            this.mappingButtonPanel.Controls.Add(this.btnDelete);
            this.mappingButtonPanel.Controls.Add(this.btnAdd);
            this.mappingButtonPanel.Controls.Add(this.btnExport);
            this.mappingButtonPanel.Controls.Add(this.btnImport);
            resources.ApplyResources(this.mappingButtonPanel, "mappingButtonPanel");
            this.mappingButtonPanel.Name = "mappingButtonPanel";
            // 
            // checkSelectAll
            // 
            resources.ApplyResources(this.checkSelectAll, "checkSelectAll");
            this.checkSelectAll.Name = "checkSelectAll";
            this.checkSelectAll.UseVisualStyleBackColor = true;
            this.checkSelectAll.CheckedChanged += new System.EventHandler(this.checkSelectAll_CheckedChanged);
            // 
            // btnDelete
            // 
            resources.ApplyResources(this.btnDelete, "btnDelete");
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnAdd
            // 
            resources.ApplyResources(this.btnAdd, "btnAdd");
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnExport
            // 
            resources.ApplyResources(this.btnExport, "btnExport");
            this.btnExport.Name = "btnExport";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            resources.ApplyResources(this.btnImport, "btnImport");
            this.btnImport.Name = "btnImport";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // excelGroupBox
            // 
            this.excelGroupBox.Controls.Add(this.excelTableLayout);
            resources.ApplyResources(this.excelGroupBox, "excelGroupBox");
            this.excelGroupBox.Name = "excelGroupBox";
            this.excelGroupBox.TabStop = false;
            // 
            // excelTableLayout
            // 
            resources.ApplyResources(this.excelTableLayout, "excelTableLayout");
            this.excelTableLayout.Controls.Add(this.txtFolderPath, 0, 0);
            this.excelTableLayout.Controls.Add(this.btnBrowseExcel, 1, 0);
            this.excelTableLayout.Name = "excelTableLayout";
            // 
            // txtFolderPath
            // 
            resources.ApplyResources(this.txtFolderPath, "txtFolderPath");
            this.txtFolderPath.Name = "txtFolderPath";
            // 
            // btnBrowseExcel
            // 
            resources.ApplyResources(this.btnBrowseExcel, "btnBrowseExcel");
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // controlPanel
            // 
            this.controlPanel.Controls.Add(this.logGroupBox);
            this.controlPanel.Controls.Add(this.updateGroupBox);
            resources.ApplyResources(this.controlPanel, "controlPanel");
            this.controlPanel.Name = "controlPanel";
            // 
            // logGroupBox
            // 
            this.logGroupBox.Controls.Add(this.txtLog);
            resources.ApplyResources(this.logGroupBox, "logGroupBox");
            this.logGroupBox.Name = "logGroupBox";
            this.logGroupBox.TabStop = false;
            // 
            // txtLog
            // 
            this.txtLog.BackColor = System.Drawing.Color.White;
            resources.ApplyResources(this.txtLog, "txtLog");
            this.txtLog.Name = "txtLog";
            // 
            // updateGroupBox
            // 
            this.updateGroupBox.Controls.Add(this.updateTableLayout);
            resources.ApplyResources(this.updateGroupBox, "updateGroupBox");
            this.updateGroupBox.Name = "updateGroupBox";
            this.updateGroupBox.TabStop = false;
            // 
            // updateTableLayout
            // 
            resources.ApplyResources(this.updateTableLayout, "updateTableLayout");
            this.updateTableLayout.Controls.Add(this.radioExcelToCAD, 0, 0);
            this.updateTableLayout.Controls.Add(this.radioCADToExcel, 1, 0);
            this.updateTableLayout.Controls.Add(this.checkBackup, 2, 0);
            this.updateTableLayout.Controls.Add(this.btnUpdate, 3, 0);
            this.updateTableLayout.Controls.Add(this.checkMergeCells, 0, 1);
            this.updateTableLayout.Name = "updateTableLayout";
            // 
            // radioExcelToCAD
            // 
            this.radioExcelToCAD.Checked = true;
            resources.ApplyResources(this.radioExcelToCAD, "radioExcelToCAD");
            this.radioExcelToCAD.Name = "radioExcelToCAD";
            this.radioExcelToCAD.TabStop = true;
            // 
            // radioCADToExcel
            // 
            resources.ApplyResources(this.radioCADToExcel, "radioCADToExcel");
            this.radioCADToExcel.Name = "radioCADToExcel";
            this.radioCADToExcel.TabStop = true;
            // 
            // checkBackup
            // 
            resources.ApplyResources(this.checkBackup, "checkBackup");
            this.checkBackup.Name = "checkBackup";
            this.checkBackup.UseVisualStyleBackColor = true;
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            resources.ApplyResources(this.btnUpdate, "btnUpdate");
            this.btnUpdate.ForeColor = System.Drawing.Color.White;
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.UseVisualStyleBackColor = false;
            // 
            // checkMergeCells
            // 
            resources.ApplyResources(this.checkMergeCells, "checkMergeCells");
            this.checkMergeCells.Name = "checkMergeCells";
            this.checkMergeCells.UseVisualStyleBackColor = true;
            // 
            // BatchUpdateTableForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.mainSplitContainer);
            this.Name = "BatchUpdateTableForm";
            this.mainSplitContainer.Panel1.ResumeLayout(false);
            this.mainSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.mainSplitContainer)).EndInit();
            this.mainSplitContainer.ResumeLayout(false);
            this.configPanel.ResumeLayout(false);
            this.mappingGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridMappings)).EndInit();
            this.mappingButtonPanel.ResumeLayout(false);
            this.mappingButtonPanel.PerformLayout();
            this.excelGroupBox.ResumeLayout(false);
            this.excelTableLayout.ResumeLayout(false);
            this.excelTableLayout.PerformLayout();
            this.controlPanel.ResumeLayout(false);
            this.logGroupBox.ResumeLayout(false);
            this.updateGroupBox.ResumeLayout(false);
            this.updateTableLayout.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        // 公共控件成员，供后端代码使用
        private SplitContainer mainSplitContainer;
        private Panel configPanel;
        private GroupBox excelGroupBox;
        private TableLayoutPanel excelTableLayout;
        private TextBox txtFolderPath;
        private Button btnBrowseExcel;
        private GroupBox mappingGroupBox;
        private Panel controlPanel;
        private GroupBox updateGroupBox;
        private TableLayoutPanel updateTableLayout;
        private RadioButton radioExcelToCAD;
        private RadioButton radioCADToExcel;
        private Button btnUpdate;
        private GroupBox logGroupBox;
        private RichTextBox txtLog;

        // 新添加的控件
        private Panel mappingButtonPanel;
        private CheckBox checkSelectAll;
        private Button btnAdd;
        private Button btnDelete;
        private Button btnImport;
        private Button btnExport;
        private CheckBox checkBackup;
        private DataGridView gridMappings;
        private DataGridViewCheckBoxColumn EnabledColumn;
        private DataGridViewTextBoxColumn DwgKeywordColumn;
        private DataGridViewTextBoxColumn TableTitleColumn;
        private DataGridViewTextBoxColumn ExcelSheetColumn;
        private CheckBox checkMergeCells;
    }
}