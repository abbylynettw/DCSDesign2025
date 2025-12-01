using System;
using System.Drawing;
using System.Windows.Forms;

namespace WinformUI.UpdateFrameTag
{
    partial class ReadExcelModifyCAD
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

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.mainTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.topTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.fileGroupBox = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.dwgFilesListView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.fileToolStrip = new System.Windows.Forms.ToolStrip();
            this.fileCountLabel = new System.Windows.Forms.ToolStripLabel();
            this.logGroupBox = new System.Windows.Forms.GroupBox();
            this.logRichTextBox = new System.Windows.Forms.RichTextBox();
            this.excelGroupBox = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tlpTitleBlock = new System.Windows.Forms.TableLayoutPanel();
            this.lblTitle1 = new System.Windows.Forms.Label();
            this.TITLE_1 = new System.Windows.Forms.TextBox();
            this.lblProjectNumber = new System.Windows.Forms.Label();
            this.PROJECT_NUMBER = new System.Windows.Forms.TextBox();
            this.lblTitle2 = new System.Windows.Forms.Label();
            this.TITLE_2 = new System.Windows.Forms.TextBox();
            this.lblDrawingNumber = new System.Windows.Forms.Label();
            this.DRAWING_NUMBER = new System.Windows.Forms.TextBox();
            this.lblTitle3 = new System.Windows.Forms.Label();
            this.TITLE_3 = new System.Windows.Forms.TextBox();
            this.lblPageNumber = new System.Windows.Forms.Label();
            this.NO1 = new System.Windows.Forms.TextBox();
            this.lblPageSeparator = new System.Windows.Forms.Label();
            this.NO2 = new System.Windows.Forms.TextBox();
            this.lblProjectName1 = new System.Windows.Forms.Label();
            this.PROJECT_NAME1 = new System.Windows.Forms.TextBox();
            this.lblRevision = new System.Windows.Forms.Label();
            this.REV = new System.Windows.Forms.TextBox();
            this.lblRevisionSeparator = new System.Windows.Forms.Label();
            this.REV2 = new System.Windows.Forms.TextBox();
            this.lblProjectName2 = new System.Windows.Forms.Label();
            this.PROJECT_NAME2 = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.REVIEWER = new System.Windows.Forms.TextBox();
            this.DRAFTER = new System.Windows.Forms.TextBox();
            this.APPROVAL = new System.Windows.Forms.TextBox();
            this.CHECKER = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tlpRevisions = new System.Windows.Forms.TableLayoutPanel();
            this.REASON_A = new System.Windows.Forms.TextBox();
            this.STAT_A = new System.Windows.Forms.TextBox();
            this.DATE_A = new System.Windows.Forms.TextBox();
            this.STAT_D = new System.Windows.Forms.TextBox();
            this.REV_A = new System.Windows.Forms.TextBox();
            this.REV_D = new System.Windows.Forms.TextBox();
            this.REASON_B = new System.Windows.Forms.TextBox();
            this.STAT_B = new System.Windows.Forms.TextBox();
            this.DATE_B = new System.Windows.Forms.TextBox();
            this.STAT_F = new System.Windows.Forms.TextBox();
            this.STAT_E = new System.Windows.Forms.TextBox();
            this.REV_B = new System.Windows.Forms.TextBox();
            this.REV_E = new System.Windows.Forms.TextBox();
            this.REASON_C = new System.Windows.Forms.TextBox();
            this.STAT_C = new System.Windows.Forms.TextBox();
            this.DATE_C = new System.Windows.Forms.TextBox();
            this.REV_C = new System.Windows.Forms.TextBox();
            this.REV_F = new System.Windows.Forms.TextBox();
            this.DATE_F = new System.Windows.Forms.TextBox();
            this.REASON_F = new System.Windows.Forms.TextBox();
            this.DATE_E = new System.Windows.Forms.TextBox();
            this.REASON_E = new System.Windows.Forms.TextBox();
            this.DATE_D = new System.Windows.Forms.TextBox();
            this.REASON_D = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.excelPathPanel = new System.Windows.Forms.Panel();
            this.browseExcelButton = new System.Windows.Forms.Button();
            this.downloadTemplateButton = new System.Windows.Forms.Button();
            this.excelPathTextBox = new System.Windows.Forms.TextBox();
            this.controlPanel = new System.Windows.Forms.Panel();
            this.matchingGroupBox = new System.Windows.Forms.GroupBox();
            this.containsMatchRadioButton = new System.Windows.Forms.RadioButton();
            this.startButton = new System.Windows.Forms.Button();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.openFileDialogDwg = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialogExcel = new System.Windows.Forms.OpenFileDialog();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.mainTableLayoutPanel.SuspendLayout();
            this.topTableLayoutPanel.SuspendLayout();
            this.fileGroupBox.SuspendLayout();
            this.fileToolStrip.SuspendLayout();
            this.logGroupBox.SuspendLayout();
            this.excelGroupBox.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tlpTitleBlock.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tlpRevisions.SuspendLayout();
            this.excelPathPanel.SuspendLayout();
            this.controlPanel.SuspendLayout();
            this.matchingGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainTableLayoutPanel
            // 
            this.mainTableLayoutPanel.ColumnCount = 1;
            this.mainTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainTableLayoutPanel.Controls.Add(this.topTableLayoutPanel, 0, 0);
            this.mainTableLayoutPanel.Controls.Add(this.excelGroupBox, 0, 1);
            this.mainTableLayoutPanel.Controls.Add(this.controlPanel, 0, 2);
            this.mainTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainTableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.mainTableLayoutPanel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.mainTableLayoutPanel.Name = "mainTableLayoutPanel";
            this.mainTableLayoutPanel.RowCount = 3;
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 46.89564F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 42.93263F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.mainTableLayoutPanel.Size = new System.Drawing.Size(1351, 821);
            this.mainTableLayoutPanel.TabIndex = 0;
            // 
            // topTableLayoutPanel
            // 
            this.topTableLayoutPanel.ColumnCount = 2;
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.topTableLayoutPanel.Controls.Add(this.fileGroupBox, 0, 0);
            this.topTableLayoutPanel.Controls.Add(this.logGroupBox, 1, 0);
            this.topTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.topTableLayoutPanel.Location = new System.Drawing.Point(4, 3);
            this.topTableLayoutPanel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.topTableLayoutPanel.Name = "topTableLayoutPanel";
            this.topTableLayoutPanel.RowCount = 1;
            this.topTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.topTableLayoutPanel.Size = new System.Drawing.Size(1343, 379);
            this.topTableLayoutPanel.TabIndex = 0;
            // 
            // fileGroupBox
            // 
            this.fileGroupBox.Controls.Add(this.button1);
            this.fileGroupBox.Controls.Add(this.dwgFilesListView);
            this.fileGroupBox.Controls.Add(this.fileToolStrip);
            this.fileGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fileGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.fileGroupBox.Location = new System.Drawing.Point(4, 3);
            this.fileGroupBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.fileGroupBox.Name = "fileGroupBox";
            this.fileGroupBox.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.fileGroupBox.Size = new System.Drawing.Size(663, 373);
            this.fileGroupBox.TabIndex = 0;
            this.fileGroupBox.TabStop = false;
            this.fileGroupBox.Text = "DWG文件 (拖拽)";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.HotTrack;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(20, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(118, 29);
            this.button1.TabIndex = 2;
            this.button1.Text = "选择文件夹";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dwgFilesListView
            // 
            this.dwgFilesListView.AllowDrop = true;
            this.dwgFilesListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
            this.dwgFilesListView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dwgFilesListView.FullRowSelect = true;
            this.dwgFilesListView.GridLines = true;
            this.dwgFilesListView.HideSelection = false;
            this.dwgFilesListView.Location = new System.Drawing.Point(4, 53);
            this.dwgFilesListView.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.dwgFilesListView.Name = "dwgFilesListView";
            this.dwgFilesListView.Size = new System.Drawing.Size(655, 317);
            this.dwgFilesListView.TabIndex = 0;
            this.dwgFilesListView.UseCompatibleStateImageBehavior = false;
            this.dwgFilesListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "文件名";
            this.columnHeader1.Width = 250;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "状态";
            this.columnHeader2.Width = 250;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "文件路径";
            this.columnHeader3.Width = 100;
            // 
            // fileToolStrip
            // 
            this.fileToolStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.fileToolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileCountLabel});
            this.fileToolStrip.Location = new System.Drawing.Point(4, 24);
            this.fileToolStrip.Name = "fileToolStrip";
            this.fileToolStrip.Padding = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.fileToolStrip.Size = new System.Drawing.Size(655, 29);
            this.fileToolStrip.TabIndex = 1;
            // 
            // fileCountLabel
            // 
            this.fileCountLabel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.fileCountLabel.Name = "fileCountLabel";
            this.fileCountLabel.Size = new System.Drawing.Size(84, 24);
            this.fileCountLabel.Text = "文件数: 0";
            // 
            // logGroupBox
            // 
            this.logGroupBox.Controls.Add(this.logRichTextBox);
            this.logGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logGroupBox.Location = new System.Drawing.Point(675, 3);
            this.logGroupBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.logGroupBox.Name = "logGroupBox";
            this.logGroupBox.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.logGroupBox.Size = new System.Drawing.Size(664, 373);
            this.logGroupBox.TabIndex = 1;
            this.logGroupBox.TabStop = false;
            this.logGroupBox.Text = "提示信息";
            // 
            // logRichTextBox
            // 
            this.logRichTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logRichTextBox.Location = new System.Drawing.Point(4, 24);
            this.logRichTextBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.logRichTextBox.Name = "logRichTextBox";
            this.logRichTextBox.Size = new System.Drawing.Size(656, 346);
            this.logRichTextBox.TabIndex = 0;
            this.logRichTextBox.Text = "";
            // 
            // excelGroupBox
            // 
            this.excelGroupBox.Controls.Add(this.tabControl1);
            this.excelGroupBox.Controls.Add(this.excelPathPanel);
            this.excelGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.excelGroupBox.Location = new System.Drawing.Point(4, 388);
            this.excelGroupBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.excelGroupBox.Name = "excelGroupBox";
            this.excelGroupBox.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.excelGroupBox.Size = new System.Drawing.Size(1343, 347);
            this.excelGroupBox.TabIndex = 1;
            this.excelGroupBox.TabStop = false;
            this.excelGroupBox.Text = "CAD图框信息";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(4, 80);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1270, 233);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.AutoScroll = true;
            this.tabPage1.Controls.Add(this.tlpTitleBlock);
            this.tabPage1.Location = new System.Drawing.Point(4, 28);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(10);
            this.tabPage1.Size = new System.Drawing.Size(1262, 201);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "全局图框信息1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tlpTitleBlock
            // 
            this.tlpTitleBlock.ColumnCount = 6;
            this.tlpTitleBlock.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.77814F));
            this.tlpTitleBlock.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.77814F));
            this.tlpTitleBlock.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.10896F));
            this.tlpTitleBlock.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.77814F));
            this.tlpTitleBlock.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 12.57679F));
            this.tlpTitleBlock.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.97983F));
            this.tlpTitleBlock.Controls.Add(this.lblTitle1, 0, 0);
            this.tlpTitleBlock.Controls.Add(this.TITLE_1, 1, 0);
            this.tlpTitleBlock.Controls.Add(this.lblProjectNumber, 2, 0);
            this.tlpTitleBlock.Controls.Add(this.PROJECT_NUMBER, 3, 0);
            this.tlpTitleBlock.Controls.Add(this.lblTitle2, 0, 1);
            this.tlpTitleBlock.Controls.Add(this.TITLE_2, 1, 1);
            this.tlpTitleBlock.Controls.Add(this.lblDrawingNumber, 2, 1);
            this.tlpTitleBlock.Controls.Add(this.DRAWING_NUMBER, 3, 1);
            this.tlpTitleBlock.Controls.Add(this.lblTitle3, 0, 2);
            this.tlpTitleBlock.Controls.Add(this.TITLE_3, 1, 2);
            this.tlpTitleBlock.Controls.Add(this.lblPageNumber, 2, 2);
            this.tlpTitleBlock.Controls.Add(this.NO1, 3, 2);
            this.tlpTitleBlock.Controls.Add(this.lblPageSeparator, 4, 2);
            this.tlpTitleBlock.Controls.Add(this.NO2, 5, 2);
            this.tlpTitleBlock.Controls.Add(this.lblProjectName1, 0, 3);
            this.tlpTitleBlock.Controls.Add(this.PROJECT_NAME1, 1, 3);
            this.tlpTitleBlock.Controls.Add(this.lblRevision, 2, 3);
            this.tlpTitleBlock.Controls.Add(this.REV, 3, 3);
            this.tlpTitleBlock.Controls.Add(this.lblRevisionSeparator, 4, 3);
            this.tlpTitleBlock.Controls.Add(this.REV2, 5, 3);
            this.tlpTitleBlock.Controls.Add(this.lblProjectName2, 0, 4);
            this.tlpTitleBlock.Controls.Add(this.PROJECT_NAME2, 1, 4);
            this.tlpTitleBlock.Controls.Add(this.label9, 2, 4);
            this.tlpTitleBlock.Controls.Add(this.label11, 0, 5);
            this.tlpTitleBlock.Controls.Add(this.REVIEWER, 1, 5);
            this.tlpTitleBlock.Controls.Add(this.DRAFTER, 3, 4);
            this.tlpTitleBlock.Controls.Add(this.APPROVAL, 3, 5);
            this.tlpTitleBlock.Controls.Add(this.CHECKER, 5, 4);
            this.tlpTitleBlock.Controls.Add(this.label10, 4, 4);
            this.tlpTitleBlock.Controls.Add(this.label12, 4, 0);
            this.tlpTitleBlock.Controls.Add(this.label13, 2, 5);
            this.tlpTitleBlock.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpTitleBlock.Location = new System.Drawing.Point(10, 10);
            this.tlpTitleBlock.Name = "tlpTitleBlock";
            this.tlpTitleBlock.RowCount = 6;
            this.tlpTitleBlock.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tlpTitleBlock.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tlpTitleBlock.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tlpTitleBlock.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tlpTitleBlock.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tlpTitleBlock.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.tlpTitleBlock.Size = new System.Drawing.Size(1242, 181);
            this.tlpTitleBlock.TabIndex = 0;
            // 
            // lblTitle1
            // 
            this.lblTitle1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTitle1.Location = new System.Drawing.Point(3, 3);
            this.lblTitle1.Name = "lblTitle1";
            this.lblTitle1.Size = new System.Drawing.Size(252, 23);
            this.lblTitle1.TabIndex = 0;
            this.lblTitle1.Text = "图纸名称_TITLE1:";
            this.lblTitle1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // TITLE_1
            // 
            this.TITLE_1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.TITLE_1.Location = new System.Drawing.Point(261, 3);
            this.TITLE_1.Name = "TITLE_1";
            this.TITLE_1.Size = new System.Drawing.Size(252, 28);
            this.TITLE_1.TabIndex = 1;
            // 
            // lblProjectNumber
            // 
            this.lblProjectNumber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblProjectNumber.Location = new System.Drawing.Point(519, 3);
            this.lblProjectNumber.Name = "lblProjectNumber";
            this.lblProjectNumber.Size = new System.Drawing.Size(169, 23);
            this.lblProjectNumber.TabIndex = 2;
            this.lblProjectNumber.Text = "PROJECT_NUMBER:";
            this.lblProjectNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // PROJECT_NUMBER
            // 
            this.PROJECT_NUMBER.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.PROJECT_NUMBER.Location = new System.Drawing.Point(694, 3);
            this.PROJECT_NUMBER.Name = "PROJECT_NUMBER";
            this.PROJECT_NUMBER.Size = new System.Drawing.Size(252, 28);
            this.PROJECT_NUMBER.TabIndex = 3;
            // 
            // lblTitle2
            // 
            this.lblTitle2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTitle2.Location = new System.Drawing.Point(3, 33);
            this.lblTitle2.Name = "lblTitle2";
            this.lblTitle2.Size = new System.Drawing.Size(252, 23);
            this.lblTitle2.TabIndex = 4;
            this.lblTitle2.Text = "图纸名称_TITLE2:";
            this.lblTitle2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // TITLE_2
            // 
            this.TITLE_2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.TITLE_2.Location = new System.Drawing.Point(261, 33);
            this.TITLE_2.Name = "TITLE_2";
            this.TITLE_2.Size = new System.Drawing.Size(252, 28);
            this.TITLE_2.TabIndex = 5;
            // 
            // lblDrawingNumber
            // 
            this.lblDrawingNumber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblDrawingNumber.Location = new System.Drawing.Point(519, 33);
            this.lblDrawingNumber.Name = "lblDrawingNumber";
            this.lblDrawingNumber.Size = new System.Drawing.Size(169, 23);
            this.lblDrawingNumber.TabIndex = 6;
            this.lblDrawingNumber.Text = "DRAWING_NUMBER:";
            this.lblDrawingNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // DRAWING_NUMBER
            // 
            this.DRAWING_NUMBER.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DRAWING_NUMBER.Location = new System.Drawing.Point(694, 33);
            this.DRAWING_NUMBER.Name = "DRAWING_NUMBER";
            this.DRAWING_NUMBER.Size = new System.Drawing.Size(252, 28);
            this.DRAWING_NUMBER.TabIndex = 7;
            // 
            // lblTitle3
            // 
            this.lblTitle3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTitle3.Location = new System.Drawing.Point(3, 63);
            this.lblTitle3.Name = "lblTitle3";
            this.lblTitle3.Size = new System.Drawing.Size(252, 23);
            this.lblTitle3.TabIndex = 8;
            this.lblTitle3.Text = "图纸名称_TITLE3:";
            this.lblTitle3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // TITLE_3
            // 
            this.TITLE_3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.TITLE_3.Location = new System.Drawing.Point(261, 63);
            this.TITLE_3.Name = "TITLE_3";
            this.TITLE_3.Size = new System.Drawing.Size(252, 28);
            this.TITLE_3.TabIndex = 9;
            // 
            // lblPageNumber
            // 
            this.lblPageNumber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPageNumber.Location = new System.Drawing.Point(519, 63);
            this.lblPageNumber.Name = "lblPageNumber";
            this.lblPageNumber.Size = new System.Drawing.Size(169, 23);
            this.lblPageNumber.TabIndex = 10;
            this.lblPageNumber.Text = "NO1/NO2:";
            this.lblPageNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // NO1
            // 
            this.NO1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.NO1.Location = new System.Drawing.Point(694, 63);
            this.NO1.Name = "NO1";
            this.NO1.Size = new System.Drawing.Size(252, 28);
            this.NO1.TabIndex = 11;
            // 
            // lblPageSeparator
            // 
            this.lblPageSeparator.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.lblPageSeparator.AutoSize = true;
            this.lblPageSeparator.Location = new System.Drawing.Point(1018, 66);
            this.lblPageSeparator.Name = "lblPageSeparator";
            this.lblPageSeparator.Size = new System.Drawing.Size(17, 18);
            this.lblPageSeparator.TabIndex = 12;
            this.lblPageSeparator.Text = "/";
            // 
            // NO2
            // 
            this.NO2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.NO2.Location = new System.Drawing.Point(1108, 63);
            this.NO2.Name = "NO2";
            this.NO2.Size = new System.Drawing.Size(131, 28);
            this.NO2.TabIndex = 13;
            // 
            // lblProjectName1
            // 
            this.lblProjectName1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblProjectName1.Location = new System.Drawing.Point(3, 93);
            this.lblProjectName1.Name = "lblProjectName1";
            this.lblProjectName1.Size = new System.Drawing.Size(252, 23);
            this.lblProjectName1.TabIndex = 14;
            this.lblProjectName1.Text = "PROJECT_NAME1";
            this.lblProjectName1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // PROJECT_NAME1
            // 
            this.PROJECT_NAME1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.PROJECT_NAME1.Location = new System.Drawing.Point(261, 93);
            this.PROJECT_NAME1.Name = "PROJECT_NAME1";
            this.PROJECT_NAME1.Size = new System.Drawing.Size(252, 28);
            this.PROJECT_NAME1.TabIndex = 15;
            // 
            // lblRevision
            // 
            this.lblRevision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblRevision.Location = new System.Drawing.Point(519, 93);
            this.lblRevision.Name = "lblRevision";
            this.lblRevision.Size = new System.Drawing.Size(169, 23);
            this.lblRevision.TabIndex = 16;
            this.lblRevision.Text = "REV/REV2:";
            this.lblRevision.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // REV
            // 
            this.REV.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV.Location = new System.Drawing.Point(694, 93);
            this.REV.Name = "REV";
            this.REV.Size = new System.Drawing.Size(252, 28);
            this.REV.TabIndex = 17;
            // 
            // lblRevisionSeparator
            // 
            this.lblRevisionSeparator.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.lblRevisionSeparator.AutoSize = true;
            this.lblRevisionSeparator.Location = new System.Drawing.Point(1018, 96);
            this.lblRevisionSeparator.Name = "lblRevisionSeparator";
            this.lblRevisionSeparator.Size = new System.Drawing.Size(17, 18);
            this.lblRevisionSeparator.TabIndex = 18;
            this.lblRevisionSeparator.Text = "/";
            // 
            // REV2
            // 
            this.REV2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV2.Location = new System.Drawing.Point(1108, 93);
            this.REV2.Name = "REV2";
            this.REV2.Size = new System.Drawing.Size(131, 28);
            this.REV2.TabIndex = 19;
            // 
            // lblProjectName2
            // 
            this.lblProjectName2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblProjectName2.Location = new System.Drawing.Point(3, 123);
            this.lblProjectName2.Name = "lblProjectName2";
            this.lblProjectName2.Size = new System.Drawing.Size(252, 23);
            this.lblProjectName2.TabIndex = 20;
            this.lblProjectName2.Text = "PROJECT_NAME2";
            this.lblProjectName2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // PROJECT_NAME2
            // 
            this.PROJECT_NAME2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.PROJECT_NAME2.Location = new System.Drawing.Point(261, 123);
            this.PROJECT_NAME2.Name = "PROJECT_NAME2";
            this.PROJECT_NAME2.Size = new System.Drawing.Size(252, 28);
            this.PROJECT_NAME2.TabIndex = 21;
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.Location = new System.Drawing.Point(519, 123);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(169, 23);
            this.label9.TabIndex = 22;
            this.label9.Text = "编制_DRAFTER";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.Location = new System.Drawing.Point(3, 154);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(252, 23);
            this.label11.TabIndex = 24;
            this.label11.Text = "审核_REVIEWER";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // REVIEWER
            // 
            this.REVIEWER.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REVIEWER.Location = new System.Drawing.Point(261, 153);
            this.REVIEWER.Name = "REVIEWER";
            this.REVIEWER.Size = new System.Drawing.Size(252, 28);
            this.REVIEWER.TabIndex = 26;
            // 
            // DRAFTER
            // 
            this.DRAFTER.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DRAFTER.Location = new System.Drawing.Point(694, 123);
            this.DRAFTER.Name = "DRAFTER";
            this.DRAFTER.Size = new System.Drawing.Size(252, 28);
            this.DRAFTER.TabIndex = 27;
            // 
            // APPROVAL
            // 
            this.APPROVAL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.APPROVAL.Location = new System.Drawing.Point(694, 153);
            this.APPROVAL.Name = "APPROVAL";
            this.APPROVAL.Size = new System.Drawing.Size(252, 28);
            this.APPROVAL.TabIndex = 28;
            // 
            // CHECKER
            // 
            this.CHECKER.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.CHECKER.Location = new System.Drawing.Point(1108, 123);
            this.CHECKER.Name = "CHECKER";
            this.CHECKER.Size = new System.Drawing.Size(131, 28);
            this.CHECKER.TabIndex = 29;
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.Location = new System.Drawing.Point(952, 123);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(150, 23);
            this.label10.TabIndex = 30;
            this.label10.Text = "校对_CHECKER";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label12.Location = new System.Drawing.Point(952, 3);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(150, 23);
            this.label12.TabIndex = 25;
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label13.Location = new System.Drawing.Point(519, 154);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(169, 23);
            this.label13.TabIndex = 31;
            this.label13.Text = "批准_APPROVAL";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tlpRevisions);
            this.tabPage2.Location = new System.Drawing.Point(4, 28);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1262, 201);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "全局图框信息2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tlpRevisions
            // 
            this.tlpRevisions.ColumnCount = 8;
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.49425F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.49425F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.49425F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.51724F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.49425F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.49425F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.49425F));
            this.tlpRevisions.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.51724F));
            this.tlpRevisions.Controls.Add(this.REASON_A, 7, 2);
            this.tlpRevisions.Controls.Add(this.STAT_A, 6, 2);
            this.tlpRevisions.Controls.Add(this.DATE_A, 5, 2);
            this.tlpRevisions.Controls.Add(this.STAT_D, 2, 2);
            this.tlpRevisions.Controls.Add(this.REV_A, 4, 2);
            this.tlpRevisions.Controls.Add(this.REV_D, 0, 2);
            this.tlpRevisions.Controls.Add(this.REASON_B, 7, 1);
            this.tlpRevisions.Controls.Add(this.STAT_B, 6, 1);
            this.tlpRevisions.Controls.Add(this.DATE_B, 5, 1);
            this.tlpRevisions.Controls.Add(this.STAT_F, 2, 0);
            this.tlpRevisions.Controls.Add(this.STAT_E, 2, 1);
            this.tlpRevisions.Controls.Add(this.REV_B, 4, 1);
            this.tlpRevisions.Controls.Add(this.REV_E, 0, 1);
            this.tlpRevisions.Controls.Add(this.REASON_C, 7, 0);
            this.tlpRevisions.Controls.Add(this.STAT_C, 6, 0);
            this.tlpRevisions.Controls.Add(this.DATE_C, 5, 0);
            this.tlpRevisions.Controls.Add(this.REV_C, 4, 0);
            this.tlpRevisions.Controls.Add(this.REV_F, 0, 0);
            this.tlpRevisions.Controls.Add(this.DATE_F, 1, 0);
            this.tlpRevisions.Controls.Add(this.REASON_F, 3, 0);
            this.tlpRevisions.Controls.Add(this.DATE_E, 1, 1);
            this.tlpRevisions.Controls.Add(this.REASON_E, 3, 1);
            this.tlpRevisions.Controls.Add(this.DATE_D, 1, 2);
            this.tlpRevisions.Controls.Add(this.REASON_D, 3, 2);
            this.tlpRevisions.Controls.Add(this.label1, 7, 3);
            this.tlpRevisions.Controls.Add(this.label3, 6, 3);
            this.tlpRevisions.Controls.Add(this.label4, 5, 3);
            this.tlpRevisions.Controls.Add(this.label5, 4, 3);
            this.tlpRevisions.Controls.Add(this.label2, 0, 3);
            this.tlpRevisions.Controls.Add(this.label6, 1, 3);
            this.tlpRevisions.Controls.Add(this.label7, 2, 3);
            this.tlpRevisions.Controls.Add(this.label8, 3, 3);
            this.tlpRevisions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpRevisions.Location = new System.Drawing.Point(3, 3);
            this.tlpRevisions.Name = "tlpRevisions";
            this.tlpRevisions.Padding = new System.Windows.Forms.Padding(10);
            this.tlpRevisions.RowCount = 4;
            this.tlpRevisions.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tlpRevisions.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tlpRevisions.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tlpRevisions.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tlpRevisions.Size = new System.Drawing.Size(1256, 195);
            this.tlpRevisions.TabIndex = 0;
            // 
            // REASON_A
            // 
            this.REASON_A.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REASON_A.Location = new System.Drawing.Point(1056, 103);
            this.REASON_A.Name = "REASON_A";
            this.REASON_A.Size = new System.Drawing.Size(187, 28);
            this.REASON_A.TabIndex = 42;
            // 
            // STAT_A
            // 
            this.STAT_A.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.STAT_A.Location = new System.Drawing.Point(914, 103);
            this.STAT_A.Name = "STAT_A";
            this.STAT_A.Size = new System.Drawing.Size(136, 28);
            this.STAT_A.TabIndex = 41;
            // 
            // DATE_A
            // 
            this.DATE_A.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DATE_A.Location = new System.Drawing.Point(772, 103);
            this.DATE_A.Name = "DATE_A";
            this.DATE_A.Size = new System.Drawing.Size(136, 28);
            this.DATE_A.TabIndex = 40;
            // 
            // STAT_D
            // 
            this.STAT_D.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.STAT_D.Location = new System.Drawing.Point(297, 103);
            this.STAT_D.Name = "STAT_D";
            this.STAT_D.Size = new System.Drawing.Size(136, 28);
            this.STAT_D.TabIndex = 39;
            // 
            // REV_A
            // 
            this.REV_A.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV_A.Location = new System.Drawing.Point(630, 103);
            this.REV_A.Name = "REV_A";
            this.REV_A.Size = new System.Drawing.Size(136, 28);
            this.REV_A.TabIndex = 38;
            // 
            // REV_D
            // 
            this.REV_D.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV_D.Location = new System.Drawing.Point(13, 103);
            this.REV_D.Name = "REV_D";
            this.REV_D.Size = new System.Drawing.Size(136, 28);
            this.REV_D.TabIndex = 37;
            // 
            // REASON_B
            // 
            this.REASON_B.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REASON_B.Location = new System.Drawing.Point(1056, 60);
            this.REASON_B.Name = "REASON_B";
            this.REASON_B.Size = new System.Drawing.Size(187, 28);
            this.REASON_B.TabIndex = 36;
            // 
            // STAT_B
            // 
            this.STAT_B.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.STAT_B.Location = new System.Drawing.Point(914, 60);
            this.STAT_B.Name = "STAT_B";
            this.STAT_B.Size = new System.Drawing.Size(136, 28);
            this.STAT_B.TabIndex = 35;
            // 
            // DATE_B
            // 
            this.DATE_B.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DATE_B.Location = new System.Drawing.Point(772, 60);
            this.DATE_B.Name = "DATE_B";
            this.DATE_B.Size = new System.Drawing.Size(136, 28);
            this.DATE_B.TabIndex = 34;
            // 
            // STAT_F
            // 
            this.STAT_F.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.STAT_F.Location = new System.Drawing.Point(297, 17);
            this.STAT_F.Name = "STAT_F";
            this.STAT_F.Size = new System.Drawing.Size(136, 28);
            this.STAT_F.TabIndex = 33;
            // 
            // STAT_E
            // 
            this.STAT_E.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.STAT_E.Location = new System.Drawing.Point(297, 60);
            this.STAT_E.Name = "STAT_E";
            this.STAT_E.Size = new System.Drawing.Size(136, 28);
            this.STAT_E.TabIndex = 32;
            // 
            // REV_B
            // 
            this.REV_B.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV_B.Location = new System.Drawing.Point(630, 60);
            this.REV_B.Name = "REV_B";
            this.REV_B.Size = new System.Drawing.Size(136, 28);
            this.REV_B.TabIndex = 31;
            // 
            // REV_E
            // 
            this.REV_E.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV_E.Location = new System.Drawing.Point(13, 60);
            this.REV_E.Name = "REV_E";
            this.REV_E.Size = new System.Drawing.Size(136, 28);
            this.REV_E.TabIndex = 30;
            // 
            // REASON_C
            // 
            this.REASON_C.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REASON_C.Location = new System.Drawing.Point(1056, 17);
            this.REASON_C.Name = "REASON_C";
            this.REASON_C.Size = new System.Drawing.Size(187, 28);
            this.REASON_C.TabIndex = 29;
            // 
            // STAT_C
            // 
            this.STAT_C.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.STAT_C.Location = new System.Drawing.Point(914, 17);
            this.STAT_C.Name = "STAT_C";
            this.STAT_C.Size = new System.Drawing.Size(136, 28);
            this.STAT_C.TabIndex = 28;
            // 
            // DATE_C
            // 
            this.DATE_C.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DATE_C.Location = new System.Drawing.Point(772, 17);
            this.DATE_C.Name = "DATE_C";
            this.DATE_C.Size = new System.Drawing.Size(136, 28);
            this.DATE_C.TabIndex = 27;
            // 
            // REV_C
            // 
            this.REV_C.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV_C.Location = new System.Drawing.Point(630, 17);
            this.REV_C.Name = "REV_C";
            this.REV_C.Size = new System.Drawing.Size(136, 28);
            this.REV_C.TabIndex = 26;
            // 
            // REV_F
            // 
            this.REV_F.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REV_F.Location = new System.Drawing.Point(13, 17);
            this.REV_F.Name = "REV_F";
            this.REV_F.Size = new System.Drawing.Size(136, 28);
            this.REV_F.TabIndex = 25;
            // 
            // DATE_F
            // 
            this.DATE_F.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DATE_F.Location = new System.Drawing.Point(155, 17);
            this.DATE_F.Name = "DATE_F";
            this.DATE_F.Size = new System.Drawing.Size(136, 28);
            this.DATE_F.TabIndex = 1;
            // 
            // REASON_F
            // 
            this.REASON_F.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REASON_F.Location = new System.Drawing.Point(439, 17);
            this.REASON_F.Name = "REASON_F";
            this.REASON_F.Size = new System.Drawing.Size(185, 28);
            this.REASON_F.TabIndex = 3;
            // 
            // DATE_E
            // 
            this.DATE_E.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DATE_E.Location = new System.Drawing.Point(155, 60);
            this.DATE_E.Name = "DATE_E";
            this.DATE_E.Size = new System.Drawing.Size(136, 28);
            this.DATE_E.TabIndex = 5;
            // 
            // REASON_E
            // 
            this.REASON_E.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REASON_E.Location = new System.Drawing.Point(439, 60);
            this.REASON_E.Name = "REASON_E";
            this.REASON_E.Size = new System.Drawing.Size(185, 28);
            this.REASON_E.TabIndex = 7;
            // 
            // DATE_D
            // 
            this.DATE_D.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.DATE_D.Location = new System.Drawing.Point(155, 103);
            this.DATE_D.Name = "DATE_D";
            this.DATE_D.Size = new System.Drawing.Size(136, 28);
            this.DATE_D.TabIndex = 9;
            // 
            // REASON_D
            // 
            this.REASON_D.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.REASON_D.Location = new System.Drawing.Point(439, 103);
            this.REASON_D.Name = "REASON_D";
            this.REASON_D.Size = new System.Drawing.Size(185, 28);
            this.REASON_D.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Location = new System.Drawing.Point(1056, 150);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(187, 23);
            this.label1.TabIndex = 16;
            this.label1.Text = "发布说明";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.Location = new System.Drawing.Point(914, 150);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(136, 23);
            this.label3.TabIndex = 18;
            this.label3.Text = "状态";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.Location = new System.Drawing.Point(772, 150);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(136, 23);
            this.label4.TabIndex = 19;
            this.label4.Text = "日期";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.Location = new System.Drawing.Point(630, 150);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(136, 23);
            this.label5.TabIndex = 20;
            this.label5.Text = "版本 ";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.Location = new System.Drawing.Point(13, 150);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 23);
            this.label2.TabIndex = 21;
            this.label2.Text = "版本 ";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.Location = new System.Drawing.Point(155, 150);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(136, 23);
            this.label6.TabIndex = 22;
            this.label6.Text = "日期";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.Location = new System.Drawing.Point(297, 150);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(136, 23);
            this.label7.TabIndex = 23;
            this.label7.Text = "状态";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.Location = new System.Drawing.Point(439, 150);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(185, 23);
            this.label8.TabIndex = 24;
            this.label8.Text = "发布说明";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // excelPathPanel
            // 
            this.excelPathPanel.Controls.Add(this.browseExcelButton);
            this.excelPathPanel.Controls.Add(this.downloadTemplateButton);
            this.excelPathPanel.Controls.Add(this.excelPathTextBox);
            this.excelPathPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.excelPathPanel.Location = new System.Drawing.Point(4, 24);
            this.excelPathPanel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.excelPathPanel.Name = "excelPathPanel";
            this.excelPathPanel.Size = new System.Drawing.Size(1335, 50);
            this.excelPathPanel.TabIndex = 0;
            // 
            // browseExcelButton
            // 
            this.browseExcelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.browseExcelButton.Location = new System.Drawing.Point(945, 4);
            this.browseExcelButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.browseExcelButton.Name = "browseExcelButton";
            this.browseExcelButton.Size = new System.Drawing.Size(193, 38);
            this.browseExcelButton.TabIndex = 1;
            this.browseExcelButton.Text = "浏览Excel...";
            this.browseExcelButton.UseVisualStyleBackColor = true;
            // 
            // downloadTemplateButton
            // 
            this.downloadTemplateButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.downloadTemplateButton.Location = new System.Drawing.Point(1145, 4);
            this.downloadTemplateButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.downloadTemplateButton.Name = "downloadTemplateButton";
            this.downloadTemplateButton.Size = new System.Drawing.Size(186, 38);
            this.downloadTemplateButton.TabIndex = 2;
            this.downloadTemplateButton.Text = "下载模板";
            this.downloadTemplateButton.UseVisualStyleBackColor = true;
            // 
            // excelPathTextBox
            // 
            this.excelPathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.excelPathTextBox.Location = new System.Drawing.Point(9, 11);
            this.excelPathTextBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.excelPathTextBox.Name = "excelPathTextBox";
            this.excelPathTextBox.Size = new System.Drawing.Size(928, 28);
            this.excelPathTextBox.TabIndex = 0;
            // 
            // controlPanel
            // 
            this.controlPanel.Controls.Add(this.matchingGroupBox);
            this.controlPanel.Controls.Add(this.startButton);
            this.controlPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.controlPanel.Location = new System.Drawing.Point(4, 741);
            this.controlPanel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.controlPanel.Name = "controlPanel";
            this.controlPanel.Size = new System.Drawing.Size(1343, 77);
            this.controlPanel.TabIndex = 2;
            // 
            // matchingGroupBox
            // 
            this.matchingGroupBox.Controls.Add(this.containsMatchRadioButton);
            this.matchingGroupBox.Location = new System.Drawing.Point(13, 3);
            this.matchingGroupBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.matchingGroupBox.Name = "matchingGroupBox";
            this.matchingGroupBox.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.matchingGroupBox.Size = new System.Drawing.Size(668, 65);
            this.matchingGroupBox.TabIndex = 0;
            this.matchingGroupBox.TabStop = false;
            this.matchingGroupBox.Text = "文件名匹配方式";
            // 
            // containsMatchRadioButton
            // 
            this.containsMatchRadioButton.AutoSize = true;
            this.containsMatchRadioButton.Checked = true;
            this.containsMatchRadioButton.Enabled = false;
            this.containsMatchRadioButton.Location = new System.Drawing.Point(11, 27);
            this.containsMatchRadioButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.containsMatchRadioButton.Name = "containsMatchRadioButton";
            this.containsMatchRadioButton.Size = new System.Drawing.Size(366, 22);
            this.containsMatchRadioButton.TabIndex = 1;
            this.containsMatchRadioButton.TabStop = true;
            this.containsMatchRadioButton.Text = "包含匹配(文件名包含Excel中的图纸名称)";
            this.containsMatchRadioButton.UseVisualStyleBackColor = true;
            // 
            // startButton
            // 
            this.startButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.startButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.startButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.startButton.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold);
            this.startButton.ForeColor = System.Drawing.Color.White;
            this.startButton.Location = new System.Drawing.Point(1145, 17);
            this.startButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(193, 39);
            this.startButton.TabIndex = 1;
            this.startButton.Text = "开始修改";
            this.startButton.UseVisualStyleBackColor = false;
            // 
            // statusStrip
            // 
            this.statusStrip.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.statusStrip.Location = new System.Drawing.Point(0, 821);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Padding = new System.Windows.Forms.Padding(1, 0, 18, 0);
            this.statusStrip.Size = new System.Drawing.Size(1351, 22);
            this.statusStrip.TabIndex = 1;
            // 
            // openFileDialogDwg
            // 
            this.openFileDialogDwg.Filter = "DWG文件 (*.dwg)|*.dwg|所有文件 (*.*)|*.*";
            this.openFileDialogDwg.Multiselect = true;
            this.openFileDialogDwg.Title = "选择DWG文件";
            // 
            // openFileDialogExcel
            // 
            this.openFileDialogExcel.Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
            this.openFileDialogExcel.Title = "选择Excel文件";
            // 
            // ReadExcelModifyCAD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1351, 843);
            this.Controls.Add(this.mainTableLayoutPanel);
            this.Controls.Add(this.statusStrip);
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MinimumSize = new System.Drawing.Size(1022, 668);
            this.Name = "ReadExcelModifyCAD";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CAD图框批量修改工具";
            this.mainTableLayoutPanel.ResumeLayout(false);
            this.topTableLayoutPanel.ResumeLayout(false);
            this.fileGroupBox.ResumeLayout(false);
            this.fileGroupBox.PerformLayout();
            this.fileToolStrip.ResumeLayout(false);
            this.fileToolStrip.PerformLayout();
            this.logGroupBox.ResumeLayout(false);
            this.excelGroupBox.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tlpTitleBlock.ResumeLayout(false);
            this.tlpTitleBlock.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tlpRevisions.ResumeLayout(false);
            this.tlpRevisions.PerformLayout();
            this.excelPathPanel.ResumeLayout(false);
            this.excelPathPanel.PerformLayout();
            this.controlPanel.ResumeLayout(false);
            this.matchingGroupBox.ResumeLayout(false);
            this.matchingGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        // 控件声明
        private TableLayoutPanel mainTableLayoutPanel;
        private TableLayoutPanel topTableLayoutPanel;

        // DWG文件区域
        private GroupBox fileGroupBox;
        private ListView dwgFilesListView;
        private ToolStrip fileToolStrip;
        private ToolStripLabel fileCountLabel;
        private ColumnHeader columnHeader1;
        private ColumnHeader columnHeader2;
        private ColumnHeader columnHeader3;

        // 日志区域
        private GroupBox logGroupBox;
        private RichTextBox logRichTextBox;

        // Excel配置区域
        private GroupBox excelGroupBox;
        private Panel excelPathPanel;
        private TextBox excelPathTextBox;
        private Button browseExcelButton;
        private Button downloadTemplateButton;

        // 控制区域
        private Panel controlPanel;
        private GroupBox matchingGroupBox;
        private RadioButton containsMatchRadioButton;
        private Button startButton;

        // 状态栏
        private StatusStrip statusStrip;

        // 对话框和工具提示
        private OpenFileDialog openFileDialogDwg;
        private OpenFileDialog openFileDialogExcel;
        private ToolTip toolTip;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Label lblTitle1;
        private Label lblTitle2;
        private Label lblTitle3;
        private TextBox TITLE_1;
        private TextBox TITLE_2;
        private TextBox TITLE_3;

        private Label lblProjectName1;
        private Label lblProjectName2;
        private TextBox PROJECT_NAME1;
        private TextBox PROJECT_NAME2;

        private Label lblProjectNumber;
        private Label lblDrawingNumber;
        private Label lblPageNumber;
        private Label lblRevision;
        private TextBox PROJECT_NUMBER;
        private TextBox DRAWING_NUMBER;
        private TextBox NO1;
        private TextBox NO2;
        private Label lblPageSeparator;
        private TextBox REV;
        private TextBox REV2;
        private Label lblRevisionSeparator;
        private TableLayoutPanel tlpTitleBlock;

        // Add these declarations to your class fields
        private TableLayoutPanel tlpRevisions;
        private TextBox DATE_F;
        private TextBox DATE_E;
        private TextBox REASON_F;
        private TextBox REASON_E;
        private TextBox DATE_D;
        private TextBox REASON_D;
        private Label label1;
        private TextBox REASON_A;
        private TextBox STAT_A;
        private TextBox DATE_A;
        private TextBox STAT_D;
        private TextBox REV_A;
        private TextBox REV_D;
        private TextBox REASON_B;
        private TextBox STAT_B;
        private TextBox DATE_B;
        private TextBox STAT_F;
        private TextBox STAT_E;
        private TextBox REV_B;
        private TextBox REV_E;
        private TextBox REASON_C;
        private TextBox STAT_C;
        private TextBox DATE_C;
        private TextBox REV_C;
        private TextBox REV_F;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label2;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private Label label11;
        private Label label12;
        private TextBox REVIEWER;
        private TextBox DRAFTER;
        private TextBox APPROVAL;
        private TextBox CHECKER;
        private Label label10;
        private Label label13;
        private Button button1;
    }
}