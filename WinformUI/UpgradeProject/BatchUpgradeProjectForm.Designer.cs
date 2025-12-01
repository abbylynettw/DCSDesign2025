namespace WinformUI
{
    partial class BatchUpgradeProjectForm
    {
        private System.Windows.Forms.Button btnStartProcessing;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;

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
            this.btnStartProcessing = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.tabPageWord = new System.Windows.Forms.TabPage();
            this.grpWordOptions = new System.Windows.Forms.GroupBox();
            this.chk页眉标题 = new System.Windows.Forms.CheckBox();
            this.chk内部编码 = new System.Windows.Forms.CheckBox();
            this.chk外部编码 = new System.Windows.Forms.CheckBox();
            this.chk标题替换 = new System.Windows.Forms.CheckBox();
            this.chk项目编号 = new System.Windows.Forms.CheckBox();
            this.chk项目名称 = new System.Windows.Forms.CheckBox();
            this.txt_projectName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chk页眉版本状态 = new System.Windows.Forms.CheckBox();
            this.chk页眉内部编码 = new System.Windows.Forms.CheckBox();
            this.chk页眉总页码 = new System.Windows.Forms.CheckBox();
            this.chk编校审批表格重置 = new System.Windows.Forms.CheckBox();
            this.chk修改记录重置 = new System.Windows.Forms.CheckBox();
            this.txt_ProjectNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_InternalCode = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_versionState = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_totalPage = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tabPageDWG = new System.Windows.Forms.TabPage();
            this.grpDwgOptions = new System.Windows.Forms.GroupBox();
            this.chkDwgBlockAttributes = new System.Windows.Forms.CheckBox();
            this.chkDwgTableContent = new System.Windows.Forms.CheckBox();
            this.chkDwgText = new System.Windows.Forms.CheckBox();
            this.chkdowngradeA = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.tabPageGeneral = new System.Windows.Forms.TabPage();
            this.btnDeleteRule = new System.Windows.Forms.Button();
            this.btnAddRule = new System.Windows.Forms.Button();
            this.dgvReplaceRules = new System.Windows.Forms.DataGridView();
            this.colRuleType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colNewText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colOldText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblReplaceText = new System.Windows.Forms.Label();
            this.lblReplaceRules = new System.Windows.Forms.Label();
            this.chkBackupFiles = new System.Windows.Forms.CheckBox();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.txtFolderPath = new System.Windows.Forms.TextBox();
            this.lblFolderPath = new System.Windows.Forms.Label();
            this.grpFileNameOptions = new System.Windows.Forms.GroupBox();
            this.chkWordFileName = new System.Windows.Forms.CheckBox();
            this.chkDwgFileName = new System.Windows.Forms.CheckBox();
            this.chkFolderName = new System.Windows.Forms.CheckBox();
            this.btn_selectExcel = new System.Windows.Forms.Button();
            this.txt_excel = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btn_import = new System.Windows.Forms.Button();
            this.btn_export = new System.Windows.Forms.Button();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPageWord.SuspendLayout();
            this.grpWordOptions.SuspendLayout();
            this.tabPageDWG.SuspendLayout();
            this.grpDwgOptions.SuspendLayout();
            this.tabPageGeneral.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvReplaceRules)).BeginInit();
            this.grpFileNameOptions.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStartProcessing
            // 
            this.btnStartProcessing.BackColor = System.Drawing.Color.ForestGreen;
            this.btnStartProcessing.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnStartProcessing.Location = new System.Drawing.Point(18, 652);
            this.btnStartProcessing.Margin = new System.Windows.Forms.Padding(4);
            this.btnStartProcessing.Name = "btnStartProcessing";
            this.btnStartProcessing.Size = new System.Drawing.Size(150, 45);
            this.btnStartProcessing.TabIndex = 1;
            this.btnStartProcessing.Text = "开始处理";
            this.btnStartProcessing.UseVisualStyleBackColor = false;
            this.btnStartProcessing.Click += new System.EventHandler(this.btnStartProcessing_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(180, 656);
            this.progressBar.Margin = new System.Windows.Forms.Padding(4);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(990, 41);
            this.progressBar.TabIndex = 2;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(616, 669);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(44, 18);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "就绪";
            // 
            // tabPageWord
            // 
            this.tabPageWord.Controls.Add(this.grpWordOptions);
            this.tabPageWord.Location = new System.Drawing.Point(4, 28);
            this.tabPageWord.Margin = new System.Windows.Forms.Padding(4);
            this.tabPageWord.Name = "tabPageWord";
            this.tabPageWord.Size = new System.Drawing.Size(1148, 568);
            this.tabPageWord.TabIndex = 2;
            this.tabPageWord.Text = "生成A版封面页";
            this.tabPageWord.UseVisualStyleBackColor = true;
            // 
            // grpWordOptions
            // 
            this.grpWordOptions.Controls.Add(this.label7);
            this.grpWordOptions.Controls.Add(this.txt_totalPage);
            this.grpWordOptions.Controls.Add(this.label6);
            this.grpWordOptions.Controls.Add(this.label5);
            this.grpWordOptions.Controls.Add(this.txt_versionState);
            this.grpWordOptions.Controls.Add(this.label4);
            this.grpWordOptions.Controls.Add(this.txt_InternalCode);
            this.grpWordOptions.Controls.Add(this.label3);
            this.grpWordOptions.Controls.Add(this.label2);
            this.grpWordOptions.Controls.Add(this.txt_ProjectNo);
            this.grpWordOptions.Controls.Add(this.chk修改记录重置);
            this.grpWordOptions.Controls.Add(this.chk编校审批表格重置);
            this.grpWordOptions.Controls.Add(this.chk页眉总页码);
            this.grpWordOptions.Controls.Add(this.chk页眉内部编码);
            this.grpWordOptions.Controls.Add(this.chk页眉版本状态);
            this.grpWordOptions.Controls.Add(this.label1);
            this.grpWordOptions.Controls.Add(this.txt_projectName);
            this.grpWordOptions.Controls.Add(this.chk项目名称);
            this.grpWordOptions.Controls.Add(this.chk项目编号);
            this.grpWordOptions.Controls.Add(this.chk标题替换);
            this.grpWordOptions.Controls.Add(this.chk外部编码);
            this.grpWordOptions.Controls.Add(this.chk内部编码);
            this.grpWordOptions.Controls.Add(this.chk页眉标题);
            this.grpWordOptions.Location = new System.Drawing.Point(22, 22);
            this.grpWordOptions.Margin = new System.Windows.Forms.Padding(4);
            this.grpWordOptions.Name = "grpWordOptions";
            this.grpWordOptions.Padding = new System.Windows.Forms.Padding(4);
            this.grpWordOptions.Size = new System.Drawing.Size(1071, 524);
            this.grpWordOptions.TabIndex = 0;
            this.grpWordOptions.TabStop = false;
            this.grpWordOptions.Text = "书签替换";
            // 
            // chk页眉标题
            // 
            this.chk页眉标题.AutoSize = true;
            this.chk页眉标题.Checked = true;
            this.chk页眉标题.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk页眉标题.Location = new System.Drawing.Point(234, 405);
            this.chk页眉标题.Margin = new System.Windows.Forms.Padding(4);
            this.chk页眉标题.Name = "chk页眉标题";
            this.chk页眉标题.Size = new System.Drawing.Size(196, 22);
            this.chk页眉标题.TabIndex = 5;
            this.chk页眉标题.Text = "页眉标题->文字替换";
            this.chk页眉标题.UseVisualStyleBackColor = true;
            // 
            // chk内部编码
            // 
            this.chk内部编码.AutoSize = true;
            this.chk内部编码.Checked = true;
            this.chk内部编码.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk内部编码.Location = new System.Drawing.Point(31, 118);
            this.chk内部编码.Margin = new System.Windows.Forms.Padding(4);
            this.chk内部编码.Name = "chk内部编码";
            this.chk内部编码.Size = new System.Drawing.Size(142, 22);
            this.chk内部编码.TabIndex = 4;
            this.chk内部编码.Text = "内部编码修改";
            this.chk内部编码.UseVisualStyleBackColor = true;
            // 
            // chk外部编码
            // 
            this.chk外部编码.AutoSize = true;
            this.chk外部编码.Checked = true;
            this.chk外部编码.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk外部编码.Location = new System.Drawing.Point(31, 337);
            this.chk外部编码.Margin = new System.Windows.Forms.Padding(4);
            this.chk外部编码.Name = "chk外部编码";
            this.chk外部编码.Size = new System.Drawing.Size(142, 22);
            this.chk外部编码.TabIndex = 3;
            this.chk外部编码.Text = "外部编码修改";
            this.chk外部编码.UseVisualStyleBackColor = true;
            // 
            // chk标题替换
            // 
            this.chk标题替换.AutoSize = true;
            this.chk标题替换.Checked = true;
            this.chk标题替换.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk标题替换.Location = new System.Drawing.Point(31, 405);
            this.chk标题替换.Margin = new System.Windows.Forms.Padding(4);
            this.chk标题替换.Name = "chk标题替换";
            this.chk标题替换.Size = new System.Drawing.Size(160, 22);
            this.chk标题替换.TabIndex = 2;
            this.chk标题替换.Text = "标题->文字替换";
            this.chk标题替换.UseVisualStyleBackColor = true;
            // 
            // chk项目编号
            // 
            this.chk项目编号.AutoSize = true;
            this.chk项目编号.Location = new System.Drawing.Point(31, 70);
            this.chk项目编号.Margin = new System.Windows.Forms.Padding(4);
            this.chk项目编号.Name = "chk项目编号";
            this.chk项目编号.Size = new System.Drawing.Size(142, 22);
            this.chk项目编号.TabIndex = 1;
            this.chk项目编号.Text = "项目编号修改";
            this.chk项目编号.UseVisualStyleBackColor = true;
            // 
            // chk项目名称
            // 
            this.chk项目名称.AutoSize = true;
            this.chk项目名称.Location = new System.Drawing.Point(31, 27);
            this.chk项目名称.Margin = new System.Windows.Forms.Padding(4);
            this.chk项目名称.Name = "chk项目名称";
            this.chk项目名称.Size = new System.Drawing.Size(142, 22);
            this.chk项目名称.TabIndex = 0;
            this.chk项目名称.Text = "项目名称修改";
            this.chk项目名称.UseVisualStyleBackColor = true;
            // 
            // txt_projectName
            // 
            this.txt_projectName.Location = new System.Drawing.Point(265, 25);
            this.txt_projectName.Margin = new System.Windows.Forms.Padding(4);
            this.txt_projectName.Name = "txt_projectName";
            this.txt_projectName.Size = new System.Drawing.Size(589, 28);
            this.txt_projectName.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(231, 31);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(26, 18);
            this.label1.TabIndex = 7;
            this.label1.Text = "->";
            // 
            // chk页眉版本状态
            // 
            this.chk页眉版本状态.AutoSize = true;
            this.chk页眉版本状态.Checked = true;
            this.chk页眉版本状态.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk页眉版本状态.Location = new System.Drawing.Point(31, 206);
            this.chk页眉版本状态.Margin = new System.Windows.Forms.Padding(4);
            this.chk页眉版本状态.Name = "chk页眉版本状态";
            this.chk页眉版本状态.Size = new System.Drawing.Size(178, 22);
            this.chk页眉版本状态.TabIndex = 8;
            this.chk页眉版本状态.Text = "页眉版本状态修改";
            this.chk页眉版本状态.UseVisualStyleBackColor = true;
            // 
            // chk页眉内部编码
            // 
            this.chk页眉内部编码.AutoSize = true;
            this.chk页眉内部编码.Checked = true;
            this.chk页眉内部编码.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk页眉内部编码.Location = new System.Drawing.Point(31, 164);
            this.chk页眉内部编码.Margin = new System.Windows.Forms.Padding(4);
            this.chk页眉内部编码.Name = "chk页眉内部编码";
            this.chk页眉内部编码.Size = new System.Drawing.Size(178, 22);
            this.chk页眉内部编码.TabIndex = 9;
            this.chk页眉内部编码.Text = "页眉内部编码修改";
            this.chk页眉内部编码.UseVisualStyleBackColor = true;
            // 
            // chk页眉总页码
            // 
            this.chk页眉总页码.AutoSize = true;
            this.chk页眉总页码.Location = new System.Drawing.Point(31, 252);
            this.chk页眉总页码.Margin = new System.Windows.Forms.Padding(4);
            this.chk页眉总页码.Name = "chk页眉总页码";
            this.chk页眉总页码.Size = new System.Drawing.Size(160, 22);
            this.chk页眉总页码.TabIndex = 10;
            this.chk页眉总页码.Text = "页眉总页码修改";
            this.chk页眉总页码.UseVisualStyleBackColor = true;
            // 
            // chk编校审批表格重置
            // 
            this.chk编校审批表格重置.AutoSize = true;
            this.chk编校审批表格重置.Checked = true;
            this.chk编校审批表格重置.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk编校审批表格重置.Location = new System.Drawing.Point(714, 405);
            this.chk编校审批表格重置.Margin = new System.Windows.Forms.Padding(4);
            this.chk编校审批表格重置.Name = "chk编校审批表格重置";
            this.chk编校审批表格重置.Size = new System.Drawing.Size(223, 22);
            this.chk编校审批表格重置.TabIndex = 11;
            this.chk编校审批表格重置.Text = "编校审批表格重置为A版";
            this.chk编校审批表格重置.UseVisualStyleBackColor = true;
            // 
            // chk修改记录重置
            // 
            this.chk修改记录重置.AutoSize = true;
            this.chk修改记录重置.Checked = true;
            this.chk修改记录重置.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk修改记录重置.Location = new System.Drawing.Point(471, 405);
            this.chk修改记录重置.Margin = new System.Windows.Forms.Padding(4);
            this.chk修改记录重置.Name = "chk修改记录重置";
            this.chk修改记录重置.Size = new System.Drawing.Size(223, 22);
            this.chk修改记录重置.TabIndex = 12;
            this.chk修改记录重置.Text = "修改记录表格重置为A版";
            this.chk修改记录重置.UseVisualStyleBackColor = true;
            // 
            // txt_ProjectNo
            // 
            this.txt_ProjectNo.Location = new System.Drawing.Point(265, 64);
            this.txt_ProjectNo.Margin = new System.Windows.Forms.Padding(4);
            this.txt_ProjectNo.Name = "txt_ProjectNo";
            this.txt_ProjectNo.Size = new System.Drawing.Size(589, 28);
            this.txt_ProjectNo.TabIndex = 13;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(231, 70);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(26, 18);
            this.label2.TabIndex = 14;
            this.label2.Text = "->";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(231, 337);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(242, 18);
            this.label3.TabIndex = 16;
            this.label3.Text = "-> 根据标题-外部编码表计算";
            // 
            // txt_InternalCode
            // 
            this.txt_InternalCode.Location = new System.Drawing.Point(265, 112);
            this.txt_InternalCode.Margin = new System.Windows.Forms.Padding(4);
            this.txt_InternalCode.Name = "txt_InternalCode";
            this.txt_InternalCode.Size = new System.Drawing.Size(589, 28);
            this.txt_InternalCode.TabIndex = 17;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(231, 118);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(26, 18);
            this.label4.TabIndex = 18;
            this.label4.Text = "->";
            // 
            // txt_versionState
            // 
            this.txt_versionState.Location = new System.Drawing.Point(265, 202);
            this.txt_versionState.Margin = new System.Windows.Forms.Padding(4);
            this.txt_versionState.Name = "txt_versionState";
            this.txt_versionState.Size = new System.Drawing.Size(589, 28);
            this.txt_versionState.TabIndex = 19;
            this.txt_versionState.Text = "A/CFC";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(231, 206);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(26, 18);
            this.label5.TabIndex = 20;
            this.label5.Text = "->";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(231, 163);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 18);
            this.label6.TabIndex = 22;
            this.label6.Text = "->  同上";
            // 
            // txt_totalPage
            // 
            this.txt_totalPage.Location = new System.Drawing.Point(265, 246);
            this.txt_totalPage.Margin = new System.Windows.Forms.Padding(4);
            this.txt_totalPage.Name = "txt_totalPage";
            this.txt_totalPage.Size = new System.Drawing.Size(589, 28);
            this.txt_totalPage.TabIndex = 23;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(231, 250);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(26, 18);
            this.label7.TabIndex = 24;
            this.label7.Text = "->";
            // 
            // tabPageDWG
            // 
            this.tabPageDWG.Controls.Add(this.label12);
            this.tabPageDWG.Controls.Add(this.label11);
            this.tabPageDWG.Controls.Add(this.label10);
            this.tabPageDWG.Controls.Add(this.label9);
            this.tabPageDWG.Controls.Add(this.chkdowngradeA);
            this.tabPageDWG.Controls.Add(this.grpDwgOptions);
            this.tabPageDWG.Location = new System.Drawing.Point(4, 28);
            this.tabPageDWG.Margin = new System.Windows.Forms.Padding(4);
            this.tabPageDWG.Name = "tabPageDWG";
            this.tabPageDWG.Padding = new System.Windows.Forms.Padding(4);
            this.tabPageDWG.Size = new System.Drawing.Size(1148, 568);
            this.tabPageDWG.TabIndex = 1;
            this.tabPageDWG.Text = "DWG文件设置";
            this.tabPageDWG.UseVisualStyleBackColor = true;
            // 
            // grpDwgOptions
            // 
            this.grpDwgOptions.Controls.Add(this.chkDwgText);
            this.grpDwgOptions.Controls.Add(this.chkDwgTableContent);
            this.grpDwgOptions.Controls.Add(this.chkDwgBlockAttributes);
            this.grpDwgOptions.Location = new System.Drawing.Point(22, 22);
            this.grpDwgOptions.Margin = new System.Windows.Forms.Padding(4);
            this.grpDwgOptions.Name = "grpDwgOptions";
            this.grpDwgOptions.Padding = new System.Windows.Forms.Padding(4);
            this.grpDwgOptions.Size = new System.Drawing.Size(1071, 150);
            this.grpDwgOptions.TabIndex = 0;
            this.grpDwgOptions.TabStop = false;
            this.grpDwgOptions.Text = "DWG替换对象";
            // 
            // chkDwgBlockAttributes
            // 
            this.chkDwgBlockAttributes.AutoSize = true;
            this.chkDwgBlockAttributes.Checked = true;
            this.chkDwgBlockAttributes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwgBlockAttributes.Location = new System.Drawing.Point(31, 97);
            this.chkDwgBlockAttributes.Margin = new System.Windows.Forms.Padding(4);
            this.chkDwgBlockAttributes.Name = "chkDwgBlockAttributes";
            this.chkDwgBlockAttributes.Size = new System.Drawing.Size(196, 22);
            this.chkDwgBlockAttributes.TabIndex = 2;
            this.chkDwgBlockAttributes.Text = "块属性（图框属性）";
            this.chkDwgBlockAttributes.UseVisualStyleBackColor = true;
            // 
            // chkDwgTableContent
            // 
            this.chkDwgTableContent.AutoSize = true;
            this.chkDwgTableContent.Checked = true;
            this.chkDwgTableContent.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwgTableContent.Location = new System.Drawing.Point(31, 64);
            this.chkDwgTableContent.Margin = new System.Windows.Forms.Padding(4);
            this.chkDwgTableContent.Name = "chkDwgTableContent";
            this.chkDwgTableContent.Size = new System.Drawing.Size(106, 22);
            this.chkDwgTableContent.TabIndex = 1;
            this.chkDwgTableContent.Text = "表格内容";
            this.chkDwgTableContent.UseVisualStyleBackColor = true;
            // 
            // chkDwgText
            // 
            this.chkDwgText.AutoSize = true;
            this.chkDwgText.Checked = true;
            this.chkDwgText.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwgText.Location = new System.Drawing.Point(31, 31);
            this.chkDwgText.Margin = new System.Windows.Forms.Padding(4);
            this.chkDwgText.Name = "chkDwgText";
            this.chkDwgText.Size = new System.Drawing.Size(196, 22);
            this.chkDwgText.TabIndex = 0;
            this.chkDwgText.Text = "文字（单行、多行）";
            this.chkDwgText.UseVisualStyleBackColor = true;
            // 
            // chkdowngradeA
            // 
            this.chkdowngradeA.AutoSize = true;
            this.chkdowngradeA.Checked = true;
            this.chkdowngradeA.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkdowngradeA.Location = new System.Drawing.Point(31, 244);
            this.chkdowngradeA.Margin = new System.Windows.Forms.Padding(4);
            this.chkdowngradeA.Name = "chkdowngradeA";
            this.chkdowngradeA.Size = new System.Drawing.Size(115, 22);
            this.chkdowngradeA.TabIndex = 1;
            this.chkdowngradeA.Text = "降成A版本";
            this.chkdowngradeA.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(50, 273);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(233, 18);
            this.label9.TabIndex = 4;
            this.label9.Text = "----目录页表格版本列改为A";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(50, 304);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(233, 18);
            this.label10.TabIndex = 5;
            this.label10.Text = "----图框版本、总版本改为A";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(50, 337);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(215, 18);
            this.label11.TabIndex = 6;
            this.label11.Text = "----清空图框B C D E F栏";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(50, 368);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(206, 18);
            this.label12.TabIndex = 7;
            this.label12.Text = "----修改图框日期 月/年";
            // 
            // tabPageGeneral
            // 
            this.tabPageGeneral.Controls.Add(this.btn_export);
            this.tabPageGeneral.Controls.Add(this.btn_import);
            this.tabPageGeneral.Controls.Add(this.label8);
            this.tabPageGeneral.Controls.Add(this.txt_excel);
            this.tabPageGeneral.Controls.Add(this.txtFolderPath);
            this.tabPageGeneral.Controls.Add(this.btn_selectExcel);
            this.tabPageGeneral.Controls.Add(this.grpFileNameOptions);
            this.tabPageGeneral.Controls.Add(this.lblFolderPath);
            this.tabPageGeneral.Controls.Add(this.btnSelectFolder);
            this.tabPageGeneral.Controls.Add(this.chkBackupFiles);
            this.tabPageGeneral.Controls.Add(this.lblReplaceRules);
            this.tabPageGeneral.Controls.Add(this.lblReplaceText);
            this.tabPageGeneral.Controls.Add(this.dgvReplaceRules);
            this.tabPageGeneral.Controls.Add(this.btnAddRule);
            this.tabPageGeneral.Controls.Add(this.btnDeleteRule);
            this.tabPageGeneral.Location = new System.Drawing.Point(4, 28);
            this.tabPageGeneral.Margin = new System.Windows.Forms.Padding(4);
            this.tabPageGeneral.Name = "tabPageGeneral";
            this.tabPageGeneral.Padding = new System.Windows.Forms.Padding(4);
            this.tabPageGeneral.Size = new System.Drawing.Size(1148, 568);
            this.tabPageGeneral.TabIndex = 0;
            this.tabPageGeneral.Text = "常规设置";
            this.tabPageGeneral.UseVisualStyleBackColor = true;
            // 
            // btnDeleteRule
            // 
            this.btnDeleteRule.Location = new System.Drawing.Point(980, 250);
            this.btnDeleteRule.Margin = new System.Windows.Forms.Padding(4);
            this.btnDeleteRule.Name = "btnDeleteRule";
            this.btnDeleteRule.Size = new System.Drawing.Size(112, 34);
            this.btnDeleteRule.TabIndex = 8;
            this.btnDeleteRule.Text = "删除";
            this.btnDeleteRule.UseVisualStyleBackColor = true;
            this.btnDeleteRule.Click += new System.EventHandler(this.btnDeleteRule_Click);
            // 
            // btnAddRule
            // 
            this.btnAddRule.Location = new System.Drawing.Point(980, 208);
            this.btnAddRule.Margin = new System.Windows.Forms.Padding(4);
            this.btnAddRule.Name = "btnAddRule";
            this.btnAddRule.Size = new System.Drawing.Size(112, 34);
            this.btnAddRule.TabIndex = 7;
            this.btnAddRule.Text = "添加";
            this.btnAddRule.UseVisualStyleBackColor = true;
            this.btnAddRule.Click += new System.EventHandler(this.btnAddRule_Click);
            // 
            // dgvReplaceRules
            // 
            this.dgvReplaceRules.AllowUserToAddRows = false;
            this.dgvReplaceRules.AllowUserToDeleteRows = false;
            this.dgvReplaceRules.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvReplaceRules.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colOldText,
            this.colNewText,
            this.colRuleType});
            this.dgvReplaceRules.Location = new System.Drawing.Point(22, 208);
            this.dgvReplaceRules.Margin = new System.Windows.Forms.Padding(4);
            this.dgvReplaceRules.Name = "dgvReplaceRules";
            this.dgvReplaceRules.RowHeadersWidth = 62;
            this.dgvReplaceRules.Size = new System.Drawing.Size(950, 257);
            this.dgvReplaceRules.TabIndex = 6;
            // 
            // colRuleType
            // 
            this.colRuleType.HeaderText = "规则类型";
            this.colRuleType.Items.AddRange(new object[] {
            "相等",
            "包含"});
            this.colRuleType.MinimumWidth = 8;
            this.colRuleType.Name = "colRuleType";
            this.colRuleType.Width = 120;
            // 
            // colNewText
            // 
            this.colNewText.HeaderText = "新文本";
            this.colNewText.MinimumWidth = 8;
            this.colNewText.Name = "colNewText";
            this.colNewText.Width = 220;
            // 
            // colOldText
            // 
            this.colOldText.HeaderText = "原文本";
            this.colOldText.MinimumWidth = 8;
            this.colOldText.Name = "colOldText";
            this.colOldText.Width = 225;
            // 
            // lblReplaceText
            // 
            this.lblReplaceText.AutoSize = true;
            this.lblReplaceText.Location = new System.Drawing.Point(121, 178);
            this.lblReplaceText.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblReplaceText.Name = "lblReplaceText";
            this.lblReplaceText.Size = new System.Drawing.Size(494, 18);
            this.lblReplaceText.TabIndex = 5;
            this.lblReplaceText.Text = "在下表中指定要替换的文本和新文本，并选择相应的规则类型";
            // 
            // lblReplaceRules
            // 
            this.lblReplaceRules.AutoSize = true;
            this.lblReplaceRules.Location = new System.Drawing.Point(23, 178);
            this.lblReplaceRules.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblReplaceRules.Name = "lblReplaceRules";
            this.lblReplaceRules.Size = new System.Drawing.Size(89, 18);
            this.lblReplaceRules.TabIndex = 4;
            this.lblReplaceRules.Text = "替换规则:";
            // 
            // chkBackupFiles
            // 
            this.chkBackupFiles.AutoSize = true;
            this.chkBackupFiles.Checked = true;
            this.chkBackupFiles.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBackupFiles.Location = new System.Drawing.Point(29, 132);
            this.chkBackupFiles.Margin = new System.Windows.Forms.Padding(4);
            this.chkBackupFiles.Name = "chkBackupFiles";
            this.chkBackupFiles.Size = new System.Drawing.Size(160, 22);
            this.chkBackupFiles.TabIndex = 3;
            this.chkBackupFiles.Text = "处理前备份文件";
            this.chkBackupFiles.UseVisualStyleBackColor = true;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(981, 15);
            this.btnSelectFolder.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(112, 34);
            this.btnSelectFolder.TabIndex = 2;
            this.btnSelectFolder.Text = "浏览...";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // txtFolderPath
            // 
            this.txtFolderPath.Location = new System.Drawing.Point(177, 18);
            this.txtFolderPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtFolderPath.Name = "txtFolderPath";
            this.txtFolderPath.Size = new System.Drawing.Size(793, 28);
            this.txtFolderPath.TabIndex = 1;
            // 
            // lblFolderPath
            // 
            this.lblFolderPath.AutoSize = true;
            this.lblFolderPath.Location = new System.Drawing.Point(26, 26);
            this.lblFolderPath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFolderPath.Name = "lblFolderPath";
            this.lblFolderPath.Size = new System.Drawing.Size(107, 18);
            this.lblFolderPath.TabIndex = 0;
            this.lblFolderPath.Text = "图纸文件夹:";
            // 
            // grpFileNameOptions
            // 
            this.grpFileNameOptions.Controls.Add(this.chkFolderName);
            this.grpFileNameOptions.Controls.Add(this.chkDwgFileName);
            this.grpFileNameOptions.Controls.Add(this.chkWordFileName);
            this.grpFileNameOptions.Location = new System.Drawing.Point(22, 482);
            this.grpFileNameOptions.Margin = new System.Windows.Forms.Padding(4);
            this.grpFileNameOptions.Name = "grpFileNameOptions";
            this.grpFileNameOptions.Padding = new System.Windows.Forms.Padding(4);
            this.grpFileNameOptions.Size = new System.Drawing.Size(948, 68);
            this.grpFileNameOptions.TabIndex = 11;
            this.grpFileNameOptions.TabStop = false;
            this.grpFileNameOptions.Text = "文件名处理选项";
            // 
            // chkWordFileName
            // 
            this.chkWordFileName.AutoSize = true;
            this.chkWordFileName.Checked = true;
            this.chkWordFileName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkWordFileName.Location = new System.Drawing.Point(367, 29);
            this.chkWordFileName.Margin = new System.Windows.Forms.Padding(4);
            this.chkWordFileName.Name = "chkWordFileName";
            this.chkWordFileName.Size = new System.Drawing.Size(124, 22);
            this.chkWordFileName.TabIndex = 1;
            this.chkWordFileName.Text = "Word文件名";
            this.chkWordFileName.UseVisualStyleBackColor = true;
            // 
            // chkDwgFileName
            // 
            this.chkDwgFileName.AutoSize = true;
            this.chkDwgFileName.Checked = true;
            this.chkDwgFileName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwgFileName.Location = new System.Drawing.Point(183, 29);
            this.chkDwgFileName.Margin = new System.Windows.Forms.Padding(4);
            this.chkDwgFileName.Name = "chkDwgFileName";
            this.chkDwgFileName.Size = new System.Drawing.Size(151, 22);
            this.chkDwgFileName.TabIndex = 0;
            this.chkDwgFileName.Text = "DWG图纸文件名";
            this.chkDwgFileName.UseVisualStyleBackColor = true;
            // 
            // chkFolderName
            // 
            this.chkFolderName.AutoSize = true;
            this.chkFolderName.Checked = true;
            this.chkFolderName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFolderName.Location = new System.Drawing.Point(33, 29);
            this.chkFolderName.Margin = new System.Windows.Forms.Padding(4);
            this.chkFolderName.Name = "chkFolderName";
            this.chkFolderName.Size = new System.Drawing.Size(124, 22);
            this.chkFolderName.TabIndex = 2;
            this.chkFolderName.Text = "文件夹名称";
            this.chkFolderName.UseVisualStyleBackColor = true;
            // 
            // btn_selectExcel
            // 
            this.btn_selectExcel.Location = new System.Drawing.Point(981, 69);
            this.btn_selectExcel.Margin = new System.Windows.Forms.Padding(4);
            this.btn_selectExcel.Name = "btn_selectExcel";
            this.btn_selectExcel.Size = new System.Drawing.Size(112, 34);
            this.btn_selectExcel.TabIndex = 14;
            this.btn_selectExcel.Text = "浏览...";
            this.btn_selectExcel.UseVisualStyleBackColor = true;
            this.btn_selectExcel.Click += new System.EventHandler(this.btn_selectExcel_Click);
            // 
            // txt_excel
            // 
            this.txt_excel.Location = new System.Drawing.Point(177, 69);
            this.txt_excel.Margin = new System.Windows.Forms.Padding(4);
            this.txt_excel.Name = "txt_excel";
            this.txt_excel.Size = new System.Drawing.Size(793, 28);
            this.txt_excel.TabIndex = 13;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(26, 77);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(143, 18);
            this.label8.TabIndex = 12;
            this.label8.Text = "标题-外部编码表";
            // 
            // btn_import
            // 
            this.btn_import.Location = new System.Drawing.Point(980, 292);
            this.btn_import.Margin = new System.Windows.Forms.Padding(4);
            this.btn_import.Name = "btn_import";
            this.btn_import.Size = new System.Drawing.Size(112, 34);
            this.btn_import.TabIndex = 15;
            this.btn_import.Text = "导入";
            this.btn_import.UseVisualStyleBackColor = true;
            this.btn_import.Click += new System.EventHandler(this.btn_import_Click);
            // 
            // btn_export
            // 
            this.btn_export.Location = new System.Drawing.Point(981, 334);
            this.btn_export.Margin = new System.Windows.Forms.Padding(4);
            this.btn_export.Name = "btn_export";
            this.btn_export.Size = new System.Drawing.Size(112, 34);
            this.btn_export.TabIndex = 16;
            this.btn_export.Text = "导出";
            this.btn_export.UseVisualStyleBackColor = true;
            this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabPageGeneral);
            this.tabControl.Controls.Add(this.tabPageDWG);
            this.tabControl.Controls.Add(this.tabPageWord);
            this.tabControl.Location = new System.Drawing.Point(18, 19);
            this.tabControl.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1156, 600);
            this.tabControl.TabIndex = 0;
            // 
            // BatchUpgradeProjectForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1206, 773);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnStartProcessing);
            this.Controls.Add(this.tabControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "BatchUpgradeProjectForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "工程图纸批量改版工具";
            this.tabPageWord.ResumeLayout(false);
            this.grpWordOptions.ResumeLayout(false);
            this.grpWordOptions.PerformLayout();
            this.tabPageDWG.ResumeLayout(false);
            this.tabPageDWG.PerformLayout();
            this.grpDwgOptions.ResumeLayout(false);
            this.grpDwgOptions.PerformLayout();
            this.tabPageGeneral.ResumeLayout(false);
            this.tabPageGeneral.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvReplaceRules)).EndInit();
            this.grpFileNameOptions.ResumeLayout(false);
            this.grpFileNameOptions.PerformLayout();
            this.tabControl.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabPage tabPageWord;
        private System.Windows.Forms.GroupBox grpWordOptions;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt_totalPage;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_versionState;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_InternalCode;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_ProjectNo;
        private System.Windows.Forms.CheckBox chk修改记录重置;
        private System.Windows.Forms.CheckBox chk编校审批表格重置;
        private System.Windows.Forms.CheckBox chk页眉总页码;
        private System.Windows.Forms.CheckBox chk页眉内部编码;
        private System.Windows.Forms.CheckBox chk页眉版本状态;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_projectName;
        private System.Windows.Forms.CheckBox chk项目名称;
        private System.Windows.Forms.CheckBox chk项目编号;
        private System.Windows.Forms.CheckBox chk标题替换;
        private System.Windows.Forms.CheckBox chk外部编码;
        private System.Windows.Forms.CheckBox chk内部编码;
        private System.Windows.Forms.CheckBox chk页眉标题;
        private System.Windows.Forms.TabPage tabPageDWG;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox chkdowngradeA;
        private System.Windows.Forms.GroupBox grpDwgOptions;
        private System.Windows.Forms.CheckBox chkDwgText;
        private System.Windows.Forms.CheckBox chkDwgTableContent;
        private System.Windows.Forms.CheckBox chkDwgBlockAttributes;
        private System.Windows.Forms.TabPage tabPageGeneral;
        private System.Windows.Forms.Button btn_export;
        private System.Windows.Forms.Button btn_import;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_excel;
        private System.Windows.Forms.TextBox txtFolderPath;
        private System.Windows.Forms.Button btn_selectExcel;
        private System.Windows.Forms.GroupBox grpFileNameOptions;
        private System.Windows.Forms.CheckBox chkFolderName;
        private System.Windows.Forms.CheckBox chkDwgFileName;
        private System.Windows.Forms.CheckBox chkWordFileName;
        private System.Windows.Forms.Label lblFolderPath;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.CheckBox chkBackupFiles;
        private System.Windows.Forms.Label lblReplaceRules;
        private System.Windows.Forms.Label lblReplaceText;
        private System.Windows.Forms.DataGridView dgvReplaceRules;
        private System.Windows.Forms.DataGridViewTextBoxColumn colOldText;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNewText;
        private System.Windows.Forms.DataGridViewComboBoxColumn colRuleType;
        private System.Windows.Forms.Button btnAddRule;
        private System.Windows.Forms.Button btnDeleteRule;
        private System.Windows.Forms.TabControl tabControl;
    }
}