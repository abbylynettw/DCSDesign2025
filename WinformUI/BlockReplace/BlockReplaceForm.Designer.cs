namespace WinformUI.BlockReplace
{
    partial class BlockReplaceForm
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
            this.lblDirectoryInfo = new System.Windows.Forms.Label();
            this.lblSourceBlock = new System.Windows.Forms.Label();
            this.lblTargetBlock = new System.Windows.Forms.Label();
            this.btnSelectDirectory = new System.Windows.Forms.Button();
            this.btnClearDirectory = new System.Windows.Forms.Button();
            this.txtSourceBlockName = new System.Windows.Forms.TextBox();
            this.txtTargetBlockName = new System.Windows.Forms.TextBox();
            this.btnBrowseSourceBlock = new System.Windows.Forms.Button();
            this.btnBrowseTargetBlock = new System.Windows.Forms.Button();
            this.chkPreserveAttributes = new System.Windows.Forms.CheckBox();
            this.chkPreservePosition = new System.Windows.Forms.CheckBox();
            this.chkPreserveRotation = new System.Windows.Forms.CheckBox();
            this.chkPreserveScale = new System.Windows.Forms.CheckBox();
            this.chkCreateBackup = new System.Windows.Forms.CheckBox();
            this.btnPreview = new System.Windows.Forms.Button();
            this.btnStartReplace = new System.Windows.Forms.Button();
            this.btnStop = new System.Windows.Forms.Button();
            this.rtbLog = new System.Windows.Forms.RichTextBox();
            this.btnClearLog = new System.Windows.Forms.Button();
            this.btnSaveReport = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.pbPreviewImage = new System.Windows.Forms.PictureBox();
            this.pbBlockThumbnail = new System.Windows.Forms.PictureBox();
            this.lblBlockThumbnail = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pbPreviewImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbBlockThumbnail)).BeginInit();
            this.SuspendLayout();
            // 
            // lblDirectoryInfo
            // 
            this.lblDirectoryInfo.AutoSize = true;
            this.lblDirectoryInfo.Location = new System.Drawing.Point(12, 15);
            this.lblDirectoryInfo.Name = "lblDirectoryInfo";
            this.lblDirectoryInfo.Size = new System.Drawing.Size(67, 15);
            this.lblDirectoryInfo.TabIndex = 0;
            this.lblDirectoryInfo.Text = "未选择目录";
            // 
            // btnSelectDirectory
            // 
            this.btnSelectDirectory.Location = new System.Drawing.Point(12, 40);
            this.btnSelectDirectory.Name = "btnSelectDirectory";
            this.btnSelectDirectory.Size = new System.Drawing.Size(100, 30);
            this.btnSelectDirectory.TabIndex = 1;
            this.btnSelectDirectory.Text = "选择目录";
            this.btnSelectDirectory.UseVisualStyleBackColor = true;
            // 
            // btnClearDirectory
            // 
            this.btnClearDirectory.Location = new System.Drawing.Point(118, 40);
            this.btnClearDirectory.Name = "btnClearDirectory";
            this.btnClearDirectory.Size = new System.Drawing.Size(100, 30);
            this.btnClearDirectory.TabIndex = 2;
            this.btnClearDirectory.Text = "清空选择";
            this.btnClearDirectory.UseVisualStyleBackColor = true;
            // 
            // lblSourceBlock
            // 
            this.lblSourceBlock.AutoSize = true;
            this.lblSourceBlock.Location = new System.Drawing.Point(12, 80);
            this.lblSourceBlock.Name = "lblSourceBlock";
            this.lblSourceBlock.Size = new System.Drawing.Size(127, 15);
            this.lblSourceBlock.TabIndex = 20;
            this.lblSourceBlock.Text = "源块文件（选择.dwg文件）:";
            // 
            // lblTargetBlock
            // 
            this.lblTargetBlock.AutoSize = true;
            this.lblTargetBlock.Location = new System.Drawing.Point(12, 125);
            this.lblTargetBlock.Name = "lblTargetBlock";
            this.lblTargetBlock.Size = new System.Drawing.Size(82, 15);
            this.lblTargetBlock.TabIndex = 21;
            this.lblTargetBlock.Text = "目标块名称:";
            // 
            // txtSourceBlockName
            // 
            this.txtSourceBlockName.Location = new System.Drawing.Point(12, 100);
            this.txtSourceBlockName.Name = "txtSourceBlockName";
            this.txtSourceBlockName.Size = new System.Drawing.Size(200, 25);
            this.txtSourceBlockName.TabIndex = 3;
            // 
            // txtTargetBlockName
            // 
            this.txtTargetBlockName.Location = new System.Drawing.Point(12, 140);
            this.txtTargetBlockName.Name = "txtTargetBlockName";
            this.txtTargetBlockName.Size = new System.Drawing.Size(200, 25);
            this.txtTargetBlockName.TabIndex = 4;
            // 
            // btnBrowseSourceBlock
            // 
            this.btnBrowseSourceBlock.Location = new System.Drawing.Point(218, 100);
            this.btnBrowseSourceBlock.Name = "btnBrowseSourceBlock";
            this.btnBrowseSourceBlock.Size = new System.Drawing.Size(30, 25);
            this.btnBrowseSourceBlock.TabIndex = 5;
            this.btnBrowseSourceBlock.Text = "...";
            this.btnBrowseSourceBlock.UseVisualStyleBackColor = true;
            // 
            // btnBrowseTargetBlock
            // 
            this.btnBrowseTargetBlock.Location = new System.Drawing.Point(218, 140);
            this.btnBrowseTargetBlock.Name = "btnBrowseTargetBlock";
            this.btnBrowseTargetBlock.Size = new System.Drawing.Size(30, 25);
            this.btnBrowseTargetBlock.TabIndex = 6;
            this.btnBrowseTargetBlock.Text = "...";
            this.btnBrowseTargetBlock.UseVisualStyleBackColor = true;
            // 
            // chkPreserveAttributes
            // 
            this.chkPreserveAttributes.AutoSize = true;
            this.chkPreserveAttributes.Checked = true;
            this.chkPreserveAttributes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPreserveAttributes.Location = new System.Drawing.Point(12, 180);
            this.chkPreserveAttributes.Name = "chkPreserveAttributes";
            this.chkPreserveAttributes.Size = new System.Drawing.Size(89, 19);
            this.chkPreserveAttributes.TabIndex = 7;
            this.chkPreserveAttributes.Text = "保留属性";
            this.chkPreserveAttributes.UseVisualStyleBackColor = true;
            // 
            // chkPreservePosition
            // 
            this.chkPreservePosition.AutoSize = true;
            this.chkPreservePosition.Checked = true;
            this.chkPreservePosition.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPreservePosition.Location = new System.Drawing.Point(120, 180);
            this.chkPreservePosition.Name = "chkPreservePosition";
            this.chkPreservePosition.Size = new System.Drawing.Size(89, 19);
            this.chkPreservePosition.TabIndex = 8;
            this.chkPreservePosition.Text = "保留位置";
            this.chkPreservePosition.UseVisualStyleBackColor = true;
            // 
            // chkPreserveRotation
            // 
            this.chkPreserveRotation.AutoSize = true;
            this.chkPreserveRotation.Checked = true;
            this.chkPreserveRotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPreserveRotation.Location = new System.Drawing.Point(12, 205);
            this.chkPreserveRotation.Name = "chkPreserveRotation";
            this.chkPreserveRotation.Size = new System.Drawing.Size(89, 19);
            this.chkPreserveRotation.TabIndex = 9;
            this.chkPreserveRotation.Text = "保留旋转";
            this.chkPreserveRotation.UseVisualStyleBackColor = true;
            // 
            // chkPreserveScale
            // 
            this.chkPreserveScale.AutoSize = true;
            this.chkPreserveScale.Checked = true;
            this.chkPreserveScale.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPreserveScale.Location = new System.Drawing.Point(120, 205);
            this.chkPreserveScale.Name = "chkPreserveScale";
            this.chkPreserveScale.Size = new System.Drawing.Size(89, 19);
            this.chkPreserveScale.TabIndex = 10;
            this.chkPreserveScale.Text = "保留缩放";
            this.chkPreserveScale.UseVisualStyleBackColor = true;
            // 
            // chkCreateBackup
            // 
            this.chkCreateBackup.AutoSize = true;
            this.chkCreateBackup.Checked = true;
            this.chkCreateBackup.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCreateBackup.Location = new System.Drawing.Point(230, 180);
            this.chkCreateBackup.Name = "chkCreateBackup";
            this.chkCreateBackup.Size = new System.Drawing.Size(89, 19);
            this.chkCreateBackup.TabIndex = 11;
            this.chkCreateBackup.Text = "创建备份";
            this.chkCreateBackup.UseVisualStyleBackColor = true;
            // 
            // btnPreview
            // 
            this.btnPreview.Location = new System.Drawing.Point(12, 240);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(100, 35);
            this.btnPreview.TabIndex = 12;
            this.btnPreview.Text = "预览";
            this.btnPreview.UseVisualStyleBackColor = true;
            // 
            // btnStartReplace
            // 
            this.btnStartReplace.Location = new System.Drawing.Point(118, 240);
            this.btnStartReplace.Name = "btnStartReplace";
            this.btnStartReplace.Size = new System.Drawing.Size(100, 35);
            this.btnStartReplace.TabIndex = 13;
            this.btnStartReplace.Text = "开始替换";
            this.btnStartReplace.UseVisualStyleBackColor = true;
            // 
            // btnStop
            // 
            this.btnStop.Enabled = false;
            this.btnStop.Location = new System.Drawing.Point(224, 240);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(100, 35);
            this.btnStop.TabIndex = 14;
            this.btnStop.Text = "停止";
            this.btnStop.UseVisualStyleBackColor = true;
            // 
            // lblBlockThumbnail
            // 
            this.lblBlockThumbnail.AutoSize = true;
            this.lblBlockThumbnail.Location = new System.Drawing.Point(10, 150);
            this.lblBlockThumbnail.Name = "lblBlockThumbnail";
            this.lblBlockThumbnail.Size = new System.Drawing.Size(65, 12);
            this.lblBlockThumbnail.TabIndex = 23;
            this.lblBlockThumbnail.Text = "块缩略图:";
            // 
            // pbBlockThumbnail
            // 
            this.pbBlockThumbnail.Location = new System.Drawing.Point(10, 170);
            this.pbBlockThumbnail.Name = "pbBlockThumbnail";
            this.pbBlockThumbnail.Size = new System.Drawing.Size(330, 200);
            this.pbBlockThumbnail.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbBlockThumbnail.TabIndex = 24;
            this.pbBlockThumbnail.TabStop = false;
            // 
            // pbPreviewImage
            // 
            this.pbPreviewImage.Location = new System.Drawing.Point(350, 40);
            this.pbPreviewImage.Name = "pbPreviewImage";
            this.pbPreviewImage.Size = new System.Drawing.Size(200, 200);
            this.pbPreviewImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbPreviewImage.TabIndex = 22;
            this.pbPreviewImage.TabStop = false;
            // 
            // rtbLog
            // 
            this.rtbLog.Location = new System.Drawing.Point(560, 40);
            this.rtbLog.Name = "rtbLog";
            this.rtbLog.Size = new System.Drawing.Size(290, 300);
            this.rtbLog.TabIndex = 15;
            this.rtbLog.Text = "";
            // 
            // btnClearLog
            // 
            this.btnClearLog.Location = new System.Drawing.Point(350, 350);
            this.btnClearLog.Name = "btnClearLog";
            this.btnClearLog.Size = new System.Drawing.Size(100, 30);
            this.btnClearLog.TabIndex = 16;
            this.btnClearLog.Text = "清空日志";
            this.btnClearLog.UseVisualStyleBackColor = true;
            // 
            // btnSaveReport
            // 
            this.btnSaveReport.Location = new System.Drawing.Point(456, 350);
            this.btnSaveReport.Name = "btnSaveReport";
            this.btnSaveReport.Size = new System.Drawing.Size(100, 30);
            this.btnSaveReport.TabIndex = 17;
            this.btnSaveReport.Text = "保存报告";
            this.btnSaveReport.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 290);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(312, 23);
            this.progressBar.TabIndex = 18;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(12, 320);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(52, 15);
            this.lblProgress.TabIndex = 19;
            this.lblProgress.Text = "就绪...";
            // 
            // BlockReplaceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(870, 400);
            this.Controls.Add(this.pbBlockThumbnail);
            this.Controls.Add(this.lblBlockThumbnail);
            this.Controls.Add(this.pbPreviewImage);
            this.Controls.Add(this.lblTargetBlock);
            this.Controls.Add(this.lblSourceBlock);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnSaveReport);
            this.Controls.Add(this.btnClearLog);
            this.Controls.Add(this.rtbLog);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.btnStartReplace);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.chkCreateBackup);
            this.Controls.Add(this.chkPreserveScale);
            this.Controls.Add(this.chkPreserveRotation);
            this.Controls.Add(this.chkPreservePosition);
            this.Controls.Add(this.chkPreserveAttributes);
            this.Controls.Add(this.btnBrowseTargetBlock);
            this.Controls.Add(this.btnBrowseSourceBlock);
            this.Controls.Add(this.txtTargetBlockName);
            this.Controls.Add(this.txtSourceBlockName);
            this.Controls.Add(this.btnClearDirectory);
            this.Controls.Add(this.btnSelectDirectory);
            this.Controls.Add(this.lblDirectoryInfo);
            this.Name = "BlockReplaceForm";
            this.Text = "块替换工具";
            ((System.ComponentModel.ISupportInitialize)(this.pbPreviewImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbBlockThumbnail)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblDirectoryInfo;
        private System.Windows.Forms.Label lblSourceBlock;
        private System.Windows.Forms.Label lblTargetBlock;
        private System.Windows.Forms.Button btnSelectDirectory;
        private System.Windows.Forms.Button btnClearDirectory;
        private System.Windows.Forms.TextBox txtSourceBlockName;
        private System.Windows.Forms.TextBox txtTargetBlockName;
        private System.Windows.Forms.Button btnBrowseSourceBlock;
        private System.Windows.Forms.Button btnBrowseTargetBlock;
        private System.Windows.Forms.CheckBox chkPreserveAttributes;
        private System.Windows.Forms.CheckBox chkPreservePosition;
        private System.Windows.Forms.CheckBox chkPreserveRotation;
        private System.Windows.Forms.CheckBox chkPreserveScale;
        private System.Windows.Forms.CheckBox chkCreateBackup;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.Button btnStartReplace;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.RichTextBox rtbLog;
        private System.Windows.Forms.Button btnClearLog;
        private System.Windows.Forms.Button btnSaveReport;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.PictureBox pbPreviewImage;
        private System.Windows.Forms.PictureBox pbBlockThumbnail;
        private System.Windows.Forms.Label lblBlockThumbnail;
    }
}