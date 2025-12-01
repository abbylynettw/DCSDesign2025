using System;
using System.Drawing;
using System.Windows.Forms;

namespace InstallerPlugin
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        // 控件变量声明
        private Label titleLabel;
        private GroupBox versionGroup;
        private RadioButton rb2022;
        private RadioButton rb2023;
        private RadioButton rb2024;
        private Button installButton;
        private Button uninstallButton;
        private Button closeButton;

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
            this.titleLabel = new System.Windows.Forms.Label();
            this.versionGroup = new System.Windows.Forms.GroupBox();
            this.rb2022 = new System.Windows.Forms.RadioButton();
            this.rb2023 = new System.Windows.Forms.RadioButton();
            this.rb2024 = new System.Windows.Forms.RadioButton();
            this.installButton = new System.Windows.Forms.Button();
            this.uninstallButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.versionGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Arial", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLabel.Location = new System.Drawing.Point(100, 20);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(252, 26);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "DCSDesign2025 插件安装工具";
            // 
            // versionGroup
            // 
            this.versionGroup.Controls.Add(this.rb2022);
            this.versionGroup.Controls.Add(this.rb2023);
            this.versionGroup.Controls.Add(this.rb2024);
            this.versionGroup.Location = new System.Drawing.Point(40, 60);
            this.versionGroup.Name = "versionGroup";
            this.versionGroup.Size = new System.Drawing.Size(350, 120);
            this.versionGroup.TabIndex = 1;
            this.versionGroup.TabStop = false;
            this.versionGroup.Text = "选择AutoCAD版本";
            // 
            // rb2022
            // 
            this.rb2022.AutoSize = true;
            this.rb2022.Checked = true;
            this.rb2022.Location = new System.Drawing.Point(20, 20);
            this.rb2022.Name = "rb2022";
            this.rb2022.Size = new System.Drawing.Size(95, 17);
            this.rb2022.TabIndex = 0;
            this.rb2022.TabStop = true;
            this.rb2022.Tag = "AutoCAD 2022";
            this.rb2022.Text = "AutoCAD 2022";
            this.rb2022.UseVisualStyleBackColor = true;
            // 
            // rb2023
            // 
            this.rb2023.AutoSize = true;
            this.rb2023.Location = new System.Drawing.Point(20, 50);
            this.rb2023.Name = "rb2023";
            this.rb2023.Size = new System.Drawing.Size(95, 17);
            this.rb2023.TabIndex = 1;
            this.rb2023.Tag = "AutoCAD 2023";
            this.rb2023.Text = "AutoCAD 2023";
            this.rb2023.UseVisualStyleBackColor = true;
            // 
            // rb2024
            // 
            this.rb2024.AutoSize = true;
            this.rb2024.Location = new System.Drawing.Point(20, 80);
            this.rb2024.Name = "rb2024";
            this.rb2024.Size = new System.Drawing.Size(95, 17);
            this.rb2024.TabIndex = 2;
            this.rb2024.Tag = "AutoCAD 2024";
            this.rb2024.Text = "AutoCAD 2024";
            this.rb2024.UseVisualStyleBackColor = true;
            // 
            // installButton
            // 
            this.installButton.Location = new System.Drawing.Point(80, 200);
            this.installButton.Name = "installButton";
            this.installButton.Size = new System.Drawing.Size(120, 40);
            this.installButton.TabIndex = 2;
            this.installButton.Text = "安装插件";
            this.installButton.UseVisualStyleBackColor = true;
            this.installButton.Click += new System.EventHandler(this.InstallButton_Click);
            // 
            // uninstallButton
            // 
            this.uninstallButton.Location = new System.Drawing.Point(230, 200);
            this.uninstallButton.Name = "uninstallButton";
            this.uninstallButton.Size = new System.Drawing.Size(120, 40);
            this.uninstallButton.TabIndex = 3;
            this.uninstallButton.Text = "卸载插件";
            this.uninstallButton.UseVisualStyleBackColor = true;
            this.uninstallButton.Click += new System.EventHandler(this.UninstallButton_Click);
            // 
            // closeButton
            // 
            this.closeButton.Location = new System.Drawing.Point(170, 260);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(90, 30);
            this.closeButton.TabIndex = 4;
            this.closeButton.Text = "关闭";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(450, 350);
            this.Controls.Add(this.titleLabel);
            this.Controls.Add(this.versionGroup);
            this.Controls.Add(this.installButton);
            this.Controls.Add(this.uninstallButton);
            this.Controls.Add(this.closeButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DCSDesign2025 安装工具";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.versionGroup.ResumeLayout(false);
            this.versionGroup.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}