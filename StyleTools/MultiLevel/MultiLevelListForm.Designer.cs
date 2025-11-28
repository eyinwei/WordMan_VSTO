using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WordMan;

namespace WordMan
{
    partial class MultiLevelListForm
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
            // 面板控件
            this.mainPanel = new System.Windows.Forms.Panel();
            this.leftPanel = new System.Windows.Forms.Panel();
            this.rightPanel = new System.Windows.Forms.Panel();
            this.quickSettingsPanel = new System.Windows.Forms.Panel();
            this.levelsScrollPanel = new System.Windows.Forms.Panel();
            this.levelsContainer = new System.Windows.Forms.Panel();
            
            // 组框控件
            this.quickSettingsGroupBox = new System.Windows.Forms.GroupBox();
            
            // 按钮控件
            this.btnApplySettings = new WordMan.StandardButton();
            this.btnSetMultiLevelList = new WordMan.StandardButton();
            this.btnClose = new WordMan.StandardButton();
            this.btnImport = new WordMan.StandardButton();
            this.btnExport = new WordMan.StandardButton();
            this.btnLoadCurrentList = new WordMan.StandardButton();
            this.btnSetLevelStyle = new WordMan.StandardButton();
            
            // 标签控件
            this.lblLinkStyleSection = new System.Windows.Forms.Label();
            this.lblProgressiveIndentSection = new System.Windows.Forms.Label();
            this.lblUnifiedIndentSection = new System.Windows.Forms.Label();
            this.lblNumberIndent = new System.Windows.Forms.Label();
            this.lblIncrementIndent = new System.Windows.Forms.Label();
            this.lblLevelCount = new System.Windows.Forms.Label();
            
            // 复选框控件
            this.chkTabPosition = new System.Windows.Forms.CheckBox();
            this.chkTextIndent = new System.Windows.Forms.CheckBox();
            this.chkNumberIndent = new System.Windows.Forms.CheckBox();
            this.chkLinkTitles = new System.Windows.Forms.CheckBox();
            this.chkUnlinkTitles = new System.Windows.Forms.CheckBox();
            this.chkProgressiveIndent = new System.Windows.Forms.CheckBox();
            
            // 下拉框控件
            this.cmbLevelCount = new WordMan.StandardComboBox();
            
            // 数值输入框控件
            this.nudNumberIndent = new WordMan.StandardNumericUpDown(null, "厘米");
            this.nudFirstLevelIndent = new WordMan.StandardNumericUpDown(null, "厘米");
            this.nudIncrementIndent = new WordMan.StandardNumericUpDown(null, "厘米");
            this.nudTextIndent = new WordMan.StandardNumericUpDown(null, "厘米");
            this.nudTabPosition = new WordMan.StandardNumericUpDown(null, "厘米");
            this.mainPanel.SuspendLayout();
            this.rightPanel.SuspendLayout();
            this.quickSettingsPanel.SuspendLayout();
            this.quickSettingsGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudNumberIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstLevelIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudIncrementIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudTextIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudTabPosition)).BeginInit();
            this.leftPanel.SuspendLayout();
            this.levelsScrollPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.rightPanel);
            this.mainPanel.Controls.Add(this.leftPanel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(1200, 480);
            this.mainPanel.TabIndex = 0;
            // 
            // rightPanel
            // 
            this.rightPanel.BackColor = System.Drawing.Color.Transparent;
            this.rightPanel.Controls.Add(this.quickSettingsPanel);
            this.rightPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.rightPanel.Location = new System.Drawing.Point(850, 0);
            this.rightPanel.Name = "rightPanel";
            this.rightPanel.Padding = new System.Windows.Forms.Padding(15);
            this.rightPanel.Size = new System.Drawing.Size(350, 480);
            this.rightPanel.TabIndex = 1;
            // 
            // quickSettingsPanel
            // 
            this.quickSettingsPanel.Controls.Add(this.quickSettingsGroupBox);
            this.quickSettingsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.quickSettingsPanel.Location = new System.Drawing.Point(15, 15);
            this.quickSettingsPanel.Name = "quickSettingsPanel";
            this.quickSettingsPanel.Size = new System.Drawing.Size(320, 450);
            this.quickSettingsPanel.TabIndex = 0;
            // 
            // quickSettingsGroupBox
            // 
            this.quickSettingsGroupBox.BackColor = System.Drawing.Color.Transparent;
            this.quickSettingsGroupBox.Controls.Add(this.btnApplySettings);
            this.quickSettingsGroupBox.Controls.Add(this.lblLinkStyleSection);
            this.quickSettingsGroupBox.Controls.Add(this.lblProgressiveIndentSection);
            this.quickSettingsGroupBox.Controls.Add(this.lblUnifiedIndentSection);
            this.quickSettingsGroupBox.Controls.Add(this.lblNumberIndent);
            this.quickSettingsGroupBox.Controls.Add(this.nudNumberIndent);
            this.quickSettingsGroupBox.Controls.Add(this.nudFirstLevelIndent);
            this.quickSettingsGroupBox.Controls.Add(this.nudIncrementIndent);
            this.quickSettingsGroupBox.Controls.Add(this.nudTextIndent);
            this.quickSettingsGroupBox.Controls.Add(this.nudTabPosition);
            this.quickSettingsGroupBox.Controls.Add(this.chkTabPosition);
            this.quickSettingsGroupBox.Controls.Add(this.chkTextIndent);
            this.quickSettingsGroupBox.Controls.Add(this.chkNumberIndent);
            this.quickSettingsGroupBox.Controls.Add(this.chkLinkTitles);
            this.quickSettingsGroupBox.Controls.Add(this.chkUnlinkTitles);
            this.quickSettingsGroupBox.Controls.Add(this.chkProgressiveIndent);
            this.quickSettingsGroupBox.Controls.Add(this.lblIncrementIndent);
            this.quickSettingsGroupBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.quickSettingsGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.quickSettingsGroupBox.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.quickSettingsGroupBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.quickSettingsGroupBox.Location = new System.Drawing.Point(0, 0);
            this.quickSettingsGroupBox.Name = "quickSettingsGroupBox";
            this.quickSettingsGroupBox.Padding = new System.Windows.Forms.Padding(15);
            this.quickSettingsGroupBox.Size = new System.Drawing.Size(320, 441);
            this.quickSettingsGroupBox.TabIndex = 0;
            this.quickSettingsGroupBox.TabStop = false;
            this.quickSettingsGroupBox.Text = "快速样式设置";
            // 
            // btnApplySettings
            // 
            this.btnApplySettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnApplySettings.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnApplySettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnApplySettings.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnApplySettings.ForeColor = System.Drawing.Color.Black;
            this.btnApplySettings.Location = new System.Drawing.Point(98, 400);
            this.btnApplySettings.Name = "btnApplySettings";
            this.btnApplySettings.Size = new System.Drawing.Size(130, 35);
            this.btnApplySettings.TabIndex = 0;
            this.btnApplySettings.Text = "应用以上设置";
            this.btnApplySettings.UseVisualStyleBackColor = false;
            // 
            // lblLinkStyleSection
            // 
            this.lblLinkStyleSection.AutoSize = true;
            this.lblLinkStyleSection.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.lblLinkStyleSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.lblLinkStyleSection.Location = new System.Drawing.Point(20, 250);
            this.lblLinkStyleSection.Name = "lblLinkStyleSection";
            this.lblLinkStyleSection.Size = new System.Drawing.Size(94, 17);
            this.lblLinkStyleSection.TabIndex = 0;
            this.lblLinkStyleSection.Text = "3. 链接样式设置";
            // 
            // lblProgressiveIndentSection
            // 
            this.lblProgressiveIndentSection.AutoSize = true;
            this.lblProgressiveIndentSection.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.lblProgressiveIndentSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.lblProgressiveIndentSection.Location = new System.Drawing.Point(20, 150);
            this.lblProgressiveIndentSection.Name = "lblProgressiveIndentSection";
            this.lblProgressiveIndentSection.Size = new System.Drawing.Size(94, 17);
            this.lblProgressiveIndentSection.TabIndex = 0;
            this.lblProgressiveIndentSection.Text = "2. 递进缩进设置";
            // 
            // lblUnifiedIndentSection
            // 
            this.lblUnifiedIndentSection.AutoSize = true;
            this.lblUnifiedIndentSection.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.lblUnifiedIndentSection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.lblUnifiedIndentSection.Location = new System.Drawing.Point(20, 20);
            this.lblUnifiedIndentSection.Name = "lblUnifiedIndentSection";
            this.lblUnifiedIndentSection.Size = new System.Drawing.Size(94, 17);
            this.lblUnifiedIndentSection.TabIndex = 0;
            this.lblUnifiedIndentSection.Text = "1. 统一缩进设置";
            // 
            // lblNumberIndent
            // 
            this.lblNumberIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblNumberIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblNumberIndent.Location = new System.Drawing.Point(40, 186);
            this.lblNumberIndent.Name = "lblNumberIndent";
            this.lblNumberIndent.Size = new System.Drawing.Size(120, 17);
            this.lblNumberIndent.TabIndex = 0;
            this.lblNumberIndent.Text = "一级编号缩进：";
            this.lblNumberIndent.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // nudNumberIndent
            // 
            this.nudNumberIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.nudNumberIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nudNumberIndent.DecimalPlaces = 1;
            this.nudNumberIndent.Enabled = false;
            this.nudNumberIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudNumberIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudNumberIndent.Location = new System.Drawing.Point(200, 50);
            this.nudNumberIndent.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudNumberIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.nudNumberIndent.Name = "nudNumberIndent";
            this.nudNumberIndent.Size = new System.Drawing.Size(100, 23);
            this.nudNumberIndent.TabIndex = 6;
            this.nudNumberIndent.Unit = "厘米";
            // 
            // nudFirstLevelIndent
            // 
            this.nudFirstLevelIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.nudFirstLevelIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nudFirstLevelIndent.DecimalPlaces = 1;
            this.nudFirstLevelIndent.Enabled = false;
            this.nudFirstLevelIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudFirstLevelIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudFirstLevelIndent.Location = new System.Drawing.Point(200, 180);
            this.nudFirstLevelIndent.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudFirstLevelIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.nudFirstLevelIndent.Name = "nudFirstLevelIndent";
            this.nudFirstLevelIndent.Size = new System.Drawing.Size(100, 23);
            this.nudFirstLevelIndent.TabIndex = 7;
            this.nudFirstLevelIndent.Unit = "厘米";
            // 
            // nudIncrementIndent
            // 
            this.nudIncrementIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.nudIncrementIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nudIncrementIndent.DecimalPlaces = 1;
            this.nudIncrementIndent.Enabled = false;
            this.nudIncrementIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudIncrementIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudIncrementIndent.Location = new System.Drawing.Point(200, 210);
            this.nudIncrementIndent.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudIncrementIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.nudIncrementIndent.Name = "nudIncrementIndent";
            this.nudIncrementIndent.Size = new System.Drawing.Size(100, 23);
            this.nudIncrementIndent.TabIndex = 8;
            this.nudIncrementIndent.Unit = "厘米";
            // 
            // nudTextIndent
            // 
            this.nudTextIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.nudTextIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nudTextIndent.DecimalPlaces = 1;
            this.nudTextIndent.Enabled = false;
            this.nudTextIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudTextIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudTextIndent.Location = new System.Drawing.Point(200, 80);
            this.nudTextIndent.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudTextIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.nudTextIndent.Name = "nudTextIndent";
            this.nudTextIndent.Size = new System.Drawing.Size(100, 23);
            this.nudTextIndent.TabIndex = 9;
            this.nudTextIndent.Unit = "厘米";
            // 
            // nudTabPosition
            // 
            this.nudTabPosition.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.nudTabPosition.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.nudTabPosition.DecimalPlaces = 1;
            this.nudTabPosition.Enabled = false;
            this.nudTabPosition.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudTabPosition.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.nudTabPosition.Location = new System.Drawing.Point(200, 110);
            this.nudTabPosition.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.nudTabPosition.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.nudTabPosition.Name = "nudTabPosition";
            this.nudTabPosition.Size = new System.Drawing.Size(100, 23);
            this.nudTabPosition.TabIndex = 10;
            this.nudTabPosition.Unit = "厘米";
            // 
            // chkTabPosition
            // 
            this.chkTabPosition.AutoSize = true;
            this.chkTabPosition.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkTabPosition.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkTabPosition.Location = new System.Drawing.Point(40, 110);
            this.chkTabPosition.Name = "chkTabPosition";
            this.chkTabPosition.Size = new System.Drawing.Size(87, 21);
            this.chkTabPosition.TabIndex = 2;
            this.chkTabPosition.Text = "制表位位置";
            this.chkTabPosition.UseVisualStyleBackColor = true;
            // 
            // chkTextIndent
            // 
            this.chkTextIndent.AutoSize = true;
            this.chkTextIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkTextIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkTextIndent.Location = new System.Drawing.Point(40, 80);
            this.chkTextIndent.Name = "chkTextIndent";
            this.chkTextIndent.Size = new System.Drawing.Size(75, 21);
            this.chkTextIndent.TabIndex = 1;
            this.chkTextIndent.Text = "文本缩进";
            this.chkTextIndent.UseVisualStyleBackColor = true;
            // 
            // chkNumberIndent
            // 
            this.chkNumberIndent.AutoSize = true;
            this.chkNumberIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkNumberIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkNumberIndent.Location = new System.Drawing.Point(40, 50);
            this.chkNumberIndent.Name = "chkNumberIndent";
            this.chkNumberIndent.Size = new System.Drawing.Size(75, 21);
            this.chkNumberIndent.TabIndex = 0;
            this.chkNumberIndent.Text = "编号缩进";
            this.chkNumberIndent.UseVisualStyleBackColor = true;
            // 
            // chkLinkTitles
            // 
            this.chkLinkTitles.AutoSize = true;
            this.chkLinkTitles.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkLinkTitles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkLinkTitles.Location = new System.Drawing.Point(40, 284);
            this.chkLinkTitles.Name = "chkLinkTitles";
            this.chkLinkTitles.Size = new System.Drawing.Size(111, 21);
            this.chkLinkTitles.TabIndex = 0;
            this.chkLinkTitles.Text = "链接到标题样式";
            this.chkLinkTitles.UseVisualStyleBackColor = true;
            // 
            // chkUnlinkTitles
            // 
            this.chkUnlinkTitles.AutoSize = true;
            this.chkUnlinkTitles.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkUnlinkTitles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkUnlinkTitles.Location = new System.Drawing.Point(40, 313);
            this.chkUnlinkTitles.Name = "chkUnlinkTitles";
            this.chkUnlinkTitles.Size = new System.Drawing.Size(111, 21);
            this.chkUnlinkTitles.TabIndex = 1;
            this.chkUnlinkTitles.Text = "不链接标题样式";
            this.chkUnlinkTitles.UseVisualStyleBackColor = true;
            // 
            // chkProgressiveIndent
            // 
            this.chkProgressiveIndent.AutoSize = true;
            this.chkProgressiveIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkProgressiveIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkProgressiveIndent.Location = new System.Drawing.Point(118, 149);
            this.chkProgressiveIndent.Name = "chkProgressiveIndent";
            this.chkProgressiveIndent.Size = new System.Drawing.Size(15, 14);
            this.chkProgressiveIndent.TabIndex = 0;
            this.chkProgressiveIndent.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProgressiveIndent.UseVisualStyleBackColor = true;
            // 
            // lblIncrementIndent
            // 
            this.lblIncrementIndent.AutoSize = true;
            this.lblIncrementIndent.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblIncrementIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblIncrementIndent.Location = new System.Drawing.Point(40, 212);
            this.lblIncrementIndent.Name = "lblIncrementIndent";
            this.lblIncrementIndent.Size = new System.Drawing.Size(92, 17);
            this.lblIncrementIndent.TabIndex = 5;
            this.lblIncrementIndent.Text = "递进缩进增量：";
            // 
            // leftPanel
            // 
            this.leftPanel.BackColor = System.Drawing.Color.Transparent;
            this.leftPanel.Controls.Add(this.btnSetMultiLevelList);
            this.leftPanel.Controls.Add(this.btnClose);
            this.leftPanel.Controls.Add(this.btnImport);
            this.leftPanel.Controls.Add(this.btnExport);
            this.leftPanel.Controls.Add(this.btnLoadCurrentList);
            this.leftPanel.Controls.Add(this.btnSetLevelStyle);
            this.leftPanel.Controls.Add(this.cmbLevelCount);
            this.leftPanel.Controls.Add(this.lblLevelCount);
            this.leftPanel.Controls.Add(this.levelsScrollPanel);
            this.leftPanel.Dock = System.Windows.Forms.DockStyle.Left;
            this.leftPanel.Location = new System.Drawing.Point(0, 0);
            this.leftPanel.Name = "leftPanel";
            this.leftPanel.Padding = new System.Windows.Forms.Padding(15);
            this.leftPanel.Size = new System.Drawing.Size(1050, 480);
            this.leftPanel.TabIndex = 0;
            // 
            // btnSetMultiLevelList
            // 
            this.btnSetMultiLevelList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSetMultiLevelList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnSetMultiLevelList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetMultiLevelList.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnSetMultiLevelList.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnSetMultiLevelList.Location = new System.Drawing.Point(600, 415);
            this.btnSetMultiLevelList.Name = "btnSetMultiLevelList";
            this.btnSetMultiLevelList.Size = new System.Drawing.Size(110, 35);
            this.btnSetMultiLevelList.TabIndex = 4;
            this.btnSetMultiLevelList.Text = "应用";
            this.btnSetMultiLevelList.UseVisualStyleBackColor = false;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnClose.ForeColor = System.Drawing.Color.Black;
            this.btnClose.Location = new System.Drawing.Point(730, 415);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(110, 35);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnImport
            // 
            this.btnImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnImport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImport.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnImport.ForeColor = System.Drawing.Color.Black;
            this.btnImport.Location = new System.Drawing.Point(200, 415);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(50, 35);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "导入";
            this.btnImport.UseVisualStyleBackColor = false;
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExport.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnExport.ForeColor = System.Drawing.Color.Black;
            this.btnExport.Location = new System.Drawing.Point(255, 415);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(50, 35);
            this.btnExport.TabIndex = 7;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = false;
            // 
            // btnLoadCurrentList
            // 
            this.btnLoadCurrentList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnLoadCurrentList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnLoadCurrentList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLoadCurrentList.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnLoadCurrentList.ForeColor = System.Drawing.Color.Black;
            this.btnLoadCurrentList.Location = new System.Drawing.Point(470, 415);
            this.btnLoadCurrentList.Name = "btnLoadCurrentList";
            this.btnLoadCurrentList.Size = new System.Drawing.Size(110, 35);
            this.btnLoadCurrentList.TabIndex = 3;
            this.btnLoadCurrentList.Text = "载入当前列表";
            this.btnLoadCurrentList.UseVisualStyleBackColor = false;
            // 
            // btnSetLevelStyle
            // 
            this.btnSetLevelStyle.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSetLevelStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnSetLevelStyle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetLevelStyle.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnSetLevelStyle.ForeColor = System.Drawing.Color.Black;
            this.btnSetLevelStyle.Location = new System.Drawing.Point(340, 415);
            this.btnSetLevelStyle.Name = "btnSetLevelStyle";
            this.btnSetLevelStyle.Size = new System.Drawing.Size(110, 35);
            this.btnSetLevelStyle.TabIndex = 2;
            this.btnSetLevelStyle.Text = "设置每级样式";
            this.btnSetLevelStyle.UseVisualStyleBackColor = false;
            this.btnSetLevelStyle.Click += new System.EventHandler(this.BtnSetLevelStyle_Click);
            // 
            // cmbLevelCount
            // 
            this.cmbLevelCount.AllowInput = false;
            this.cmbLevelCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbLevelCount.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.cmbLevelCount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLevelCount.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmbLevelCount.Items.AddRange(new object[] {
            "0",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9"});
            this.cmbLevelCount.Location = new System.Drawing.Point(105, 420);
            this.cmbLevelCount.Name = "cmbLevelCount";
            this.cmbLevelCount.Size = new System.Drawing.Size(80, 25);
            this.cmbLevelCount.TabIndex = 1;
            // 
            // lblLevelCount
            // 
            this.lblLevelCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLevelCount.AutoSize = true;
            this.lblLevelCount.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.lblLevelCount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblLevelCount.Location = new System.Drawing.Point(15, 425);
            this.lblLevelCount.Name = "lblLevelCount";
            this.lblLevelCount.Size = new System.Drawing.Size(107, 20);
            this.lblLevelCount.TabIndex = 0;
            this.lblLevelCount.Text = "设计列表级数：";
            // 
            // levelsScrollPanel
            // 
            this.levelsScrollPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.levelsScrollPanel.AutoScroll = true;
            this.levelsScrollPanel.BackColor = System.Drawing.Color.Transparent;
            this.levelsScrollPanel.Controls.Add(this.levelsContainer);
            this.levelsScrollPanel.Location = new System.Drawing.Point(15, 15);
            this.levelsScrollPanel.Name = "levelsScrollPanel";
            this.levelsScrollPanel.Padding = new System.Windows.Forms.Padding(5);
            this.levelsScrollPanel.Size = new System.Drawing.Size(970, 360);
            this.levelsScrollPanel.TabIndex = 1;
            // 
            // levelsContainer
            // 
            this.levelsContainer.AutoSize = true;
            this.levelsContainer.Dock = System.Windows.Forms.DockStyle.Top;
            this.levelsContainer.Location = new System.Drawing.Point(5, 5);
            this.levelsContainer.Name = "levelsContainer";
            this.levelsContainer.Size = new System.Drawing.Size(960, 0);
            this.levelsContainer.TabIndex = 0;
            // 
            // MultiLevelListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(219)))), ((int)(((byte)(233)))), ((int)(((byte)(247)))));
            this.ClientSize = new System.Drawing.Size(1200, 480);
            this.Controls.Add(this.mainPanel);
            this.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MultiLevelListForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "多级列表设置";
            this.mainPanel.ResumeLayout(false);
            this.rightPanel.ResumeLayout(false);
            this.quickSettingsPanel.ResumeLayout(false);
            this.quickSettingsGroupBox.ResumeLayout(false);
            this.quickSettingsGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudNumberIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstLevelIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudIncrementIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudTextIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudTabPosition)).EndInit();
            this.leftPanel.ResumeLayout(false);
            this.leftPanel.PerformLayout();
            this.levelsScrollPanel.ResumeLayout(false);
            this.levelsScrollPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Panel leftPanel;
        private StandardButton btnSetMultiLevelList;
        private StandardButton btnClose;
        private StandardButton btnLoadCurrentList;
        private StandardButton btnSetLevelStyle;
        private StandardComboBox cmbLevelCount;
        private System.Windows.Forms.Label lblLevelCount;
        private System.Windows.Forms.Panel levelsScrollPanel;
        private System.Windows.Forms.Panel levelsContainer;
        private System.Windows.Forms.Panel rightPanel;
        private StandardButton btnApplySettings;
        private System.Windows.Forms.Panel quickSettingsPanel;
        private System.Windows.Forms.CheckBox chkUnlinkTitles;
        private System.Windows.Forms.CheckBox chkLinkTitles;
        private System.Windows.Forms.CheckBox chkProgressiveIndent;
        private System.Windows.Forms.GroupBox quickSettingsGroupBox;
        private System.Windows.Forms.CheckBox chkTabPosition;
        private System.Windows.Forms.CheckBox chkTextIndent;
        private System.Windows.Forms.CheckBox chkNumberIndent;
        private System.Windows.Forms.Label lblIncrementIndent;
        private System.Windows.Forms.Label lblUnifiedIndentSection;
        private System.Windows.Forms.Label lblProgressiveIndentSection;
        private System.Windows.Forms.Label lblLinkStyleSection;
        private StandardButton btnImport;
        private StandardButton btnExport;


        // 自定义控件
        private StandardNumericUpDown nudNumberIndent;
        private StandardNumericUpDown nudFirstLevelIndent;
        private StandardNumericUpDown nudIncrementIndent;
        private StandardNumericUpDown nudTextIndent;
        private StandardNumericUpDown nudTabPosition;
        private Label lblNumberIndent;
    }
}
