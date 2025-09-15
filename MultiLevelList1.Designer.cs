namespace WordMan_VSTO
{
    partial class MultiLevelList
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
            this.mainPanel = new System.Windows.Forms.Panel();
            this.leftPanel = new System.Windows.Forms.Panel();
            this.btnSetMultiLevelList = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnLoadCurrentList = new System.Windows.Forms.Button();
            this.btnSetLevelStyle = new System.Windows.Forms.Button();
            this.cmbLevelCount = new System.Windows.Forms.ComboBox();
            this.lblLevelCount = new System.Windows.Forms.Label();
            this.levelsScrollPanel = new System.Windows.Forms.Panel();
            this.levelsContainer = new System.Windows.Forms.Panel();
            this.rightPanel = new System.Windows.Forms.Panel();
            this.btnApplySettings = new System.Windows.Forms.Button();
            this.quickSettingsPanel = new System.Windows.Forms.Panel();
            this.linkStyleGroupBox = new System.Windows.Forms.GroupBox();
            this.chkUnlinkTitles = new System.Windows.Forms.CheckBox();
            this.chkLinkTitles = new System.Windows.Forms.CheckBox();
            this.progressiveIndentGroupBox = new System.Windows.Forms.GroupBox();
            this.numericUpDown3 = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.chkProgressiveIndent = new System.Windows.Forms.CheckBox();
            this.quickSettingsGroupBox = new System.Windows.Forms.GroupBox();
            this.numericUpDown5 = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown4 = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.chkTabPosition = new System.Windows.Forms.CheckBox();
            this.chkTextIndent = new System.Windows.Forms.CheckBox();
            this.chkNumberIndent = new System.Windows.Forms.CheckBox();
            this.lblFirstLevelIndent = new System.Windows.Forms.Label();
            this.lblIncrementIndent = new System.Windows.Forms.Label();
            this.lblSection1 = new System.Windows.Forms.Label();
            this.lblSection2 = new System.Windows.Forms.Label();
            this.lblSection3 = new System.Windows.Forms.Label();
            this.lblNumberIndent = new System.Windows.Forms.Label();
            
            // 设计器友好的静态控件初始化
            this.sampleLevelPanel = new System.Windows.Forms.Panel();
            this.lblSampleLevel = new System.Windows.Forms.Label();
            this.cmbSampleNumberStyle = new System.Windows.Forms.ComboBox();
            this.txtSampleNumberFormat = new System.Windows.Forms.TextBox();
            this.nudSampleNumberIndent = new System.Windows.Forms.NumericUpDown();
            this.nudSampleTextIndent = new System.Windows.Forms.NumericUpDown();
            this.cmbSampleAfterNumber = new System.Windows.Forms.ComboBox();
            this.nudSampleTabPosition = new System.Windows.Forms.NumericUpDown();
            this.cmbSampleLinkedStyle = new System.Windows.Forms.ComboBox();
            
            // 自定义控件 - 使用Word API进行单位转换
            this.numericUpDownWithUnit1 = new NumericUpDownWithUnit(Globals.ThisAddIn.Application, "厘米");
            this.numericUpDownWithUnit2 = new NumericUpDownWithUnit(Globals.ThisAddIn.Application, "厘米");
            this.numericUpDownWithUnit3 = new NumericUpDownWithUnit(Globals.ThisAddIn.Application, "厘米");
            this.numericUpDownWithUnit4 = new NumericUpDownWithUnit(Globals.ThisAddIn.Application, "厘米");
            this.numericUpDownWithUnit5 = new NumericUpDownWithUnit(Globals.ThisAddIn.Application, "厘米");
            this.mainPanel.SuspendLayout();
            this.leftPanel.SuspendLayout();
            this.levelsScrollPanel.SuspendLayout();
            this.rightPanel.SuspendLayout();
            this.quickSettingsPanel.SuspendLayout();
            this.linkStyleGroupBox.SuspendLayout();
            this.progressiveIndentGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            this.quickSettingsGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.rightPanel);
            this.mainPanel.Controls.Add(this.leftPanel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(1400, 480);
            this.mainPanel.TabIndex = 0;
            // 
            // leftPanel
            // 
            this.leftPanel.BackColor = System.Drawing.Color.Transparent;
            this.leftPanel.Controls.Add(this.btnSetMultiLevelList);
            this.leftPanel.Controls.Add(this.btnClose);
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
            // 
            // btnSetMultiLevelList
            // 
            this.btnSetMultiLevelList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnSetMultiLevelList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnSetMultiLevelList.FlatAppearance.BorderSize = 1;
            this.btnSetMultiLevelList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetMultiLevelList.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnSetMultiLevelList.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnSetMultiLevelList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSetMultiLevelList.Location = new System.Drawing.Point(540, 415);
            this.btnSetMultiLevelList.Name = "btnSetMultiLevelList";
            this.btnSetMultiLevelList.Size = new System.Drawing.Size(110, 35);
            this.btnSetMultiLevelList.TabIndex = 4;
            this.btnSetMultiLevelList.Text = "应用";
            this.btnSetMultiLevelList.UseVisualStyleBackColor = false;
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnClose.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnClose.FlatAppearance.BorderSize = 1;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnClose.ForeColor = System.Drawing.Color.Black;
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClose.Location = new System.Drawing.Point(660, 415);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(110, 35);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnLoadCurrentList
            // 
            this.btnLoadCurrentList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnLoadCurrentList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnLoadCurrentList.FlatAppearance.BorderSize = 1;
            this.btnLoadCurrentList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLoadCurrentList.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnLoadCurrentList.ForeColor = System.Drawing.Color.Black;
            this.btnLoadCurrentList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnLoadCurrentList.Location = new System.Drawing.Point(420, 415);
            this.btnLoadCurrentList.Name = "btnLoadCurrentList";
            this.btnLoadCurrentList.Size = new System.Drawing.Size(110, 35);
            this.btnLoadCurrentList.TabIndex = 3;
            this.btnLoadCurrentList.Text = "载入当前列表";
            this.btnLoadCurrentList.UseVisualStyleBackColor = false;
            // 
            // btnSetLevelStyle
            // 
            this.btnSetLevelStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnSetLevelStyle.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnSetLevelStyle.FlatAppearance.BorderSize = 1;
            this.btnSetLevelStyle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetLevelStyle.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnSetLevelStyle.ForeColor = System.Drawing.Color.Black;
            this.btnSetLevelStyle.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSetLevelStyle.Location = new System.Drawing.Point(300, 415);
            this.btnSetLevelStyle.Name = "btnSetLevelStyle";
            this.btnSetLevelStyle.Size = new System.Drawing.Size(110, 35);
            this.btnSetLevelStyle.TabIndex = 2;
            this.btnSetLevelStyle.Text = "设置每级样式";
            this.btnSetLevelStyle.UseVisualStyleBackColor = false;
            // 
            // cmbLevelCount
            // 
            this.cmbLevelCount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLevelCount.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Regular);
            this.cmbLevelCount.FormattingEnabled = true;
            this.cmbLevelCount.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9"});
            this.cmbLevelCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmbLevelCount.Location = new System.Drawing.Point(100, 420);
            this.cmbLevelCount.Name = "cmbLevelCount";
            this.cmbLevelCount.Size = new System.Drawing.Size(80, 25);
            this.cmbLevelCount.TabIndex = 1;
            // 
            // lblLevelCount
            // 
            this.lblLevelCount.AutoSize = true;
            this.lblLevelCount.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Regular);
            this.lblLevelCount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblLevelCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLevelCount.Location = new System.Drawing.Point(15, 425);
            this.lblLevelCount.Name = "lblLevelCount";
            this.lblLevelCount.Size = new System.Drawing.Size(79, 20);
            this.lblLevelCount.TabIndex = 0;
            this.lblLevelCount.Text = "设计列表级数：";
            // 
            // levelsScrollPanel
            // 
            this.levelsScrollPanel.AutoScroll = true;
            this.levelsScrollPanel.BackColor = System.Drawing.Color.Transparent;
            this.levelsScrollPanel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.levelsScrollPanel.Controls.Add(this.levelsContainer);
            this.levelsScrollPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
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
            this.levelsContainer.Size = new System.Drawing.Size(558, 0);
            this.levelsContainer.TabIndex = 0;
            // 
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
            // 
            // btnApplySettings
            // 
            this.btnApplySettings.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.btnApplySettings.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnApplySettings.FlatAppearance.BorderSize = 1;
            this.btnApplySettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnApplySettings.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnApplySettings.ForeColor = System.Drawing.Color.Black;
            this.btnApplySettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnApplySettings.Location = new System.Drawing.Point(100, 380);
            this.btnApplySettings.Name = "btnApplySettings";
            this.btnApplySettings.Size = new System.Drawing.Size(130, 35);
            this.btnApplySettings.TabIndex = 0;
            this.btnApplySettings.Text = "应用以上设置";
            this.btnApplySettings.UseVisualStyleBackColor = false;
            // 
            // quickSettingsPanel
            // 
            this.quickSettingsPanel.Controls.Add(this.quickSettingsGroupBox);
            this.quickSettingsPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.quickSettingsPanel.Location = new System.Drawing.Point(15, 15);
            this.quickSettingsPanel.Name = "quickSettingsPanel";
            this.quickSettingsPanel.Size = new System.Drawing.Size(320, 420);
            this.quickSettingsPanel.TabIndex = 0;
            // 
            // linkStyleGroupBox
            // 
            this.linkStyleGroupBox.BackColor = System.Drawing.Color.Transparent;
            this.linkStyleGroupBox.Controls.Add(this.chkUnlinkTitles);
            this.linkStyleGroupBox.Controls.Add(this.chkLinkTitles);
            this.linkStyleGroupBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.linkStyleGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.linkStyleGroupBox.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.linkStyleGroupBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.linkStyleGroupBox.Location = new System.Drawing.Point(0, 300);
            this.linkStyleGroupBox.Name = "linkStyleGroupBox";
            this.linkStyleGroupBox.Padding = new System.Windows.Forms.Padding(15);
            this.linkStyleGroupBox.Size = new System.Drawing.Size(370, 100);
            this.linkStyleGroupBox.TabIndex = 2;
            this.linkStyleGroupBox.TabStop = false;
            this.linkStyleGroupBox.Text = "3. 链接样式设置";
            // 
            // chkUnlinkTitles
            // 
            this.chkUnlinkTitles.AutoSize = true;
            this.chkUnlinkTitles.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.chkUnlinkTitles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkUnlinkTitles.Location = new System.Drawing.Point(40, 310);
            this.chkUnlinkTitles.Name = "chkUnlinkTitles";
            this.chkUnlinkTitles.Size = new System.Drawing.Size(75, 21);
            this.chkUnlinkTitles.TabIndex = 1;
            this.chkUnlinkTitles.Text = "不链接标题样式";
            this.chkUnlinkTitles.UseVisualStyleBackColor = true;
            // 
            // chkLinkTitles
            // 
            this.chkLinkTitles.AutoSize = true;
            this.chkLinkTitles.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.chkLinkTitles.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkLinkTitles.Location = new System.Drawing.Point(40, 280);
            this.chkLinkTitles.Name = "chkLinkTitles";
            this.chkLinkTitles.Size = new System.Drawing.Size(123, 21);
            this.chkLinkTitles.TabIndex = 0;
            this.chkLinkTitles.Text = "链接到标题样式";
            this.chkLinkTitles.UseVisualStyleBackColor = true;
            // 
            // progressiveIndentGroupBox
            // 
            this.progressiveIndentGroupBox.BackColor = System.Drawing.Color.Transparent;
            this.progressiveIndentGroupBox.Controls.Add(this.numericUpDown3);
            this.progressiveIndentGroupBox.Controls.Add(this.numericUpDown2);
            this.progressiveIndentGroupBox.Controls.Add(this.chkProgressiveIndent);
            this.progressiveIndentGroupBox.Controls.Add(this.lblFirstLevelIndent);
            this.progressiveIndentGroupBox.Controls.Add(this.lblIncrementIndent);
            this.progressiveIndentGroupBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.progressiveIndentGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.progressiveIndentGroupBox.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.progressiveIndentGroupBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.progressiveIndentGroupBox.Location = new System.Drawing.Point(0, 100);
            this.progressiveIndentGroupBox.Name = "progressiveIndentGroupBox";
            this.progressiveIndentGroupBox.Padding = new System.Windows.Forms.Padding(15);
            this.progressiveIndentGroupBox.Size = new System.Drawing.Size(370, 100);
            this.progressiveIndentGroupBox.TabIndex = 1;
            this.progressiveIndentGroupBox.TabStop = false;
            this.progressiveIndentGroupBox.Text = "2. 递进缩进设置";
            // 
            // numericUpDown3
            // 
            this.numericUpDown3.DecimalPlaces = 2;
            this.numericUpDown3.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDown3.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDown3.Location = new System.Drawing.Point(200, 210);
            this.numericUpDown3.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDown3.Name = "numericUpDown3";
            this.numericUpDown3.Size = new System.Drawing.Size(100, 25);
            this.numericUpDown3.TabIndex = 2;
            this.numericUpDown3.Enabled = false;
            // 
            // numericUpDown2
            // 
            this.numericUpDown2.DecimalPlaces = 2;
            this.numericUpDown2.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDown2.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDown2.Location = new System.Drawing.Point(200, 180);
            this.numericUpDown2.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(100, 25);
            this.numericUpDown2.TabIndex = 1;
            this.numericUpDown2.Enabled = false;
            // 
            // chkProgressiveIndent
            // 
            this.chkProgressiveIndent.AutoSize = true;
            this.chkProgressiveIndent.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.chkProgressiveIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkProgressiveIndent.Location = new System.Drawing.Point(200, 150);
            this.chkProgressiveIndent.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProgressiveIndent.Name = "chkProgressiveIndent";
            this.chkProgressiveIndent.Size = new System.Drawing.Size(120, 21);
            this.chkProgressiveIndent.TabIndex = 0;
            this.chkProgressiveIndent.Text = " ";
            this.chkProgressiveIndent.UseVisualStyleBackColor = true;
            // 
            // quickSettingsGroupBox
            // 
            this.quickSettingsGroupBox.BackColor = System.Drawing.Color.Transparent;
            this.quickSettingsGroupBox.Controls.Add(this.btnApplySettings);
            this.quickSettingsGroupBox.Controls.Add(this.lblSection3);
            this.quickSettingsGroupBox.Controls.Add(this.lblSection2);
            this.quickSettingsGroupBox.Controls.Add(this.lblSection1);
            this.quickSettingsGroupBox.Controls.Add(this.lblNumberIndent);
            this.quickSettingsGroupBox.Controls.Add(this.numericUpDownWithUnit1);
            this.quickSettingsGroupBox.Controls.Add(this.numericUpDownWithUnit2);
            this.quickSettingsGroupBox.Controls.Add(this.numericUpDownWithUnit3);
            this.quickSettingsGroupBox.Controls.Add(this.numericUpDownWithUnit4);
            this.quickSettingsGroupBox.Controls.Add(this.numericUpDownWithUnit5);
            this.quickSettingsGroupBox.Controls.Add(this.chkUnlinkTitles);
            this.quickSettingsGroupBox.Controls.Add(this.chkLinkTitles);
            this.quickSettingsGroupBox.Controls.Add(this.chkProgressiveIndent);
            this.quickSettingsGroupBox.Controls.Add(this.chkTabPosition);
            this.quickSettingsGroupBox.Controls.Add(this.chkTextIndent);
            this.quickSettingsGroupBox.Controls.Add(this.chkNumberIndent);
            this.quickSettingsGroupBox.Controls.Add(this.lblFirstLevelIndent);
            this.quickSettingsGroupBox.Controls.Add(this.lblIncrementIndent);
            this.quickSettingsGroupBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.quickSettingsGroupBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.quickSettingsGroupBox.Font = new System.Drawing.Font("Microsoft YaHei", 10F, System.Drawing.FontStyle.Bold);
            this.quickSettingsGroupBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.quickSettingsGroupBox.Location = new System.Drawing.Point(0, 0);
            this.quickSettingsGroupBox.Name = "quickSettingsGroupBox";
            this.quickSettingsGroupBox.Padding = new System.Windows.Forms.Padding(15);
            this.quickSettingsGroupBox.Size = new System.Drawing.Size(320, 420);
            this.quickSettingsGroupBox.TabIndex = 0;
            this.quickSettingsGroupBox.TabStop = false;
            this.quickSettingsGroupBox.Text = "快速样式设置";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.DecimalPlaces = 2;
            this.numericUpDown1.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDown1.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDown1.Location = new System.Drawing.Point(200, 50);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(100, 25);
            this.numericUpDown1.TabIndex = 3;
            this.numericUpDown1.Enabled = false;
            // 
            // chkTabPosition
            // 
            this.chkTabPosition.AutoSize = true;
            this.chkTabPosition.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
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
            this.chkTextIndent.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
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
            this.chkNumberIndent.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.chkNumberIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.chkNumberIndent.Location = new System.Drawing.Point(40, 50);
            this.chkNumberIndent.Name = "chkNumberIndent";
            this.chkNumberIndent.Size = new System.Drawing.Size(75, 21);
            this.chkNumberIndent.TabIndex = 0;
            this.chkNumberIndent.Text = "编号缩进";
            this.chkNumberIndent.UseVisualStyleBackColor = true;
            // 
            // lblFirstLevelIndent
            // 
            this.lblFirstLevelIndent.AutoSize = false;
            this.lblFirstLevelIndent.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.lblFirstLevelIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblFirstLevelIndent.Location = new System.Drawing.Point(20, 180);
            this.lblFirstLevelIndent.Name = "lblFirstLevelIndent";
            this.lblFirstLevelIndent.Size = new System.Drawing.Size(120, 17);
            this.lblFirstLevelIndent.TabIndex = 4;
            this.lblFirstLevelIndent.Text = "各级编号递缩：";
            this.lblFirstLevelIndent.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblIncrementIndent
            // 
            this.lblIncrementIndent.AutoSize = true;
            this.lblIncrementIndent.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.lblIncrementIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblIncrementIndent.Location = new System.Drawing.Point(20, 210);
            this.lblIncrementIndent.Name = "lblIncrementIndent";
            this.lblIncrementIndent.Size = new System.Drawing.Size(120, 17);
            this.lblIncrementIndent.TabIndex = 5;
            this.lblIncrementIndent.Text = "递进缩进增量：";
            // 
            // numericUpDown4
            // 
            this.numericUpDown4.DecimalPlaces = 2;
            this.numericUpDown4.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDown4.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDown4.Location = new System.Drawing.Point(200, 80);
            this.numericUpDown4.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDown4.Name = "numericUpDown4";
            this.numericUpDown4.Size = new System.Drawing.Size(100, 25);
            this.numericUpDown4.TabIndex = 4;
            this.numericUpDown4.Enabled = false;
            // 
            // lblSection1
            // 
            this.lblSection1.AutoSize = true;
            this.lblSection1.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold);
            this.lblSection1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.lblSection1.Location = new System.Drawing.Point(20, 20);
            this.lblSection1.Name = "lblSection1";
            this.lblSection1.Size = new System.Drawing.Size(107, 17);
            this.lblSection1.TabIndex = 0;
            this.lblSection1.Text = "1. 统一缩进设置";
            // 
            // lblSection2
            // 
            this.lblSection2.AutoSize = true;
            this.lblSection2.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold);
            this.lblSection2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.lblSection2.Location = new System.Drawing.Point(20, 150);
            this.lblSection2.Name = "lblSection2";
            this.lblSection2.Size = new System.Drawing.Size(107, 17);
            this.lblSection2.TabIndex = 0;
            this.lblSection2.Text = "2. 递进缩进设置";
            // 
            // lblSection3
            // 
            this.lblSection3.AutoSize = true;
            this.lblSection3.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Bold);
            this.lblSection3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.lblSection3.Location = new System.Drawing.Point(20, 250);
            this.lblSection3.Name = "lblSection3";
            this.lblSection3.Size = new System.Drawing.Size(107, 17);
            this.lblSection3.TabIndex = 0;
            this.lblSection3.Text = "3. 链接样式设置";
            // 
            // lblNumberIndent
            // 
            this.lblNumberIndent.AutoSize = false;
            this.lblNumberIndent.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.lblNumberIndent.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.lblNumberIndent.Location = new System.Drawing.Point(20, 180);
            this.lblNumberIndent.Name = "lblNumberIndent";
            this.lblNumberIndent.Size = new System.Drawing.Size(120, 17);
            this.lblNumberIndent.TabIndex = 0;
            this.lblNumberIndent.Text = "一级编号缩进：";
            this.lblNumberIndent.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // numericUpDown5
            // 
            this.numericUpDown5.DecimalPlaces = 2;
            this.numericUpDown5.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDown5.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDown5.Location = new System.Drawing.Point(200, 110);
            this.numericUpDown5.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDown5.Name = "numericUpDown5";
            this.numericUpDown5.Size = new System.Drawing.Size(100, 25);
            this.numericUpDown5.TabIndex = 5;
            this.numericUpDown5.Enabled = false;
            // 
            // numericUpDownWithUnit1
            // 
            this.numericUpDownWithUnit1.DecimalPlaces = 2;
            this.numericUpDownWithUnit1.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDownWithUnit1.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDownWithUnit1.Location = new System.Drawing.Point(200, 50);
            this.numericUpDownWithUnit1.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDownWithUnit1.Name = "numericUpDownWithUnit1";
            this.numericUpDownWithUnit1.Size = new System.Drawing.Size(100, 25);
            this.numericUpDownWithUnit1.TabIndex = 6;
            this.numericUpDownWithUnit1.Enabled = false;
            // 
            // numericUpDownWithUnit2
            // 
            this.numericUpDownWithUnit2.DecimalPlaces = 2;
            this.numericUpDownWithUnit2.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDownWithUnit2.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDownWithUnit2.Location = new System.Drawing.Point(200, 180);
            this.numericUpDownWithUnit2.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDownWithUnit2.Name = "numericUpDownWithUnit2";
            this.numericUpDownWithUnit2.Size = new System.Drawing.Size(100, 25);
            this.numericUpDownWithUnit2.TabIndex = 7;
            this.numericUpDownWithUnit2.Enabled = false;
            // 
            // numericUpDownWithUnit3
            // 
            this.numericUpDownWithUnit3.DecimalPlaces = 2;
            this.numericUpDownWithUnit3.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDownWithUnit3.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDownWithUnit3.Location = new System.Drawing.Point(200, 210);
            this.numericUpDownWithUnit3.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDownWithUnit3.Name = "numericUpDownWithUnit3";
            this.numericUpDownWithUnit3.Size = new System.Drawing.Size(100, 25);
            this.numericUpDownWithUnit3.TabIndex = 8;
            this.numericUpDownWithUnit3.Enabled = false;
            // 
            // numericUpDownWithUnit4
            // 
            this.numericUpDownWithUnit4.DecimalPlaces = 2;
            this.numericUpDownWithUnit4.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDownWithUnit4.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDownWithUnit4.Location = new System.Drawing.Point(200, 80);
            this.numericUpDownWithUnit4.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDownWithUnit4.Name = "numericUpDownWithUnit4";
            this.numericUpDownWithUnit4.Size = new System.Drawing.Size(100, 25);
            this.numericUpDownWithUnit4.TabIndex = 9;
            this.numericUpDownWithUnit4.Enabled = false;
            // 
            // numericUpDownWithUnit5
            // 
            this.numericUpDownWithUnit5.DecimalPlaces = 2;
            this.numericUpDownWithUnit5.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.numericUpDownWithUnit5.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.numericUpDownWithUnit5.Location = new System.Drawing.Point(200, 110);
            this.numericUpDownWithUnit5.Maximum = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this.numericUpDownWithUnit5.Name = "numericUpDownWithUnit5";
            this.numericUpDownWithUnit5.Size = new System.Drawing.Size(100, 25);
            this.numericUpDownWithUnit5.TabIndex = 10;
            this.numericUpDownWithUnit5.Enabled = false;
            // 
            // MultiLevelList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(248)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(1200, 480);
            this.Controls.Add(this.mainPanel);
            this.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MultiLevelList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "多级列表设置";
            this.mainPanel.ResumeLayout(false);
            this.leftPanel.ResumeLayout(false);
            this.leftPanel.PerformLayout();
            this.levelsScrollPanel.ResumeLayout(false);
            this.levelsScrollPanel.PerformLayout();
            this.rightPanel.ResumeLayout(false);
            this.quickSettingsPanel.ResumeLayout(false);
            this.linkStyleGroupBox.ResumeLayout(false);
            this.linkStyleGroupBox.PerformLayout();
            this.progressiveIndentGroupBox.ResumeLayout(false);
            this.progressiveIndentGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            this.quickSettingsGroupBox.ResumeLayout(false);
            this.quickSettingsGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Panel leftPanel;
        private System.Windows.Forms.Button btnSetMultiLevelList;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnLoadCurrentList;
        private System.Windows.Forms.Button btnSetLevelStyle;
        private System.Windows.Forms.ComboBox cmbLevelCount;
        private System.Windows.Forms.Label lblLevelCount;
        private System.Windows.Forms.Panel levelsScrollPanel;
        private System.Windows.Forms.Panel levelsContainer;
        private System.Windows.Forms.Panel rightPanel;
        private System.Windows.Forms.Button btnApplySettings;
        private System.Windows.Forms.Panel quickSettingsPanel;
        private System.Windows.Forms.GroupBox linkStyleGroupBox;
        private System.Windows.Forms.CheckBox chkUnlinkTitles;
        private System.Windows.Forms.CheckBox chkLinkTitles;
        private System.Windows.Forms.GroupBox progressiveIndentGroupBox;
        private System.Windows.Forms.NumericUpDown numericUpDown3;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private System.Windows.Forms.CheckBox chkProgressiveIndent;
        private System.Windows.Forms.GroupBox quickSettingsGroupBox;
        private System.Windows.Forms.NumericUpDown numericUpDown5;
        private System.Windows.Forms.NumericUpDown numericUpDown4;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.CheckBox chkTabPosition;
        private System.Windows.Forms.CheckBox chkTextIndent;
        private System.Windows.Forms.CheckBox chkNumberIndent;
        private System.Windows.Forms.Label lblFirstLevelIndent;
        private System.Windows.Forms.Label lblIncrementIndent;
        private System.Windows.Forms.Label lblSection1;
        private System.Windows.Forms.Label lblSection2;
        private System.Windows.Forms.Label lblSection3;
        private System.Windows.Forms.Label lblNumberIndent;
        
        // 设计器友好的静态控件 - 用于示例级别
        private System.Windows.Forms.Panel sampleLevelPanel;
        private System.Windows.Forms.Label lblSampleLevel;
        private System.Windows.Forms.ComboBox cmbSampleNumberStyle;
        private System.Windows.Forms.TextBox txtSampleNumberFormat;
        private System.Windows.Forms.NumericUpDown nudSampleNumberIndent;
        private System.Windows.Forms.NumericUpDown nudSampleTextIndent;
        private System.Windows.Forms.ComboBox cmbSampleAfterNumber;
        private System.Windows.Forms.NumericUpDown nudSampleTabPosition;
        private System.Windows.Forms.ComboBox cmbSampleLinkedStyle;
        
        // 自定义控件
        private NumericUpDownWithUnit numericUpDownWithUnit1;
        private NumericUpDownWithUnit numericUpDownWithUnit2;
        private NumericUpDownWithUnit numericUpDownWithUnit3;
        private NumericUpDownWithUnit numericUpDownWithUnit4;
        private NumericUpDownWithUnit numericUpDownWithUnit5;
    }
}