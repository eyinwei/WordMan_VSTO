namespace WordMan_VSTO.MultiLevel
{
    partial class LevelStyleSettingsForm
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
            this.Dta_StyleList = new System.Windows.Forms.DataGridView();
            this.Grp_SetSelectedStyle = new System.Windows.Forms.GroupBox();
            this.Txt_RightIndent = new WordMan_VSTO.StandardNumericUpDown();
            this.Cmb_SpaceAfter = new WordMan_VSTO.StandardComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Btn_FontColor = new System.Windows.Forms.Button();
            this.Txt_LeftIndent = new WordMan_VSTO.StandardNumericUpDown();
            this.Btn_BreakBefore = new WordMan_VSTO.ToggleButton();
            this.Cmb_Alignment = new WordMan_VSTO.StandardComboBox();
            this.Cmb_SpaceBefore = new WordMan_VSTO.StandardComboBox();
            this.Cmb_FontSize = new WordMan_VSTO.StandardComboBox();
            this.Cmb_LineSpacing = new WordMan_VSTO.StandardComboBox();
            this.Cmb_EngFontName = new WordMan_VSTO.StandardComboBox();
            this.Cmb_ChnFontName = new WordMan_VSTO.StandardComboBox();
            this.Btn_Underline = new WordMan_VSTO.ToggleButton();
            this.Btn_Italic = new WordMan_VSTO.ToggleButton();
            this.Btn_Bold = new WordMan_VSTO.ToggleButton();
            this.Btn_SetStyles = new System.Windows.Forms.Button();
            this.Btn_Cancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.Dta_StyleList)).BeginInit();
            this.Grp_SetSelectedStyle.SuspendLayout();
            this.SuspendLayout();
            // 
            // Dta_StyleList
            // 
            this.Dta_StyleList.BackgroundColor = System.Drawing.SystemColors.Window;
            this.Dta_StyleList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Dta_StyleList.Dock = System.Windows.Forms.DockStyle.Top;
            this.Dta_StyleList.Location = new System.Drawing.Point(0, 0);
            this.Dta_StyleList.Name = "Dta_StyleList";
            this.Dta_StyleList.RowTemplate.Height = 23;
            this.Dta_StyleList.Size = new System.Drawing.Size(1200, 265);
            this.Dta_StyleList.TabIndex = 0;
            // 
            // Grp_SetSelectedStyle
            // 
            this.Grp_SetSelectedStyle.Controls.Add(this.Txt_RightIndent);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_SpaceAfter);
            this.Grp_SetSelectedStyle.Controls.Add(this.label14);
            this.Grp_SetSelectedStyle.Controls.Add(this.label13);
            this.Grp_SetSelectedStyle.Controls.Add(this.label12);
            this.Grp_SetSelectedStyle.Controls.Add(this.label11);
            this.Grp_SetSelectedStyle.Controls.Add(this.label10);
            this.Grp_SetSelectedStyle.Controls.Add(this.label9);
            this.Grp_SetSelectedStyle.Controls.Add(this.label8);
            this.Grp_SetSelectedStyle.Controls.Add(this.label7);
            this.Grp_SetSelectedStyle.Controls.Add(this.label6);
            this.Grp_SetSelectedStyle.Controls.Add(this.label5);
            this.Grp_SetSelectedStyle.Controls.Add(this.label4);
            this.Grp_SetSelectedStyle.Controls.Add(this.label3);
            this.Grp_SetSelectedStyle.Controls.Add(this.label2);
            this.Grp_SetSelectedStyle.Controls.Add(this.label1);
            this.Grp_SetSelectedStyle.Controls.Add(this.Btn_FontColor);
            this.Grp_SetSelectedStyle.Controls.Add(this.Txt_LeftIndent);
            this.Grp_SetSelectedStyle.Controls.Add(this.Btn_BreakBefore);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_Alignment);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_SpaceBefore);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_FontSize);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_LineSpacing);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_EngFontName);
            this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_ChnFontName);
            this.Grp_SetSelectedStyle.Controls.Add(this.Btn_Underline);
            this.Grp_SetSelectedStyle.Controls.Add(this.Btn_Italic);
            this.Grp_SetSelectedStyle.Controls.Add(this.Btn_Bold);
            this.Grp_SetSelectedStyle.Dock = System.Windows.Forms.DockStyle.Top;
            this.Grp_SetSelectedStyle.Location = new System.Drawing.Point(0, 265);
            this.Grp_SetSelectedStyle.Name = "Grp_SetSelectedStyle";
            this.Grp_SetSelectedStyle.Size = new System.Drawing.Size(1200, 195);
            this.Grp_SetSelectedStyle.TabIndex = 1;
            this.Grp_SetSelectedStyle.TabStop = false;
            this.Grp_SetSelectedStyle.Text = "为选中样式应用下列设置";
            // 
            // Txt_RightIndent
            // 
            this.Txt_RightIndent.Location = new System.Drawing.Point(495, 93);
            this.Txt_RightIndent.Name = "Txt_RightIndent";
            this.Txt_RightIndent.Size = new System.Drawing.Size(150, 23);
            this.Txt_RightIndent.TabIndex = 39;
            // 
            // Cmb_SpaceAfter
            // 
            this.Cmb_SpaceAfter.FormattingEnabled = true;
            this.Cmb_SpaceAfter.Location = new System.Drawing.Point(495, 126);
            this.Cmb_SpaceAfter.Name = "Cmb_SpaceAfter";
            this.Cmb_SpaceAfter.Size = new System.Drawing.Size(150, 25);
            this.Cmb_SpaceAfter.TabIndex = 27;
            this.Cmb_SpaceAfter.Validated += new System.EventHandler(this.Cmb_SpaceValue_Validated);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(426, 64);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(56, 17);
            this.label14.TabIndex = 54;
            this.label14.Text = "字体颜色";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(54, 164);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(56, 17);
            this.label13.TabIndex = 53;
            this.label13.Text = "段落对齐";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(745, 158);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(56, 17);
            this.label12.TabIndex = 52;
            this.label12.Text = "段前分页";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(426, 130);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(56, 17);
            this.label11.TabIndex = 51;
            this.label11.Text = "段后间距";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(54, 130);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(56, 17);
            this.label10.TabIndex = 50;
            this.label10.Text = "段前间距";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(426, 164);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(56, 17);
            this.label9.TabIndex = 49;
            this.label9.Text = "段落行距";
            this.label9.Click += new System.EventHandler(this.label9_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(426, 98);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(44, 17);
            this.label8.TabIndex = 48;
            this.label8.Text = "右缩进";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(54, 96);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(44, 17);
            this.label7.TabIndex = 47;
            this.label7.Text = "左缩进";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(757, 116);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 17);
            this.label6.TabIndex = 46;
            this.label6.Text = "下划线";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(769, 75);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 17);
            this.label5.TabIndex = 45;
            this.label5.Text = "斜体";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(770, 34);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 17);
            this.label4.TabIndex = 44;
            this.label4.Text = "粗体";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(54, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 17);
            this.label3.TabIndex = 43;
            this.label3.Text = "字体大小";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(425, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 17);
            this.label2.TabIndex = 42;
            this.label2.Text = "西文字体";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(54, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "中文字体";
            // 
            // Btn_FontColor
            // 
            this.Btn_FontColor.BackColor = System.Drawing.Color.Black;
            this.Btn_FontColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Btn_FontColor.Location = new System.Drawing.Point(494, 58);
            this.Btn_FontColor.Name = "Btn_FontColor";
            this.Btn_FontColor.Size = new System.Drawing.Size(151, 27);
            this.Btn_FontColor.TabIndex = 41;
            this.Btn_FontColor.UseVisualStyleBackColor = false;
            this.Btn_FontColor.Click += new System.EventHandler(this.Btn_FontColor_Click);
            // 
            // Txt_LeftIndent
            // 
            this.Txt_LeftIndent.Location = new System.Drawing.Point(124, 93);
            this.Txt_LeftIndent.Name = "Txt_LeftIndent";
            this.Txt_LeftIndent.Size = new System.Drawing.Size(150, 23);
            this.Txt_LeftIndent.TabIndex = 38;
            // 
            // Btn_BreakBefore
            // 
            this.Btn_BreakBefore.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_BreakBefore.Location = new System.Drawing.Point(803, 152);
            this.Btn_BreakBefore.Name = "Btn_BreakBefore";
            this.Btn_BreakBefore.Pressed = false;
            this.Btn_BreakBefore.Size = new System.Drawing.Size(40, 30);
            this.Btn_BreakBefore.TabIndex = 34;
            this.Btn_BreakBefore.Text = "否";
            this.Btn_BreakBefore.UseVisualStyleBackColor = false;
            this.Btn_BreakBefore.PressedChanged += new System.EventHandler(this.ToggleButton_PressedChanged);
            // 
            // Cmb_Alignment
            // 
            this.Cmb_Alignment.AllowInput = false;
            this.Cmb_Alignment.FormattingEnabled = true;
            this.Cmb_Alignment.Location = new System.Drawing.Point(124, 160);
            this.Cmb_Alignment.Name = "Cmb_Alignment";
            this.Cmb_Alignment.Size = new System.Drawing.Size(150, 25);
            this.Cmb_Alignment.TabIndex = 30;
            // 
            // Cmb_SpaceBefore
            // 
            this.Cmb_SpaceBefore.FormattingEnabled = true;
            this.Cmb_SpaceBefore.Location = new System.Drawing.Point(124, 126);
            this.Cmb_SpaceBefore.Name = "Cmb_SpaceBefore";
            this.Cmb_SpaceBefore.Size = new System.Drawing.Size(150, 25);
            this.Cmb_SpaceBefore.TabIndex = 26;
            this.Cmb_SpaceBefore.Validated += new System.EventHandler(this.Cmb_SpaceValue_Validated);
            // 
            // Cmb_FontSize
            // 
            this.Cmb_FontSize.FormattingEnabled = true;
            this.Cmb_FontSize.Location = new System.Drawing.Point(124, 58);
            this.Cmb_FontSize.Name = "Cmb_FontSize";
            this.Cmb_FontSize.Size = new System.Drawing.Size(150, 25);
            this.Cmb_FontSize.TabIndex = 24;
            this.Cmb_FontSize.Validated += new System.EventHandler(this.Cmb_FontSize_Validated);
            // 
            // Cmb_LineSpacing
            // 
            this.Cmb_LineSpacing.FormattingEnabled = true;
            this.Cmb_LineSpacing.Location = new System.Drawing.Point(495, 160);
            this.Cmb_LineSpacing.Name = "Cmb_LineSpacing";
            this.Cmb_LineSpacing.Size = new System.Drawing.Size(150, 25);
            this.Cmb_LineSpacing.TabIndex = 23;
            this.Cmb_LineSpacing.Validated += new System.EventHandler(this.Cmb_LineSpace_Validated);
            // 
            // Cmb_EngFontName
            // 
            this.Cmb_EngFontName.AllowInput = false;
            this.Cmb_EngFontName.FormattingEnabled = true;
            this.Cmb_EngFontName.Location = new System.Drawing.Point(495, 24);
            this.Cmb_EngFontName.Name = "Cmb_EngFontName";
            this.Cmb_EngFontName.Size = new System.Drawing.Size(150, 25);
            this.Cmb_EngFontName.TabIndex = 15;
            // 
            // Cmb_ChnFontName
            // 
            this.Cmb_ChnFontName.AllowInput = false;
            this.Cmb_ChnFontName.FormattingEnabled = true;
            this.Cmb_ChnFontName.Location = new System.Drawing.Point(124, 24);
            this.Cmb_ChnFontName.Name = "Cmb_ChnFontName";
            this.Cmb_ChnFontName.Size = new System.Drawing.Size(150, 25);
            this.Cmb_ChnFontName.TabIndex = 14;
            // 
            // Btn_Underline
            // 
            this.Btn_Underline.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_Underline.Location = new System.Drawing.Point(803, 110);
            this.Btn_Underline.Name = "Btn_Underline";
            this.Btn_Underline.Pressed = false;
            this.Btn_Underline.Size = new System.Drawing.Size(40, 30);
            this.Btn_Underline.TabIndex = 5;
            this.Btn_Underline.Text = "否";
            this.Btn_Underline.UseVisualStyleBackColor = false;
            this.Btn_Underline.PressedChanged += new System.EventHandler(this.ToggleButton_PressedChanged);
            // 
            // Btn_Italic
            // 
            this.Btn_Italic.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_Italic.Location = new System.Drawing.Point(803, 69);
            this.Btn_Italic.Name = "Btn_Italic";
            this.Btn_Italic.Pressed = false;
            this.Btn_Italic.Size = new System.Drawing.Size(40, 30);
            this.Btn_Italic.TabIndex = 4;
            this.Btn_Italic.Text = "否";
            this.Btn_Italic.UseVisualStyleBackColor = false;
            this.Btn_Italic.PressedChanged += new System.EventHandler(this.ToggleButton_PressedChanged);
            // 
            // Btn_Bold
            // 
            this.Btn_Bold.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_Bold.Location = new System.Drawing.Point(803, 28);
            this.Btn_Bold.Name = "Btn_Bold";
            this.Btn_Bold.Pressed = false;
            this.Btn_Bold.Size = new System.Drawing.Size(40, 30);
            this.Btn_Bold.TabIndex = 3;
            this.Btn_Bold.Text = "否";
            this.Btn_Bold.UseVisualStyleBackColor = false;
            this.Btn_Bold.PressedChanged += new System.EventHandler(this.ToggleButton_PressedChanged);
            // 
            // Btn_SetStyles
            // 
            this.Btn_SetStyles.Location = new System.Drawing.Point(946, 478);
            this.Btn_SetStyles.Name = "Btn_SetStyles";
            this.Btn_SetStyles.Size = new System.Drawing.Size(100, 30);
            this.Btn_SetStyles.TabIndex = 2;
            this.Btn_SetStyles.Text = "确定";
            this.Btn_SetStyles.UseVisualStyleBackColor = true;
            this.Btn_SetStyles.Click += new System.EventHandler(this.Btn_SetStyles_Click);
            // 
            // Btn_Cancel
            // 
            this.Btn_Cancel.Location = new System.Drawing.Point(1074, 478);
            this.Btn_Cancel.Name = "Btn_Cancel";
            this.Btn_Cancel.Size = new System.Drawing.Size(100, 30);
            this.Btn_Cancel.TabIndex = 3;
            this.Btn_Cancel.Text = "取消";
            this.Btn_Cancel.UseVisualStyleBackColor = true;
            this.Btn_Cancel.Click += new System.EventHandler(this.Btn_Cancel_Click);
            // 
            // LevelStyleSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(1200, 520);
            this.Controls.Add(this.Btn_Cancel);
            this.Controls.Add(this.Btn_SetStyles);
            this.Controls.Add(this.Grp_SetSelectedStyle);
            this.Controls.Add(this.Dta_StyleList);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LevelStyleSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "多级段落样式设置";
            ((System.ComponentModel.ISupportInitialize)(this.Dta_StyleList)).EndInit();
            this.Grp_SetSelectedStyle.ResumeLayout(false);
            this.Grp_SetSelectedStyle.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView Dta_StyleList;
        private System.Windows.Forms.GroupBox Grp_SetSelectedStyle;
        private WordMan_VSTO.StandardNumericUpDown Txt_RightIndent;
        private WordMan_VSTO.StandardComboBox Cmb_SpaceAfter;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Btn_FontColor;
        private WordMan_VSTO.StandardNumericUpDown Txt_LeftIndent;
        private WordMan_VSTO.ToggleButton Btn_BreakBefore;
        private WordMan_VSTO.StandardComboBox Cmb_Alignment;
        private WordMan_VSTO.StandardComboBox Cmb_SpaceBefore;
        private WordMan_VSTO.StandardComboBox Cmb_FontSize;
        private WordMan_VSTO.StandardComboBox Cmb_LineSpacing;
        private WordMan_VSTO.StandardComboBox Cmb_EngFontName;
        private WordMan_VSTO.StandardComboBox Cmb_ChnFontName;
        private WordMan_VSTO.ToggleButton Btn_Underline;
        private WordMan_VSTO.ToggleButton Btn_Italic;
        private WordMan_VSTO.ToggleButton Btn_Bold;
        private System.Windows.Forms.Button Btn_SetStyles;
        private System.Windows.Forms.Button Btn_Cancel;
    }
}
