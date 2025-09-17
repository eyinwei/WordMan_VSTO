namespace WordMan_VSTO.StylePane
{
    partial class DocumentStyleSettings
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改
        /// 此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Lbl_Title = new System.Windows.Forms.Label();
            this.Lst_Styles = new System.Windows.Forms.ListBox();
            this.Lbl_StyleList = new System.Windows.Forms.Label();
            this.Grp_Font = new System.Windows.Forms.GroupBox();
            this.Lbl_ChnFont = new System.Windows.Forms.Label();
            this.Cmb_ChnFontName = new System.Windows.Forms.ComboBox();
            this.Lbl_EngFont = new System.Windows.Forms.Label();
            this.Cmb_EngFontName = new System.Windows.Forms.ComboBox();
            this.Lbl_FontSize = new System.Windows.Forms.Label();
            this.Cmb_FontSize = new System.Windows.Forms.ComboBox();
            this.Chk_Bold = new System.Windows.Forms.CheckBox();
            this.Chk_Italic = new System.Windows.Forms.CheckBox();
            this.Chk_Underline = new System.Windows.Forms.CheckBox();
            this.Grp_Paragraph = new System.Windows.Forms.GroupBox();
            this.Lbl_Alignment = new System.Windows.Forms.Label();
            this.Cmb_Alignment = new System.Windows.Forms.ComboBox();
            this.Lbl_LineSpacing = new System.Windows.Forms.Label();
            this.Cmb_LineSpacing = new System.Windows.Forms.ComboBox();
            this.Lbl_SpaceBefore = new System.Windows.Forms.Label();
            this.Cmb_SpaceBefore = new System.Windows.Forms.ComboBox();
            this.Lbl_SpaceAfter = new System.Windows.Forms.Label();
            this.Cmb_SpaceAfter = new System.Windows.Forms.ComboBox();
            this.Grp_Indent = new System.Windows.Forms.GroupBox();
            this.Lbl_LeftIndent = new System.Windows.Forms.Label();
            this.Txt_LeftIndent = new System.Windows.Forms.TextBox();
            this.Lbl_RightIndent = new System.Windows.Forms.Label();
            this.Txt_RightIndent = new System.Windows.Forms.TextBox();
            this.Lbl_FirstLineIndent = new System.Windows.Forms.Label();
            this.Txt_FirstLineIndent = new System.Windows.Forms.TextBox();
            this.Btn_Apply = new System.Windows.Forms.Button();
            this.Btn_ApplyAll = new System.Windows.Forms.Button();
            this.Btn_Reset = new System.Windows.Forms.Button();
            this.Btn_ResetAll = new System.Windows.Forms.Button();
            this.Btn_Save = new System.Windows.Forms.Button();
            this.Btn_Load = new System.Windows.Forms.Button();
            this.Grp_Font.SuspendLayout();
            this.Grp_Paragraph.SuspendLayout();
            this.Grp_Indent.SuspendLayout();
            this.SuspendLayout();
            // 
            // Lbl_Title
            // 
            this.Lbl_Title.AutoSize = true;
            this.Lbl_Title.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Lbl_Title.Location = new System.Drawing.Point(15, 15);
            this.Lbl_Title.Name = "Lbl_Title";
            this.Lbl_Title.Size = new System.Drawing.Size(89, 22);
            this.Lbl_Title.TabIndex = 0;
            this.Lbl_Title.Text = "文档样式设置";
            // 
            // Lst_Styles
            // 
            this.Lst_Styles.FormattingEnabled = true;
            this.Lst_Styles.ItemHeight = 12;
            this.Lst_Styles.Location = new System.Drawing.Point(15, 50);
            this.Lst_Styles.Name = "Lst_Styles";
            this.Lst_Styles.Size = new System.Drawing.Size(150, 400);
            this.Lst_Styles.TabIndex = 1;
            this.Lst_Styles.SelectedIndexChanged += new System.EventHandler(this.Lst_Styles_SelectedIndexChanged);
            // 
            // Lbl_StyleList
            // 
            this.Lbl_StyleList.AutoSize = true;
            this.Lbl_StyleList.Location = new System.Drawing.Point(15, 35);
            this.Lbl_StyleList.Name = "Lbl_StyleList";
            this.Lbl_StyleList.Size = new System.Drawing.Size(65, 12);
            this.Lbl_StyleList.TabIndex = 2;
            this.Lbl_StyleList.Text = "样式列表：";
            // 
            // Grp_Font
            // 
            this.Grp_Font.Controls.Add(this.Chk_Underline);
            this.Grp_Font.Controls.Add(this.Chk_Italic);
            this.Grp_Font.Controls.Add(this.Chk_Bold);
            this.Grp_Font.Controls.Add(this.Cmb_FontSize);
            this.Grp_Font.Controls.Add(this.Lbl_FontSize);
            this.Grp_Font.Controls.Add(this.Cmb_EngFontName);
            this.Grp_Font.Controls.Add(this.Lbl_EngFont);
            this.Grp_Font.Controls.Add(this.Cmb_ChnFontName);
            this.Grp_Font.Controls.Add(this.Lbl_ChnFont);
            this.Grp_Font.Location = new System.Drawing.Point(180, 50);
            this.Grp_Font.Name = "Grp_Font";
            this.Grp_Font.Size = new System.Drawing.Size(350, 120);
            this.Grp_Font.TabIndex = 3;
            this.Grp_Font.TabStop = false;
            this.Grp_Font.Text = "字体设置";
            // 
            // Lbl_ChnFont
            // 
            this.Lbl_ChnFont.AutoSize = true;
            this.Lbl_ChnFont.Location = new System.Drawing.Point(15, 25);
            this.Lbl_ChnFont.Name = "Lbl_ChnFont";
            this.Lbl_ChnFont.Size = new System.Drawing.Size(65, 12);
            this.Lbl_ChnFont.TabIndex = 0;
            this.Lbl_ChnFont.Text = "中文字体：";
            // 
            // Cmb_ChnFontName
            // 
            this.Cmb_ChnFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_ChnFontName.FormattingEnabled = true;
            this.Cmb_ChnFontName.Location = new System.Drawing.Point(85, 22);
            this.Cmb_ChnFontName.Name = "Cmb_ChnFontName";
            this.Cmb_ChnFontName.Size = new System.Drawing.Size(120, 20);
            this.Cmb_ChnFontName.TabIndex = 1;
            // 
            // Lbl_EngFont
            // 
            this.Lbl_EngFont.AutoSize = true;
            this.Lbl_EngFont.Location = new System.Drawing.Point(15, 55);
            this.Lbl_EngFont.Name = "Lbl_EngFont";
            this.Lbl_EngFont.Size = new System.Drawing.Size(65, 12);
            this.Lbl_EngFont.TabIndex = 2;
            this.Lbl_EngFont.Text = "西文字体：";
            // 
            // Cmb_EngFontName
            // 
            this.Cmb_EngFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_EngFontName.FormattingEnabled = true;
            this.Cmb_EngFontName.Location = new System.Drawing.Point(85, 52);
            this.Cmb_EngFontName.Name = "Cmb_EngFontName";
            this.Cmb_EngFontName.Size = new System.Drawing.Size(120, 20);
            this.Cmb_EngFontName.TabIndex = 3;
            // 
            // Lbl_FontSize
            // 
            this.Lbl_FontSize.AutoSize = true;
            this.Lbl_FontSize.Location = new System.Drawing.Point(15, 85);
            this.Lbl_FontSize.Name = "Lbl_FontSize";
            this.Lbl_FontSize.Size = new System.Drawing.Size(65, 12);
            this.Lbl_FontSize.TabIndex = 4;
            this.Lbl_FontSize.Text = "字体大小：";
            // 
            // Cmb_FontSize
            // 
            this.Cmb_FontSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_FontSize.FormattingEnabled = true;
            this.Cmb_FontSize.Location = new System.Drawing.Point(85, 82);
            this.Cmb_FontSize.Name = "Cmb_FontSize";
            this.Cmb_FontSize.Size = new System.Drawing.Size(80, 20);
            this.Cmb_FontSize.TabIndex = 5;
            // 
            // Chk_Bold
            // 
            this.Chk_Bold.AutoSize = true;
            this.Chk_Bold.Location = new System.Drawing.Point(220, 25);
            this.Chk_Bold.Name = "Chk_Bold";
            this.Chk_Bold.Size = new System.Drawing.Size(48, 16);
            this.Chk_Bold.TabIndex = 6;
            this.Chk_Bold.Text = "加粗";
            this.Chk_Bold.UseVisualStyleBackColor = true;
            // 
            // Chk_Italic
            // 
            this.Chk_Italic.AutoSize = true;
            this.Chk_Italic.Location = new System.Drawing.Point(280, 25);
            this.Chk_Italic.Name = "Chk_Italic";
            this.Chk_Italic.Size = new System.Drawing.Size(48, 16);
            this.Chk_Italic.TabIndex = 7;
            this.Chk_Italic.Text = "斜体";
            this.Chk_Italic.UseVisualStyleBackColor = true;
            // 
            // Chk_Underline
            // 
            this.Chk_Underline.AutoSize = true;
            this.Chk_Underline.Location = new System.Drawing.Point(220, 50);
            this.Chk_Underline.Name = "Chk_Underline";
            this.Chk_Underline.Size = new System.Drawing.Size(60, 16);
            this.Chk_Underline.TabIndex = 8;
            this.Chk_Underline.Text = "下划线";
            this.Chk_Underline.UseVisualStyleBackColor = true;
            // 
            // Grp_Paragraph
            // 
            this.Grp_Paragraph.Controls.Add(this.Cmb_SpaceAfter);
            this.Grp_Paragraph.Controls.Add(this.Lbl_SpaceAfter);
            this.Grp_Paragraph.Controls.Add(this.Cmb_SpaceBefore);
            this.Grp_Paragraph.Controls.Add(this.Lbl_SpaceBefore);
            this.Grp_Paragraph.Controls.Add(this.Cmb_LineSpacing);
            this.Grp_Paragraph.Controls.Add(this.Lbl_LineSpacing);
            this.Grp_Paragraph.Controls.Add(this.Cmb_Alignment);
            this.Grp_Paragraph.Controls.Add(this.Lbl_Alignment);
            this.Grp_Paragraph.Location = new System.Drawing.Point(180, 185);
            this.Grp_Paragraph.Name = "Grp_Paragraph";
            this.Grp_Paragraph.Size = new System.Drawing.Size(350, 120);
            this.Grp_Paragraph.TabIndex = 4;
            this.Grp_Paragraph.TabStop = false;
            this.Grp_Paragraph.Text = "段落设置";
            // 
            // Lbl_Alignment
            // 
            this.Lbl_Alignment.AutoSize = true;
            this.Lbl_Alignment.Location = new System.Drawing.Point(15, 25);
            this.Lbl_Alignment.Name = "Lbl_Alignment";
            this.Lbl_Alignment.Size = new System.Drawing.Size(65, 12);
            this.Lbl_Alignment.TabIndex = 0;
            this.Lbl_Alignment.Text = "对齐方式：";
            // 
            // Cmb_Alignment
            // 
            this.Cmb_Alignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_Alignment.FormattingEnabled = true;
            this.Cmb_Alignment.Location = new System.Drawing.Point(85, 22);
            this.Cmb_Alignment.Name = "Cmb_Alignment";
            this.Cmb_Alignment.Size = new System.Drawing.Size(100, 20);
            this.Cmb_Alignment.TabIndex = 1;
            // 
            // Lbl_LineSpacing
            // 
            this.Lbl_LineSpacing.AutoSize = true;
            this.Lbl_LineSpacing.Location = new System.Drawing.Point(15, 55);
            this.Lbl_LineSpacing.Name = "Lbl_LineSpacing";
            this.Lbl_LineSpacing.Size = new System.Drawing.Size(65, 12);
            this.Lbl_LineSpacing.TabIndex = 2;
            this.Lbl_LineSpacing.Text = "行距设置：";
            // 
            // Cmb_LineSpacing
            // 
            this.Cmb_LineSpacing.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_LineSpacing.FormattingEnabled = true;
            this.Cmb_LineSpacing.Location = new System.Drawing.Point(85, 52);
            this.Cmb_LineSpacing.Name = "Cmb_LineSpacing";
            this.Cmb_LineSpacing.Size = new System.Drawing.Size(100, 20);
            this.Cmb_LineSpacing.TabIndex = 3;
            // 
            // Lbl_SpaceBefore
            // 
            this.Lbl_SpaceBefore.AutoSize = true;
            this.Lbl_SpaceBefore.Location = new System.Drawing.Point(200, 25);
            this.Lbl_SpaceBefore.Name = "Lbl_SpaceBefore";
            this.Lbl_SpaceBefore.Size = new System.Drawing.Size(53, 12);
            this.Lbl_SpaceBefore.TabIndex = 4;
            this.Lbl_SpaceBefore.Text = "段前距：";
            // 
            // Cmb_SpaceBefore
            // 
            this.Cmb_SpaceBefore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_SpaceBefore.FormattingEnabled = true;
            this.Cmb_SpaceBefore.Location = new System.Drawing.Point(260, 22);
            this.Cmb_SpaceBefore.Name = "Cmb_SpaceBefore";
            this.Cmb_SpaceBefore.Size = new System.Drawing.Size(80, 20);
            this.Cmb_SpaceBefore.TabIndex = 5;
            // 
            // Lbl_SpaceAfter
            // 
            this.Lbl_SpaceAfter.AutoSize = true;
            this.Lbl_SpaceAfter.Location = new System.Drawing.Point(200, 55);
            this.Lbl_SpaceAfter.Name = "Lbl_SpaceAfter";
            this.Lbl_SpaceAfter.Size = new System.Drawing.Size(53, 12);
            this.Lbl_SpaceAfter.TabIndex = 6;
            this.Lbl_SpaceAfter.Text = "段后距：";
            // 
            // Cmb_SpaceAfter
            // 
            this.Cmb_SpaceAfter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_SpaceAfter.FormattingEnabled = true;
            this.Cmb_SpaceAfter.Location = new System.Drawing.Point(260, 52);
            this.Cmb_SpaceAfter.Name = "Cmb_SpaceAfter";
            this.Cmb_SpaceAfter.Size = new System.Drawing.Size(80, 20);
            this.Cmb_SpaceAfter.TabIndex = 7;
            // 
            // Grp_Indent
            // 
            this.Grp_Indent.Controls.Add(this.Txt_FirstLineIndent);
            this.Grp_Indent.Controls.Add(this.Lbl_FirstLineIndent);
            this.Grp_Indent.Controls.Add(this.Txt_RightIndent);
            this.Grp_Indent.Controls.Add(this.Lbl_RightIndent);
            this.Grp_Indent.Controls.Add(this.Txt_LeftIndent);
            this.Grp_Indent.Controls.Add(this.Lbl_LeftIndent);
            this.Grp_Indent.Location = new System.Drawing.Point(180, 320);
            this.Grp_Indent.Name = "Grp_Indent";
            this.Grp_Indent.Size = new System.Drawing.Size(350, 100);
            this.Grp_Indent.TabIndex = 5;
            this.Grp_Indent.TabStop = false;
            this.Grp_Indent.Text = "缩进设置";
            // 
            // Lbl_LeftIndent
            // 
            this.Lbl_LeftIndent.AutoSize = true;
            this.Lbl_LeftIndent.Location = new System.Drawing.Point(15, 25);
            this.Lbl_LeftIndent.Name = "Lbl_LeftIndent";
            this.Lbl_LeftIndent.Size = new System.Drawing.Size(65, 12);
            this.Lbl_LeftIndent.TabIndex = 0;
            this.Lbl_LeftIndent.Text = "左缩进：";
            // 
            // Txt_LeftIndent
            // 
            this.Txt_LeftIndent.Location = new System.Drawing.Point(85, 22);
            this.Txt_LeftIndent.Name = "Txt_LeftIndent";
            this.Txt_LeftIndent.Size = new System.Drawing.Size(80, 21);
            this.Txt_LeftIndent.TabIndex = 1;
            // 
            // Lbl_RightIndent
            // 
            this.Lbl_RightIndent.AutoSize = true;
            this.Lbl_RightIndent.Location = new System.Drawing.Point(15, 55);
            this.Lbl_RightIndent.Name = "Lbl_RightIndent";
            this.Lbl_RightIndent.Size = new System.Drawing.Size(65, 12);
            this.Lbl_RightIndent.TabIndex = 2;
            this.Lbl_RightIndent.Text = "右缩进：";
            // 
            // Txt_RightIndent
            // 
            this.Txt_RightIndent.Location = new System.Drawing.Point(85, 52);
            this.Txt_RightIndent.Name = "Txt_RightIndent";
            this.Txt_RightIndent.Size = new System.Drawing.Size(80, 21);
            this.Txt_RightIndent.TabIndex = 3;
            // 
            // Lbl_FirstLineIndent
            // 
            this.Lbl_FirstLineIndent.AutoSize = true;
            this.Lbl_FirstLineIndent.Location = new System.Drawing.Point(180, 25);
            this.Lbl_FirstLineIndent.Name = "Lbl_FirstLineIndent";
            this.Lbl_FirstLineIndent.Size = new System.Drawing.Size(77, 12);
            this.Lbl_FirstLineIndent.TabIndex = 4;
            this.Lbl_FirstLineIndent.Text = "首行缩进：";
            // 
            // Txt_FirstLineIndent
            // 
            this.Txt_FirstLineIndent.Location = new System.Drawing.Point(260, 22);
            this.Txt_FirstLineIndent.Name = "Txt_FirstLineIndent";
            this.Txt_FirstLineIndent.Size = new System.Drawing.Size(80, 21);
            this.Txt_FirstLineIndent.TabIndex = 5;
            // 
            // Btn_Apply
            // 
            this.Btn_Apply.Location = new System.Drawing.Point(15, 470);
            this.Btn_Apply.Name = "Btn_Apply";
            this.Btn_Apply.Size = new System.Drawing.Size(75, 30);
            this.Btn_Apply.TabIndex = 6;
            this.Btn_Apply.Text = "应用当前";
            this.Btn_Apply.UseVisualStyleBackColor = true;
            this.Btn_Apply.Click += new System.EventHandler(this.Btn_Apply_Click);
            // 
            // Btn_ApplyAll
            // 
            this.Btn_ApplyAll.Location = new System.Drawing.Point(100, 470);
            this.Btn_ApplyAll.Name = "Btn_ApplyAll";
            this.Btn_ApplyAll.Size = new System.Drawing.Size(75, 30);
            this.Btn_ApplyAll.TabIndex = 7;
            this.Btn_ApplyAll.Text = "应用全部";
            this.Btn_ApplyAll.UseVisualStyleBackColor = true;
            this.Btn_ApplyAll.Click += new System.EventHandler(this.Btn_ApplyAll_Click);
            // 
            // Btn_Reset
            // 
            this.Btn_Reset.Location = new System.Drawing.Point(185, 470);
            this.Btn_Reset.Name = "Btn_Reset";
            this.Btn_Reset.Size = new System.Drawing.Size(75, 30);
            this.Btn_Reset.TabIndex = 8;
            this.Btn_Reset.Text = "重置当前";
            this.Btn_Reset.UseVisualStyleBackColor = true;
            this.Btn_Reset.Click += new System.EventHandler(this.Btn_Reset_Click);
            // 
            // Btn_ResetAll
            // 
            this.Btn_ResetAll.Location = new System.Drawing.Point(270, 470);
            this.Btn_ResetAll.Name = "Btn_ResetAll";
            this.Btn_ResetAll.Size = new System.Drawing.Size(75, 30);
            this.Btn_ResetAll.TabIndex = 9;
            this.Btn_ResetAll.Text = "重置全部";
            this.Btn_ResetAll.UseVisualStyleBackColor = true;
            this.Btn_ResetAll.Click += new System.EventHandler(this.Btn_ResetAll_Click);
            // 
            // Btn_Save
            // 
            this.Btn_Save.Location = new System.Drawing.Point(360, 470);
            this.Btn_Save.Name = "Btn_Save";
            this.Btn_Save.Size = new System.Drawing.Size(75, 30);
            this.Btn_Save.TabIndex = 10;
            this.Btn_Save.Text = "保存配置";
            this.Btn_Save.UseVisualStyleBackColor = true;
            this.Btn_Save.Click += new System.EventHandler(this.Btn_Save_Click);
            // 
            // Btn_Load
            // 
            this.Btn_Load.Location = new System.Drawing.Point(445, 470);
            this.Btn_Load.Name = "Btn_Load";
            this.Btn_Load.Size = new System.Drawing.Size(75, 30);
            this.Btn_Load.TabIndex = 11;
            this.Btn_Load.Text = "加载配置";
            this.Btn_Load.UseVisualStyleBackColor = true;
            this.Btn_Load.Click += new System.EventHandler(this.Btn_Load_Click);
            // 
            // DocumentStyleSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.Btn_Load);
            this.Controls.Add(this.Btn_Save);
            this.Controls.Add(this.Btn_ResetAll);
            this.Controls.Add(this.Btn_Reset);
            this.Controls.Add(this.Btn_ApplyAll);
            this.Controls.Add(this.Btn_Apply);
            this.Controls.Add(this.Grp_Indent);
            this.Controls.Add(this.Grp_Paragraph);
            this.Controls.Add(this.Grp_Font);
            this.Controls.Add(this.Lbl_StyleList);
            this.Controls.Add(this.Lst_Styles);
            this.Controls.Add(this.Lbl_Title);
            this.Name = "DocumentStyleSettings";
            this.Size = new System.Drawing.Size(550, 520);
            this.Grp_Font.ResumeLayout(false);
            this.Grp_Font.PerformLayout();
            this.Grp_Paragraph.ResumeLayout(false);
            this.Grp_Paragraph.PerformLayout();
            this.Grp_Indent.ResumeLayout(false);
            this.Grp_Indent.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Lbl_Title;
        private System.Windows.Forms.ListBox Lst_Styles;
        private System.Windows.Forms.Label Lbl_StyleList;
        private System.Windows.Forms.GroupBox Grp_Font;
        private System.Windows.Forms.Label Lbl_ChnFont;
        private System.Windows.Forms.ComboBox Cmb_ChnFontName;
        private System.Windows.Forms.Label Lbl_EngFont;
        private System.Windows.Forms.ComboBox Cmb_EngFontName;
        private System.Windows.Forms.Label Lbl_FontSize;
        private System.Windows.Forms.ComboBox Cmb_FontSize;
        private System.Windows.Forms.CheckBox Chk_Bold;
        private System.Windows.Forms.CheckBox Chk_Italic;
        private System.Windows.Forms.CheckBox Chk_Underline;
        private System.Windows.Forms.GroupBox Grp_Paragraph;
        private System.Windows.Forms.Label Lbl_Alignment;
        private System.Windows.Forms.ComboBox Cmb_Alignment;
        private System.Windows.Forms.Label Lbl_LineSpacing;
        private System.Windows.Forms.ComboBox Cmb_LineSpacing;
        private System.Windows.Forms.Label Lbl_SpaceBefore;
        private System.Windows.Forms.ComboBox Cmb_SpaceBefore;
        private System.Windows.Forms.Label Lbl_SpaceAfter;
        private System.Windows.Forms.ComboBox Cmb_SpaceAfter;
        private System.Windows.Forms.GroupBox Grp_Indent;
        private System.Windows.Forms.Label Lbl_LeftIndent;
        private System.Windows.Forms.TextBox Txt_LeftIndent;
        private System.Windows.Forms.Label Lbl_RightIndent;
        private System.Windows.Forms.TextBox Txt_RightIndent;
        private System.Windows.Forms.Label Lbl_FirstLineIndent;
        private System.Windows.Forms.TextBox Txt_FirstLineIndent;
        private System.Windows.Forms.Button Btn_Apply;
        private System.Windows.Forms.Button Btn_ApplyAll;
        private System.Windows.Forms.Button Btn_Reset;
        private System.Windows.Forms.Button Btn_ResetAll;
        private System.Windows.Forms.Button Btn_Save;
        private System.Windows.Forms.Button Btn_Load;
    }
}
