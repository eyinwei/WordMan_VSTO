using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WordMan_VSTO;
using WordMan_VSTO.MultiLevel;

namespace WordMan_VSTO
{
    partial class StyleSettings
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region 控件声明
        private GroupBox groupBox3;
        private StandardButton Btn_DelStyle;
        private Label label9;
        private Label label10;
        private Label label11;
        private GroupBox Pal_Font;
        private Label label3;
        private StandardComboBox Cmb_ChnFontName;
        private Label label17;
        private StandardComboBox Cmb_EngFontName;
        private Label label4;
        private StandardComboBox Cmb_FontSize;
        private StandardButton Btn_FontColor;
        private ToggleButton Btn_Bold;
        private ToggleButton Btn_Italic;
        private ToggleButton Btn_UnderLine;
        private GroupBox Pal_ParaIndent;
        private Label label15;
        private StandardComboBox Cmb_ParaAligment;
        private Label label6;
        private Label label7;
        private Label label18;
        private StandardNumericUpDown Nud_RightIndent;
        private Label label19;
        private StandardComboBox Cmb_FirstLineIndentType;
        private StandardButton Btn_AddStyle;
        private StandardComboBox Cmb_SetLevel;
        private Label label12;
        private ListBox Lst_Styles;
        private StandardButton Btn_ApplySet;
        private StandardComboBox Cmb_PreSettings;
        private Label label16;
        private StandardTextBox Txt_AddStyleName;
        private StandardComboBox Cmb_LineSpacing;
        private StandardNumericUpDown Nud_LineSpacing;
        private StandardNumericUpDown Nud_BefreSpacing;
        private StandardNumericUpDown Nud_AfterSpacing;
        private StandardNumericUpDown Nud_LeftIndent;
        private StandardNumericUpDown Nud_FirstLineIndent;
        private StandardNumericUpDown Nud_FirstLineIndentByChar;

        #endregion

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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改
        /// 此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.Txt_AddStyleName = new WordMan_VSTO.StandardTextBox();
            this.Btn_DelStyle = new WordMan_VSTO.StandardButton();
            this.Pal_Font = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Cmb_ChnFontName = new WordMan_VSTO.StandardComboBox();
            this.label17 = new System.Windows.Forms.Label();
            this.Cmb_EngFontName = new WordMan_VSTO.StandardComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Cmb_FontSize = new WordMan_VSTO.StandardComboBox();
            this.Btn_FontColor = new WordMan_VSTO.StandardButton();
            this.Btn_Bold = new WordMan_VSTO.ToggleButton();
            this.Btn_Italic = new WordMan_VSTO.ToggleButton();
            this.Btn_UnderLine = new WordMan_VSTO.ToggleButton();
            this.Pal_ParaIndent = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.Nud_LineSpacing = new WordMan_VSTO.StandardNumericUpDown();
            this.label10 = new System.Windows.Forms.Label();
            this.Nud_BefreSpacing = new WordMan_VSTO.StandardNumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.Nud_AfterSpacing = new WordMan_VSTO.StandardNumericUpDown();
            this.label15 = new System.Windows.Forms.Label();
            this.Cmb_ParaAligment = new WordMan_VSTO.StandardComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.Nud_LeftIndent = new WordMan_VSTO.StandardNumericUpDown();
            this.label18 = new System.Windows.Forms.Label();
            this.Nud_RightIndent = new WordMan_VSTO.StandardNumericUpDown();
            this.label19 = new System.Windows.Forms.Label();
            this.Cmb_FirstLineIndentType = new WordMan_VSTO.StandardComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.Nud_FirstLineIndent = new WordMan_VSTO.StandardNumericUpDown();
            this.Nud_FirstLineIndentByChar = new WordMan_VSTO.StandardNumericUpDown();
            this.Btn_ReadDocumentStyle = new WordMan_VSTO.StandardButton();
            this.Btn_AddStyle = new WordMan_VSTO.StandardButton();
            this.Lst_Styles = new System.Windows.Forms.ListBox();
            this.Cmb_SetLevel = new WordMan_VSTO.StandardComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.Btn_ApplySet = new WordMan_VSTO.StandardButton();
            this.Cmb_PreSettings = new WordMan_VSTO.StandardComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.Cmb_LineSpacing = new WordMan_VSTO.StandardComboBox();
            this.groupBox3.SuspendLayout();
            this.Pal_Font.SuspendLayout();
            this.Pal_ParaIndent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_LineSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_BefreSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_AfterSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_LeftIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_RightIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_FirstLineIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_FirstLineIndentByChar)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.Txt_AddStyleName);
            this.groupBox3.Controls.Add(this.Btn_ApplySet);
            this.groupBox3.Controls.Add(this.Btn_DelStyle);
            this.groupBox3.Controls.Add(this.Cmb_SetLevel);
            this.groupBox3.Controls.Add(this.Cmb_PreSettings);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.Pal_Font);
            this.groupBox3.Controls.Add(this.label16);
            this.groupBox3.Controls.Add(this.Pal_ParaIndent);
            this.groupBox3.Controls.Add(this.Btn_ReadDocumentStyle);
            this.groupBox3.Controls.Add(this.Btn_AddStyle);
            this.groupBox3.Controls.Add(this.Lst_Styles);
            this.groupBox3.Location = new System.Drawing.Point(0, 0);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(679, 374);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "样式设置";
            // 
            // Txt_AddStyleName
            // 
            this.Txt_AddStyleName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Txt_AddStyleName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Txt_AddStyleName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Txt_AddStyleName.Location = new System.Drawing.Point(14, 305);
            this.Txt_AddStyleName.MaxLength = 0;
            this.Txt_AddStyleName.Name = "Txt_AddStyleName";
            this.Txt_AddStyleName.Size = new System.Drawing.Size(158, 23);
            this.Txt_AddStyleName.TabIndex = 33;
            // 
            // Btn_DelStyle
            // 
            this.Btn_DelStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.Btn_DelStyle.Enabled = false;
            this.Btn_DelStyle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Btn_DelStyle.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.Btn_DelStyle.ForeColor = System.Drawing.Color.Black;
            this.Btn_DelStyle.Location = new System.Drawing.Point(100, 334);
            this.Btn_DelStyle.Name = "Btn_DelStyle";
            this.Btn_DelStyle.Size = new System.Drawing.Size(72, 31);
            this.Btn_DelStyle.TabIndex = 29;
            this.Btn_DelStyle.Text = "删除样式";
            this.Btn_DelStyle.UseVisualStyleBackColor = false;
            // 
            // Pal_Font
            // 
            this.Pal_Font.Controls.Add(this.label3);
            this.Pal_Font.Controls.Add(this.Cmb_ChnFontName);
            this.Pal_Font.Controls.Add(this.label17);
            this.Pal_Font.Controls.Add(this.Cmb_EngFontName);
            this.Pal_Font.Controls.Add(this.label4);
            this.Pal_Font.Controls.Add(this.Cmb_FontSize);
            this.Pal_Font.Controls.Add(this.Btn_FontColor);
            this.Pal_Font.Controls.Add(this.Btn_Bold);
            this.Pal_Font.Controls.Add(this.Btn_Italic);
            this.Pal_Font.Controls.Add(this.Btn_UnderLine);
            this.Pal_Font.Enabled = false;
            this.Pal_Font.Location = new System.Drawing.Point(182, 17);
            this.Pal_Font.Name = "Pal_Font";
            this.Pal_Font.Size = new System.Drawing.Size(486, 100);
            this.Pal_Font.TabIndex = 7;
            this.Pal_Font.TabStop = false;
            this.Pal_Font.Text = "字体设置";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(10, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 28);
            this.label3.TabIndex = 1;
            this.label3.Text = "中文字体";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Cmb_ChnFontName
            // 
            this.Cmb_ChnFontName.AllowInput = false;
            this.Cmb_ChnFontName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_ChnFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_ChnFontName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_ChnFontName.FormattingEnabled = true;
            this.Cmb_ChnFontName.Location = new System.Drawing.Point(96, 22);
            this.Cmb_ChnFontName.Name = "Cmb_ChnFontName";
            this.Cmb_ChnFontName.Size = new System.Drawing.Size(120, 25);
            this.Cmb_ChnFontName.TabIndex = 2;
            // 
            // label17
            // 
            this.label17.Location = new System.Drawing.Point(288, 20);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(60, 28);
            this.label17.TabIndex = 2;
            this.label17.Text = "西文字体";
            this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Cmb_EngFontName
            // 
            this.Cmb_EngFontName.AllowInput = false;
            this.Cmb_EngFontName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_EngFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_EngFontName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_EngFontName.FormattingEnabled = true;
            this.Cmb_EngFontName.Location = new System.Drawing.Point(354, 22);
            this.Cmb_EngFontName.Name = "Cmb_EngFontName";
            this.Cmb_EngFontName.Size = new System.Drawing.Size(120, 25);
            this.Cmb_EngFontName.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(10, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 28);
            this.label4.TabIndex = 4;
            this.label4.Text = "字体大小";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Cmb_FontSize
            // 
            this.Cmb_FontSize.AllowInput = false;
            this.Cmb_FontSize.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_FontSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_FontSize.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_FontSize.FormattingEnabled = true;
            this.Cmb_FontSize.Location = new System.Drawing.Point(96, 57);
            this.Cmb_FontSize.Name = "Cmb_FontSize";
            this.Cmb_FontSize.Size = new System.Drawing.Size(94, 25);
            this.Cmb_FontSize.TabIndex = 5;
            // 
            // Btn_FontColor
            // 
            this.Btn_FontColor.BackColor = System.Drawing.Color.Black;
            this.Btn_FontColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Btn_FontColor.Location = new System.Drawing.Point(191, 57);
            this.Btn_FontColor.Name = "Btn_FontColor";
            this.Btn_FontColor.Size = new System.Drawing.Size(25, 25);
            this.Btn_FontColor.TabIndex = 6;
            this.Btn_FontColor.UseVisualStyleBackColor = false;
            // 
            // Btn_Bold
            // 
            this.Btn_Bold.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_Bold.Location = new System.Drawing.Point(288, 53);
            this.Btn_Bold.Name = "Btn_Bold";
            this.Btn_Bold.Pressed = false;
            this.Btn_Bold.Size = new System.Drawing.Size(55, 30);
            this.Btn_Bold.TabIndex = 7;
            this.Btn_Bold.Text = "粗体";
            this.Btn_Bold.UseVisualStyleBackColor = false;
            // 
            // Btn_Italic
            // 
            this.Btn_Italic.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_Italic.Location = new System.Drawing.Point(353, 54);
            this.Btn_Italic.Name = "Btn_Italic";
            this.Btn_Italic.Pressed = false;
            this.Btn_Italic.Size = new System.Drawing.Size(55, 30);
            this.Btn_Italic.TabIndex = 8;
            this.Btn_Italic.Text = "斜体";
            this.Btn_Italic.UseVisualStyleBackColor = false;
            // 
            // Btn_UnderLine
            // 
            this.Btn_UnderLine.BackColor = System.Drawing.Color.AliceBlue;
            this.Btn_UnderLine.Location = new System.Drawing.Point(418, 54);
            this.Btn_UnderLine.Name = "Btn_UnderLine";
            this.Btn_UnderLine.Pressed = false;
            this.Btn_UnderLine.Size = new System.Drawing.Size(55, 30);
            this.Btn_UnderLine.TabIndex = 9;
            this.Btn_UnderLine.Text = "下划线";
            this.Btn_UnderLine.UseVisualStyleBackColor = false;
            // 
            // Pal_ParaIndent
            // 
            this.Pal_ParaIndent.Controls.Add(this.label9);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_LineSpacing);
            this.Pal_ParaIndent.Controls.Add(this.Nud_LineSpacing);
            this.Pal_ParaIndent.Controls.Add(this.label10);
            this.Pal_ParaIndent.Controls.Add(this.Nud_BefreSpacing);
            this.Pal_ParaIndent.Controls.Add(this.label11);
            this.Pal_ParaIndent.Controls.Add(this.Nud_AfterSpacing);
            this.Pal_ParaIndent.Controls.Add(this.label15);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_ParaAligment);
            this.Pal_ParaIndent.Controls.Add(this.label6);
            this.Pal_ParaIndent.Controls.Add(this.Nud_LeftIndent);
            this.Pal_ParaIndent.Controls.Add(this.label18);
            this.Pal_ParaIndent.Controls.Add(this.Nud_RightIndent);
            this.Pal_ParaIndent.Controls.Add(this.label19);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_FirstLineIndentType);
            this.Pal_ParaIndent.Controls.Add(this.label7);
            this.Pal_ParaIndent.Controls.Add(this.Nud_FirstLineIndent);
            this.Pal_ParaIndent.Controls.Add(this.Nud_FirstLineIndentByChar);
            this.Pal_ParaIndent.Enabled = false;
            this.Pal_ParaIndent.Location = new System.Drawing.Point(182, 120);
            this.Pal_ParaIndent.Name = "Pal_ParaIndent";
            this.Pal_ParaIndent.Size = new System.Drawing.Size(486, 144);
            this.Pal_ParaIndent.TabIndex = 7;
            this.Pal_ParaIndent.TabStop = false;
            this.Pal_ParaIndent.Text = "段落设置";
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(288, 20);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(65, 30);
            this.label9.TabIndex = 15;
            this.label9.Text = "段落行距";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Nud_LineSpacing
            // 
            this.Nud_LineSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_LineSpacing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_LineSpacing.DecimalPlaces = 1;
            this.Nud_LineSpacing.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_LineSpacing.Location = new System.Drawing.Point(354, 24);
            this.Nud_LineSpacing.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_LineSpacing.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_LineSpacing.Name = "Nud_LineSpacing";
            this.Nud_LineSpacing.Size = new System.Drawing.Size(120, 23);
            this.Nud_LineSpacing.TabIndex = 16;
            this.Nud_LineSpacing.Unit = "行";
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(10, 110);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(65, 30);
            this.label10.TabIndex = 17;
            this.label10.Text = "段前间距";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Nud_BefreSpacing
            // 
            this.Nud_BefreSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_BefreSpacing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_BefreSpacing.DecimalPlaces = 1;
            this.Nud_BefreSpacing.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_BefreSpacing.Location = new System.Drawing.Point(96, 114);
            this.Nud_BefreSpacing.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_BefreSpacing.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_BefreSpacing.Name = "Nud_BefreSpacing";
            this.Nud_BefreSpacing.Size = new System.Drawing.Size(120, 23);
            this.Nud_BefreSpacing.TabIndex = 18;
            this.Nud_BefreSpacing.Unit = "磅";
            // 
            // label11
            // 
            this.label11.Location = new System.Drawing.Point(288, 110);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(65, 30);
            this.label11.TabIndex = 19;
            this.label11.Text = "段后间距";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Nud_AfterSpacing
            // 
            this.Nud_AfterSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_AfterSpacing.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_AfterSpacing.DecimalPlaces = 1;
            this.Nud_AfterSpacing.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_AfterSpacing.Location = new System.Drawing.Point(354, 114);
            this.Nud_AfterSpacing.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_AfterSpacing.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_AfterSpacing.Name = "Nud_AfterSpacing";
            this.Nud_AfterSpacing.Size = new System.Drawing.Size(120, 23);
            this.Nud_AfterSpacing.TabIndex = 21;
            this.Nud_AfterSpacing.Unit = "磅";
            // 
            // label15
            // 
            this.label15.Location = new System.Drawing.Point(10, 20);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(80, 30);
            this.label15.TabIndex = 14;
            this.label15.Text = "段落对齐";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Cmb_ParaAligment
            // 
            this.Cmb_ParaAligment.AllowInput = false;
            this.Cmb_ParaAligment.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_ParaAligment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_ParaAligment.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_ParaAligment.FormattingEnabled = true;
            this.Cmb_ParaAligment.Items.AddRange(new object[] {
            "左对齐",
            "居中对齐",
            "右对齐",
            "两端对齐",
            "分散对齐"});
            this.Cmb_ParaAligment.Location = new System.Drawing.Point(96, 23);
            this.Cmb_ParaAligment.Name = "Cmb_ParaAligment";
            this.Cmb_ParaAligment.Size = new System.Drawing.Size(120, 25);
            this.Cmb_ParaAligment.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(10, 50);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 30);
            this.label6.TabIndex = 8;
            this.label6.Text = "左缩进";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Nud_LeftIndent
            // 
            this.Nud_LeftIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_LeftIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_LeftIndent.DecimalPlaces = 1;
            this.Nud_LeftIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_LeftIndent.Location = new System.Drawing.Point(96, 54);
            this.Nud_LeftIndent.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_LeftIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_LeftIndent.Name = "Nud_LeftIndent";
            this.Nud_LeftIndent.Size = new System.Drawing.Size(120, 23);
            this.Nud_LeftIndent.TabIndex = 9;
            this.Nud_LeftIndent.Unit = "厘米";
            // 
            // label18
            // 
            this.label18.Location = new System.Drawing.Point(288, 50);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(50, 30);
            this.label18.TabIndex = 10;
            this.label18.Text = "右缩进";
            this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Nud_RightIndent
            // 
            this.Nud_RightIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_RightIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_RightIndent.DecimalPlaces = 1;
            this.Nud_RightIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_RightIndent.Location = new System.Drawing.Point(354, 54);
            this.Nud_RightIndent.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_RightIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_RightIndent.Name = "Nud_RightIndent";
            this.Nud_RightIndent.Size = new System.Drawing.Size(120, 23);
            this.Nud_RightIndent.TabIndex = 11;
            this.Nud_RightIndent.Unit = "厘米";
            // 
            // label19
            // 
            this.label19.Location = new System.Drawing.Point(10, 80);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(80, 30);
            this.label19.TabIndex = 12;
            this.label19.Text = "首行缩进方式";
            this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Cmb_FirstLineIndentType
            // 
            this.Cmb_FirstLineIndentType.AllowInput = false;
            this.Cmb_FirstLineIndentType.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_FirstLineIndentType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_FirstLineIndentType.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_FirstLineIndentType.FormattingEnabled = true;
            this.Cmb_FirstLineIndentType.Items.AddRange(new object[] {
            "无",
            "悬挂缩进",
            "首行缩进"});
            this.Cmb_FirstLineIndentType.Location = new System.Drawing.Point(96, 83);
            this.Cmb_FirstLineIndentType.Name = "Cmb_FirstLineIndentType";
            this.Cmb_FirstLineIndentType.Size = new System.Drawing.Size(120, 25);
            this.Cmb_FirstLineIndentType.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(288, 80);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 30);
            this.label7.TabIndex = 14;
            this.label7.Text = "首行缩进";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Nud_FirstLineIndent
            // 
            this.Nud_FirstLineIndent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_FirstLineIndent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_FirstLineIndent.DecimalPlaces = 1;
            this.Nud_FirstLineIndent.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_FirstLineIndent.Location = new System.Drawing.Point(354, 83);
            this.Nud_FirstLineIndent.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_FirstLineIndent.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_FirstLineIndent.Name = "Nud_FirstLineIndent";
            this.Nud_FirstLineIndent.Size = new System.Drawing.Size(120, 23);
            this.Nud_FirstLineIndent.TabIndex = 15;
            this.Nud_FirstLineIndent.Unit = "厘米";
            // 
            // Nud_FirstLineIndentByChar
            // 
            this.Nud_FirstLineIndentByChar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Nud_FirstLineIndentByChar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Nud_FirstLineIndentByChar.DecimalPlaces = 1;
            this.Nud_FirstLineIndentByChar.Increment = new decimal(new int[] {
            1,
            0,
            0,
            131072});
            this.Nud_FirstLineIndentByChar.Location = new System.Drawing.Point(96, 113);
            this.Nud_FirstLineIndentByChar.Maximum = new decimal(new int[] {
            -1,
            -1,
            -1,
            0});
            this.Nud_FirstLineIndentByChar.Minimum = new decimal(new int[] {
            -1,
            -1,
            -1,
            -2147483648});
            this.Nud_FirstLineIndentByChar.Name = "Nud_FirstLineIndentByChar";
            this.Nud_FirstLineIndentByChar.Size = new System.Drawing.Size(100, 23);
            this.Nud_FirstLineIndentByChar.TabIndex = 17;
            this.Nud_FirstLineIndentByChar.Unit = "字符";
            // 
            // Btn_ReadDocumentStyle
            // 
            this.Btn_ReadDocumentStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.Btn_ReadDocumentStyle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Btn_ReadDocumentStyle.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.Btn_ReadDocumentStyle.ForeColor = System.Drawing.Color.Black;
            this.Btn_ReadDocumentStyle.Location = new System.Drawing.Point(278, 334);
            this.Btn_ReadDocumentStyle.Name = "Btn_ReadDocumentStyle";
            this.Btn_ReadDocumentStyle.Size = new System.Drawing.Size(120, 31);
            this.Btn_ReadDocumentStyle.TabIndex = 32;
            this.Btn_ReadDocumentStyle.Text = "读取文中样式";
            this.Btn_ReadDocumentStyle.UseVisualStyleBackColor = false;
            // 
            // Btn_AddStyle
            // 
            this.Btn_AddStyle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.Btn_AddStyle.Enabled = false;
            this.Btn_AddStyle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Btn_AddStyle.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.Btn_AddStyle.ForeColor = System.Drawing.Color.Black;
            this.Btn_AddStyle.Location = new System.Drawing.Point(14, 334);
            this.Btn_AddStyle.Name = "Btn_AddStyle";
            this.Btn_AddStyle.Size = new System.Drawing.Size(72, 31);
            this.Btn_AddStyle.TabIndex = 27;
            this.Btn_AddStyle.Text = "添加样式";
            this.Btn_AddStyle.UseVisualStyleBackColor = false;
            // 
            // Lst_Styles
            // 
            this.Lst_Styles.FormattingEnabled = true;
            this.Lst_Styles.ItemHeight = 17;
            this.Lst_Styles.Location = new System.Drawing.Point(14, 23);
            this.Lst_Styles.Name = "Lst_Styles";
            this.Lst_Styles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.Lst_Styles.Size = new System.Drawing.Size(158, 276);
            this.Lst_Styles.TabIndex = 0;
            // 
            // Cmb_SetLevel
            // 
            this.Cmb_SetLevel.AllowInput = false;
            this.Cmb_SetLevel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_SetLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_SetLevel.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_SetLevel.FormattingEnabled = true;
            this.Cmb_SetLevel.Items.AddRange(new object[] {
            "无",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9"});
            this.Cmb_SetLevel.Location = new System.Drawing.Point(535, 274);
            this.Cmb_SetLevel.Name = "Cmb_SetLevel";
            this.Cmb_SetLevel.Size = new System.Drawing.Size(120, 25);
            this.Cmb_SetLevel.TabIndex = 21;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(470, 278);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(68, 17);
            this.label12.TabIndex = 20;
            this.label12.Text = "显示标题数";
            // 
            // Btn_ApplySet
            // 
            this.Btn_ApplySet.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.Btn_ApplySet.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Btn_ApplySet.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.Btn_ApplySet.ForeColor = System.Drawing.Color.Black;
            this.Btn_ApplySet.Location = new System.Drawing.Point(535, 334);
            this.Btn_ApplySet.Name = "Btn_ApplySet";
            this.Btn_ApplySet.Size = new System.Drawing.Size(120, 31);
            this.Btn_ApplySet.TabIndex = 7;
            this.Btn_ApplySet.Text = "应用设置";
            this.Btn_ApplySet.UseVisualStyleBackColor = false;
            // 
            // Cmb_PreSettings
            // 
            this.Cmb_PreSettings.AllowInput = false;
            this.Cmb_PreSettings.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_PreSettings.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_PreSettings.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_PreSettings.FormattingEnabled = true;
            this.Cmb_PreSettings.Items.AddRange(new object[] {
            "公文风格",
            "论文风格",
            "报告风格",
            "条文风格"});
            this.Cmb_PreSettings.Location = new System.Drawing.Point(278, 274);
            this.Cmb_PreSettings.Name = "Cmb_PreSettings";
            this.Cmb_PreSettings.Size = new System.Drawing.Size(120, 25);
            this.Cmb_PreSettings.TabIndex = 8;
            // 
            // Cmb_LineSpacing
            // 
            this.Cmb_LineSpacing.AllowInput = false;
            this.Cmb_LineSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_LineSpacing.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cmb_LineSpacing.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_LineSpacing.FormattingEnabled = true;
            this.Cmb_LineSpacing.Items.AddRange(new object[] {
            "单倍行距",
            "1.5倍行距",
            "2倍行距",
            "最小值",
            "固定值",
            "多倍行距"});
            this.Cmb_LineSpacing.Location = new System.Drawing.Point(354, 24);
            this.Cmb_LineSpacing.Name = "Cmb_LineSpacing";
            this.Cmb_LineSpacing.Size = new System.Drawing.Size(120, 25);
            this.Cmb_LineSpacing.TabIndex = 16;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(192, 278);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(56, 17);
            this.label16.TabIndex = 9;
            this.label16.Text = "预设风格";
            this.label16.Click += new System.EventHandler(this.label16_Click);
            // 
            // StyleSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(678, 374);
            this.Controls.Add(this.groupBox3);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MinimumSize = new System.Drawing.Size(400, 300);
            this.Name = "StyleSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "样式设置";
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.Pal_Font.ResumeLayout(false);
            this.Pal_ParaIndent.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Nud_LineSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_BefreSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_AfterSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_LeftIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_RightIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_FirstLineIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_FirstLineIndentByChar)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private StandardButton Btn_ReadDocumentStyle;
    }
}
