using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WordMan;
using WordMan.MultiLevel;

namespace WordMan
{
    partial class StyleSettings
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region 控件声明
        private GroupBox groupBox3;
        private StandardButton 删除;
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
        private StandardComboBox 首行缩进方式下拉框;
        private StandardButton 添加;
        private StandardComboBox 显示标题数下拉框;
        private Label label12;
        private ListBox Lst_Styles;
        private StandardButton Btn_ApplySet;
        private StandardComboBox 风格下拉框;
        private Label label16;
        private StandardTextBox Txt_AddStyleName;
        private StandardComboBox Cmb_LineSpacing;
        private StandardComboBox Cmb_BefreSpacing;
        private StandardComboBox Cmb_AfterSpacing;
        private StandardNumericUpDown Nud_LeftIndent;
        private StandardNumericUpDown Nud_FirstLineIndent;
        private StandardButton 关闭;

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
            this.label12 = new System.Windows.Forms.Label();
            this.Pal_Font = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.Pal_ParaIndent = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.Lst_Styles = new System.Windows.Forms.ListBox();
            this.导入 = new WordMan.StandardButton();
            this.加载 = new WordMan.StandardButton();
            this.关闭 = new WordMan.StandardButton();
            this.Txt_AddStyleName = new WordMan.StandardTextBox();
            this.Btn_ApplySet = new WordMan.StandardButton();
            this.删除 = new WordMan.StandardButton();
            this.显示标题数下拉框 = new WordMan.StandardComboBox();
            this.风格下拉框 = new WordMan.StandardComboBox();
            this.Cmb_ChnFontName = new WordMan.StandardComboBox();
            this.Cmb_EngFontName = new WordMan.StandardComboBox();
            this.Cmb_FontSize = new WordMan.StandardComboBox();
            this.Btn_FontColor = new WordMan.StandardButton();
            this.Btn_Bold = new WordMan.ToggleButton();
            this.Btn_Italic = new WordMan.ToggleButton();
            this.Btn_UnderLine = new WordMan.ToggleButton();
            this.Cmb_LineSpacing = new WordMan.StandardComboBox();
            this.Cmb_BefreSpacing = new WordMan.StandardComboBox();
            this.Cmb_AfterSpacing = new WordMan.StandardComboBox();
            this.Cmb_ParaAligment = new WordMan.StandardComboBox();
            this.Nud_LeftIndent = new WordMan.StandardNumericUpDown();
            this.Nud_RightIndent = new WordMan.StandardNumericUpDown();
            this.首行缩进方式下拉框 = new WordMan.StandardComboBox();
            this.Nud_FirstLineIndent = new WordMan.StandardNumericUpDown();
            this.导出 = new WordMan.StandardButton();
            this.添加 = new WordMan.StandardButton();
            this.读取文档样式 = new WordMan.StandardButton();
            this.groupBox3.SuspendLayout();
            this.Pal_Font.SuspendLayout();
            this.Pal_ParaIndent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_LeftIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_RightIndent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_FirstLineIndent)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.读取文档样式);
            this.groupBox3.Controls.Add(this.导入);
            this.groupBox3.Controls.Add(this.加载);
            this.groupBox3.Controls.Add(this.关闭);
            this.groupBox3.Controls.Add(this.Txt_AddStyleName);
            this.groupBox3.Controls.Add(this.Btn_ApplySet);
            this.groupBox3.Controls.Add(this.删除);
            this.groupBox3.Controls.Add(this.显示标题数下拉框);
            this.groupBox3.Controls.Add(this.风格下拉框);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.Pal_Font);
            this.groupBox3.Controls.Add(this.label16);
            this.groupBox3.Controls.Add(this.Pal_ParaIndent);
            this.groupBox3.Controls.Add(this.导出);
            this.groupBox3.Controls.Add(this.添加);
            this.groupBox3.Controls.Add(this.Lst_Styles);
            this.groupBox3.Location = new System.Drawing.Point(0, 0);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(679, 374);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "样式设置";
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
            // label17
            // 
            this.label17.Location = new System.Drawing.Point(288, 20);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(60, 28);
            this.label17.TabIndex = 2;
            this.label17.Text = "西文字体";
            this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            // Pal_ParaIndent
            // 
            this.Pal_ParaIndent.Controls.Add(this.label9);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_LineSpacing);
            this.Pal_ParaIndent.Controls.Add(this.label10);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_BefreSpacing);
            this.Pal_ParaIndent.Controls.Add(this.label11);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_AfterSpacing);
            this.Pal_ParaIndent.Controls.Add(this.label15);
            this.Pal_ParaIndent.Controls.Add(this.Cmb_ParaAligment);
            this.Pal_ParaIndent.Controls.Add(this.label6);
            this.Pal_ParaIndent.Controls.Add(this.Nud_LeftIndent);
            this.Pal_ParaIndent.Controls.Add(this.label18);
            this.Pal_ParaIndent.Controls.Add(this.Nud_RightIndent);
            this.Pal_ParaIndent.Controls.Add(this.label19);
            this.Pal_ParaIndent.Controls.Add(this.首行缩进方式下拉框);
            this.Pal_ParaIndent.Controls.Add(this.label7);
            this.Pal_ParaIndent.Controls.Add(this.Nud_FirstLineIndent);
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
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(10, 110);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(65, 30);
            this.label10.TabIndex = 17;
            this.label10.Text = "段前间距";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            // label15
            // 
            this.label15.Location = new System.Drawing.Point(10, 20);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(80, 30);
            this.label15.TabIndex = 14;
            this.label15.Text = "段落对齐";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            // label18
            // 
            this.label18.Location = new System.Drawing.Point(288, 50);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(50, 30);
            this.label18.TabIndex = 10;
            this.label18.Text = "右缩进";
            this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(288, 80);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 30);
            this.label7.TabIndex = 14;
            this.label7.Text = "首行缩进";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            // 导入
            // 
            this.导入.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.导入.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.导入.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.导入.ForeColor = System.Drawing.Color.Black;
            this.导入.Location = new System.Drawing.Point(278, 334);
            this.导入.Name = "导入";
            this.导入.Size = new System.Drawing.Size(50, 31);
            this.导入.TabIndex = 40;
            this.导入.Text = "导入";
            this.导入.UseVisualStyleBackColor = false;
            // 
            // 加载
            // 
            this.加载.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.加载.Enabled = false;
            this.加载.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.加载.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.加载.ForeColor = System.Drawing.Color.Black;
            this.加载.Location = new System.Drawing.Point(14, 334);
            this.加载.Name = "加载";
            this.加载.Size = new System.Drawing.Size(50, 31);
            this.加载.TabIndex = 35;
            this.加载.Text = "加载";
            this.加载.UseVisualStyleBackColor = false;
            // 
            // 关闭
            // 
            this.关闭.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.关闭.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.关闭.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.关闭.ForeColor = System.Drawing.Color.Black;
            this.关闭.Location = new System.Drawing.Point(536, 334);
            this.关闭.Name = "关闭";
            this.关闭.Size = new System.Drawing.Size(120, 31);
            this.关闭.TabIndex = 34;
            this.关闭.Text = "关闭";
            this.关闭.UseVisualStyleBackColor = false;
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
            // Btn_ApplySet
            // 
            this.Btn_ApplySet.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.Btn_ApplySet.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Btn_ApplySet.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.Btn_ApplySet.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.Btn_ApplySet.Location = new System.Drawing.Point(407, 334);
            this.Btn_ApplySet.Name = "Btn_ApplySet";
            this.Btn_ApplySet.Size = new System.Drawing.Size(120, 31);
            this.Btn_ApplySet.TabIndex = 7;
            this.Btn_ApplySet.Text = "应用设置";
            this.Btn_ApplySet.UseVisualStyleBackColor = false;
            // 
            // 删除
            // 
            this.删除.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.删除.Enabled = false;
            this.删除.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.删除.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.删除.ForeColor = System.Drawing.Color.Black;
            this.删除.Location = new System.Drawing.Point(122, 334);
            this.删除.Name = "删除";
            this.删除.Size = new System.Drawing.Size(50, 31);
            this.删除.TabIndex = 29;
            this.删除.Text = "删除";
            this.删除.UseVisualStyleBackColor = false;
            // 
            // 显示标题数下拉框
            // 
            this.显示标题数下拉框.AllowInput = false;
            this.显示标题数下拉框.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.显示标题数下拉框.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.显示标题数下拉框.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.显示标题数下拉框.FormattingEnabled = true;
            this.显示标题数下拉框.Items.AddRange(new object[] {
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
            this.显示标题数下拉框.Location = new System.Drawing.Point(535, 274);
            this.显示标题数下拉框.Name = "显示标题数下拉框";
            this.显示标题数下拉框.Size = new System.Drawing.Size(120, 25);
            this.显示标题数下拉框.TabIndex = 21;
            // 
            // 风格下拉框
            // 
            this.风格下拉框.AllowInput = false;
            this.风格下拉框.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.风格下拉框.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.风格下拉框.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.风格下拉框.FormattingEnabled = true;
            this.风格下拉框.Items.AddRange(new object[] {
            "公文风格",
            "论文风格",
            "报告风格",
            "条文风格"});
            this.风格下拉框.Location = new System.Drawing.Point(278, 274);
            this.风格下拉框.Name = "风格下拉框";
            this.风格下拉框.Size = new System.Drawing.Size(120, 25);
            this.风格下拉框.TabIndex = 8;
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
            this.Btn_FontColor.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.Btn_FontColor.ForeColor = System.Drawing.Color.Black;
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
            this.Btn_Italic.Location = new System.Drawing.Point(353, 53);
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
            this.Btn_UnderLine.Location = new System.Drawing.Point(418, 53);
            this.Btn_UnderLine.Name = "Btn_UnderLine";
            this.Btn_UnderLine.Pressed = false;
            this.Btn_UnderLine.Size = new System.Drawing.Size(55, 30);
            this.Btn_UnderLine.TabIndex = 9;
            this.Btn_UnderLine.Text = "下划线";
            this.Btn_UnderLine.UseVisualStyleBackColor = false;
            // 
            // Cmb_LineSpacing
            // 
            this.Cmb_LineSpacing.AllowInput = true;
            this.Cmb_LineSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
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
            // Cmb_BefreSpacing
            // 
            this.Cmb_BefreSpacing.AllowInput = true;
            this.Cmb_BefreSpacing.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.Cmb_BefreSpacing.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.Cmb_BefreSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_BefreSpacing.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            this.Cmb_BefreSpacing.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_BefreSpacing.FormattingEnabled = true;
            this.Cmb_BefreSpacing.Location = new System.Drawing.Point(96, 114);
            this.Cmb_BefreSpacing.Name = "Cmb_BefreSpacing";
            this.Cmb_BefreSpacing.Size = new System.Drawing.Size(120, 25);
            this.Cmb_BefreSpacing.TabIndex = 18;
            // 
            // Cmb_AfterSpacing
            // 
            this.Cmb_AfterSpacing.AllowInput = true;
            this.Cmb_AfterSpacing.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.Cmb_AfterSpacing.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.Cmb_AfterSpacing.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.Cmb_AfterSpacing.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            this.Cmb_AfterSpacing.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.Cmb_AfterSpacing.FormattingEnabled = true;
            this.Cmb_AfterSpacing.Location = new System.Drawing.Point(354, 114);
            this.Cmb_AfterSpacing.Name = "Cmb_AfterSpacing";
            this.Cmb_AfterSpacing.Size = new System.Drawing.Size(120, 25);
            this.Cmb_AfterSpacing.TabIndex = 21;
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
            // 首行缩进方式下拉框
            // 
            this.首行缩进方式下拉框.AllowInput = false;
            this.首行缩进方式下拉框.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.首行缩进方式下拉框.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.首行缩进方式下拉框.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.首行缩进方式下拉框.FormattingEnabled = true;
            this.首行缩进方式下拉框.Items.AddRange(new object[] {
            "无",
            "悬挂缩进",
            "首行缩进"});
            this.首行缩进方式下拉框.Location = new System.Drawing.Point(96, 83);
            this.首行缩进方式下拉框.Name = "首行缩进方式下拉框";
            this.首行缩进方式下拉框.Size = new System.Drawing.Size(120, 25);
            this.首行缩进方式下拉框.TabIndex = 13;
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
            // 导出
            // 
            this.导出.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.导出.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.导出.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.导出.ForeColor = System.Drawing.Color.Black;
            this.导出.Location = new System.Drawing.Point(348, 334);
            this.导出.Name = "导出";
            this.导出.Size = new System.Drawing.Size(50, 31);
            this.导出.TabIndex = 32;
            this.导出.Text = "导出";
            this.导出.UseVisualStyleBackColor = false;
            // 
            // 添加
            // 
            this.添加.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.添加.Enabled = false;
            this.添加.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.添加.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.添加.ForeColor = System.Drawing.Color.Black;
            this.添加.Location = new System.Drawing.Point(68, 334);
            this.添加.Name = "添加";
            this.添加.Size = new System.Drawing.Size(50, 31);
            this.添加.TabIndex = 27;
            this.添加.Text = "添加";
            this.添加.UseVisualStyleBackColor = false;
            // 
            // 读取文档样式
            // 
            this.读取文档样式.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(245)))), ((int)(((byte)(245)))));
            this.读取文档样式.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.读取文档样式.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.读取文档样式.ForeColor = System.Drawing.Color.Black;
            this.读取文档样式.Location = new System.Drawing.Point(180, 334);
            this.读取文档样式.Name = "读取文档样式";
            this.读取文档样式.Size = new System.Drawing.Size(90, 31);
            this.读取文档样式.TabIndex = 41;
            this.读取文档样式.Text = "读取文档样式";
            this.读取文档样式.UseVisualStyleBackColor = false;
            // 
            // StyleSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(219)))), ((int)(((byte)(233)))), ((int)(((byte)(247)))));
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
            ((System.ComponentModel.ISupportInitialize)(this.Nud_LeftIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_RightIndent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Nud_FirstLineIndent)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private StandardButton 导出;
        private StandardButton 加载;
        private StandardButton 导入;
        private StandardButton 读取文档样式;
    }
}
