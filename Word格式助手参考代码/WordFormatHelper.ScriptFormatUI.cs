// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.ScriptFormatUI
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordFormatHelper;

public class ScriptFormatUI : UserControl
{
	private IContainer components;

	private Button Btn_Apply;

	private Label label1;

	private TextBox Txt_ScriptText;

	private RadioButton Rdo_Superscript;

	private RadioButton Rdo_Subscript;

	private Label label2;

	private CheckBox Chk_UseRegex;

	private Button Btn_RegularExpress;

	private TextBox textBox1;

	public ScriptFormatUI()
	{
		InitializeComponent();
		Btn_RegularExpress.Text = "▲" + Btn_RegularExpress.Text;
	}

	private void Btn_Apply_Click(object sender, EventArgs e)
	{
		if (!string.IsNullOrEmpty(Txt_ScriptText.Text) && Globals.ThisAddIn.Application.Selection.Type == WdSelectionType.wdSelectionNormal)
		{
			ThisAddIn.SetSuperscriptOrSubscript(Globals.ThisAddIn.Application.Selection.Range, Txt_ScriptText.Text, Rdo_Superscript.Checked, Chk_UseRegex.Checked);
		}
	}

	private void Btn_RegularExpress_Click(object sender, EventArgs e)
	{
		if (Btn_RegularExpress.Text.StartsWith("▼"))
		{
			Btn_RegularExpress.Text = Btn_RegularExpress.Text.Replace("▼", "▲");
			(base.Parent as Form).ClientSize = new Size(base.Width, 300);
		}
		else
		{
			Btn_RegularExpress.Text = Btn_RegularExpress.Text.Replace("▲", "▼");
			(base.Parent as Form).ClientSize = new Size(base.Width, 140);
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	private void InitializeComponent()
	{
		this.Btn_Apply = new System.Windows.Forms.Button();
		this.label1 = new System.Windows.Forms.Label();
		this.Txt_ScriptText = new System.Windows.Forms.TextBox();
		this.Rdo_Superscript = new System.Windows.Forms.RadioButton();
		this.Rdo_Subscript = new System.Windows.Forms.RadioButton();
		this.label2 = new System.Windows.Forms.Label();
		this.Chk_UseRegex = new System.Windows.Forms.CheckBox();
		this.Btn_RegularExpress = new System.Windows.Forms.Button();
		this.textBox1 = new System.Windows.Forms.TextBox();
		base.SuspendLayout();
		this.Btn_Apply.Location = new System.Drawing.Point(220, 77);
		this.Btn_Apply.Name = "Btn_Apply";
		this.Btn_Apply.Size = new System.Drawing.Size(120, 30);
		this.Btn_Apply.TabIndex = 0;
		this.Btn_Apply.Text = "应用";
		this.Btn_Apply.UseVisualStyleBackColor = true;
		this.Btn_Apply.Click += new System.EventHandler(Btn_Apply_Click);
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(10, 14);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(78, 18);
		this.label1.TabIndex = 1;
		this.label1.Text = "文字内容：";
		this.Txt_ScriptText.Location = new System.Drawing.Point(99, 11);
		this.Txt_ScriptText.Name = "Txt_ScriptText";
		this.Txt_ScriptText.Size = new System.Drawing.Size(243, 25);
		this.Txt_ScriptText.TabIndex = 2;
		this.Rdo_Superscript.AutoSize = true;
		this.Rdo_Superscript.Checked = true;
		this.Rdo_Superscript.Location = new System.Drawing.Point(104, 46);
		this.Rdo_Superscript.Name = "Rdo_Superscript";
		this.Rdo_Superscript.Size = new System.Drawing.Size(54, 22);
		this.Rdo_Superscript.TabIndex = 3;
		this.Rdo_Superscript.TabStop = true;
		this.Rdo_Superscript.Text = "上标";
		this.Rdo_Superscript.UseVisualStyleBackColor = true;
		this.Rdo_Subscript.AutoSize = true;
		this.Rdo_Subscript.Location = new System.Drawing.Point(188, 46);
		this.Rdo_Subscript.Name = "Rdo_Subscript";
		this.Rdo_Subscript.Size = new System.Drawing.Size(54, 22);
		this.Rdo_Subscript.TabIndex = 4;
		this.Rdo_Subscript.Text = "下标";
		this.Rdo_Subscript.UseVisualStyleBackColor = true;
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(10, 48);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(64, 18);
		this.label2.TabIndex = 5;
		this.label2.Text = "设定为：";
		this.Chk_UseRegex.AutoSize = true;
		this.Chk_UseRegex.Location = new System.Drawing.Point(14, 82);
		this.Chk_UseRegex.Name = "Chk_UseRegex";
		this.Chk_UseRegex.Size = new System.Drawing.Size(125, 22);
		this.Chk_UseRegex.TabIndex = 6;
		this.Chk_UseRegex.Text = "使用正则表达式";
		this.Chk_UseRegex.UseVisualStyleBackColor = true;
		this.Btn_RegularExpress.FlatAppearance.BorderSize = 0;
		this.Btn_RegularExpress.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
		this.Btn_RegularExpress.Location = new System.Drawing.Point(3, 110);
		this.Btn_RegularExpress.Name = "Btn_RegularExpress";
		this.Btn_RegularExpress.Size = new System.Drawing.Size(344, 25);
		this.Btn_RegularExpress.TabIndex = 7;
		this.Btn_RegularExpress.Text = "常用正则匹配";
		this.Btn_RegularExpress.UseVisualStyleBackColor = true;
		this.Btn_RegularExpress.Click += new System.EventHandler(Btn_RegularExpress_Click);
		this.textBox1.Location = new System.Drawing.Point(3, 141);
		this.textBox1.Multiline = true;
		this.textBox1.Name = "textBox1";
		this.textBox1.ReadOnly = true;
		this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
		this.textBox1.Size = new System.Drawing.Size(344, 156);
		this.textBox1.TabIndex = 8;
		this.textBox1.Text = "1.所有数字：[0-9] 或 \\d\r\n2.所有大写字母：[A-Z]\r\n3.所有小写字母：[a-z]\r\n4.所有汉字：[\\u4E00-\\u9FA5]\r\n5.所有希腊字母：[\\u0370-\\u03FF\\u1F00-\\u1FFF]\r\n6.面积单位平方：(?<=m)2\r\n7.带圈或括号数字：[\\u2460-\\u24FF]\r\n8.罗马数字：[\\u2160-\\u217F]";
		base.AutoScaleDimensions = new System.Drawing.SizeF(96f, 96f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.textBox1);
		base.Controls.Add(this.Btn_RegularExpress);
		base.Controls.Add(this.Chk_UseRegex);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.Rdo_Subscript);
		base.Controls.Add(this.Rdo_Superscript);
		base.Controls.Add(this.Txt_ScriptText);
		base.Controls.Add(this.label1);
		base.Controls.Add(this.Btn_Apply);
		this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		base.Name = "ScriptFormatUI";
		base.Size = new System.Drawing.Size(350, 300);
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
