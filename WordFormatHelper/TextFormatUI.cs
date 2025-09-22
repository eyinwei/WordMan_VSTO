using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace WordFormatHelper{

public class TextFormatUI : UserControl
{
	private IContainer components;

	private GroupBox groupBox1;

	private CheckBox Chk_CharSet3;

	private CheckBox Chk_CharSet2;

	private CheckBox Chk_CharSet1;

	private GroupBox groupBox2;

	private RadioButton Rdo_AllSpace;

	private RadioButton Rdo_PunctuationSpace;

	private GroupBox groupBox3;

	private CheckBox Chk_RemoveCurrentBracket;

	private CheckBox Chk_FullWidthBracket;

	public TextFormatUI()
	{
		InitializeComponent();
		Chk_CharSet1.Checked = ThisAddIn.textFormatSet.SetA;
		Chk_CharSet2.Checked = ThisAddIn.textFormatSet.SetB;
		Chk_CharSet3.Checked = ThisAddIn.textFormatSet.SetC;
		if (ThisAddIn.textFormatSet.PunctuationSpace)
		{
			Rdo_PunctuationSpace.Checked = true;
		}
		else
		{
			Rdo_AllSpace.Checked = true;
		}
		Chk_FullWidthBracket.Checked = ThisAddIn.textFormatSet.FullWidthBracket;
		Chk_RemoveCurrentBracket.Checked = ThisAddIn.textFormatSet.RemoveBrackets;
	}

	private void Rdo_AllSpace_CheckedChanged(object sender, EventArgs e)
	{
		if (Rdo_AllSpace.Checked)
		{
			ThisAddIn.textFormatSet.PunctuationSpace = false;
		}
		else
		{
			ThisAddIn.textFormatSet.PunctuationSpace = true;
		}
	}

	private void Chk_CharSet1_CheckedChanged(object sender, EventArgs e)
	{
		switch ((sender as CheckBox).Name)
		{
		case "Chk_CharSet1":
			ThisAddIn.textFormatSet.SetA = (sender as CheckBox).Checked;
			break;
		case "Chk_CharSet2":
			ThisAddIn.textFormatSet.SetB = (sender as CheckBox).Checked;
			break;
		case "Chk_CharSet3":
			ThisAddIn.textFormatSet.SetC = (sender as CheckBox).Checked;
			break;
		case "Chk_FullWidthBracket":
			ThisAddIn.textFormatSet.FullWidthBracket = (sender as CheckBox).Checked;
			break;
		case "Chk_RemoveCurrentBracket":
			ThisAddIn.textFormatSet.RemoveBrackets = (sender as CheckBox).Checked;
			break;
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
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.Chk_CharSet3 = new System.Windows.Forms.CheckBox();
		this.Chk_CharSet2 = new System.Windows.Forms.CheckBox();
		this.Chk_CharSet1 = new System.Windows.Forms.CheckBox();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.Rdo_PunctuationSpace = new System.Windows.Forms.RadioButton();
		this.Rdo_AllSpace = new System.Windows.Forms.RadioButton();
		this.groupBox3 = new System.Windows.Forms.GroupBox();
		this.Chk_FullWidthBracket = new System.Windows.Forms.CheckBox();
		this.Chk_RemoveCurrentBracket = new System.Windows.Forms.CheckBox();
		this.groupBox1.SuspendLayout();
		this.groupBox2.SuspendLayout();
		this.groupBox3.SuspendLayout();
		base.SuspendLayout();
		this.groupBox1.Controls.Add(this.Chk_CharSet3);
		this.groupBox1.Controls.Add(this.Chk_CharSet2);
		this.groupBox1.Controls.Add(this.Chk_CharSet1);
		this.groupBox1.Location = new System.Drawing.Point(4, 5);
		this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.groupBox1.Size = new System.Drawing.Size(250, 100);
		this.groupBox1.TabIndex = 0;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "标点符号转化";
		this.Chk_CharSet3.AutoSize = true;
		this.Chk_CharSet3.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Chk_CharSet3.Location = new System.Drawing.Point(20, 66);
		this.Chk_CharSet3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Chk_CharSet3.Name = "Chk_CharSet3";
		this.Chk_CharSet3.Size = new System.Drawing.Size(57, 24);
		this.Chk_CharSet3.TabIndex = 2;
		this.Chk_CharSet3.Text = "? ! ~";
		this.Chk_CharSet3.UseVisualStyleBackColor = true;
		this.Chk_CharSet3.CheckedChanged += new System.EventHandler(Chk_CharSet1_CheckedChanged);
		this.Chk_CharSet2.AutoSize = true;
		this.Chk_CharSet2.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Chk_CharSet2.Location = new System.Drawing.Point(116, 33);
		this.Chk_CharSet2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Chk_CharSet2.Name = "Chk_CharSet2";
		this.Chk_CharSet2.Size = new System.Drawing.Size(90, 24);
		this.Chk_CharSet2.TabIndex = 1;
		this.Chk_CharSet2.Text = "() [] {} <>";
		this.Chk_CharSet2.UseVisualStyleBackColor = true;
		this.Chk_CharSet2.CheckedChanged += new System.EventHandler(Chk_CharSet1_CheckedChanged);
		this.Chk_CharSet1.AutoSize = true;
		this.Chk_CharSet1.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Chk_CharSet1.Location = new System.Drawing.Point(20, 33);
		this.Chk_CharSet1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Chk_CharSet1.Name = "Chk_CharSet1";
		this.Chk_CharSet1.Size = new System.Drawing.Size(84, 24);
		this.Chk_CharSet1.TabIndex = 0;
		this.Chk_CharSet1.Text = "，。：；";
		this.Chk_CharSet1.UseVisualStyleBackColor = true;
		this.Chk_CharSet1.CheckedChanged += new System.EventHandler(Chk_CharSet1_CheckedChanged);
		this.groupBox2.Controls.Add(this.Rdo_PunctuationSpace);
		this.groupBox2.Controls.Add(this.Rdo_AllSpace);
		this.groupBox2.Location = new System.Drawing.Point(4, 113);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.Size = new System.Drawing.Size(250, 88);
		this.groupBox2.TabIndex = 1;
		this.groupBox2.TabStop = false;
		this.groupBox2.Text = "删空格";
		this.Rdo_PunctuationSpace.AutoSize = true;
		this.Rdo_PunctuationSpace.Location = new System.Drawing.Point(21, 55);
		this.Rdo_PunctuationSpace.Name = "Rdo_PunctuationSpace";
		this.Rdo_PunctuationSpace.Size = new System.Drawing.Size(209, 24);
		this.Rdo_PunctuationSpace.TabIndex = 1;
		this.Rdo_PunctuationSpace.Text = "仅标点，。：；、！？后空格";
		this.Rdo_PunctuationSpace.UseVisualStyleBackColor = true;
		this.Rdo_PunctuationSpace.CheckedChanged += new System.EventHandler(Rdo_AllSpace_CheckedChanged);
		this.Rdo_AllSpace.AutoSize = true;
		this.Rdo_AllSpace.Checked = true;
		this.Rdo_AllSpace.Location = new System.Drawing.Point(21, 25);
		this.Rdo_AllSpace.Name = "Rdo_AllSpace";
		this.Rdo_AllSpace.Size = new System.Drawing.Size(83, 24);
		this.Rdo_AllSpace.TabIndex = 0;
		this.Rdo_AllSpace.TabStop = true;
		this.Rdo_AllSpace.Text = "所有空格";
		this.Rdo_AllSpace.UseVisualStyleBackColor = true;
		this.Rdo_AllSpace.CheckedChanged += new System.EventHandler(Rdo_AllSpace_CheckedChanged);
		this.groupBox3.Controls.Add(this.Chk_RemoveCurrentBracket);
		this.groupBox3.Controls.Add(this.Chk_FullWidthBracket);
		this.groupBox3.Location = new System.Drawing.Point(4, 207);
		this.groupBox3.Name = "groupBox3";
		this.groupBox3.Size = new System.Drawing.Size(250, 93);
		this.groupBox3.TabIndex = 2;
		this.groupBox3.TabStop = false;
		this.groupBox3.Text = "括号";
		this.Chk_FullWidthBracket.AutoSize = true;
		this.Chk_FullWidthBracket.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Chk_FullWidthBracket.Location = new System.Drawing.Point(20, 27);
		this.Chk_FullWidthBracket.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Chk_FullWidthBracket.Name = "Chk_FullWidthBracket";
		this.Chk_FullWidthBracket.Size = new System.Drawing.Size(174, 24);
		this.Chk_FullWidthBracket.TabIndex = 2;
		this.Chk_FullWidthBracket.Text = "() [] {} <>使用全角括号";
		this.Chk_FullWidthBracket.UseVisualStyleBackColor = true;
		this.Chk_FullWidthBracket.CheckedChanged += new System.EventHandler(Chk_CharSet1_CheckedChanged);
		this.Chk_RemoveCurrentBracket.AutoSize = true;
		this.Chk_RemoveCurrentBracket.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Chk_RemoveCurrentBracket.Location = new System.Drawing.Point(20, 61);
		this.Chk_RemoveCurrentBracket.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Chk_RemoveCurrentBracket.Name = "Chk_RemoveCurrentBracket";
		this.Chk_RemoveCurrentBracket.Size = new System.Drawing.Size(112, 24);
		this.Chk_RemoveCurrentBracket.TabIndex = 3;
		this.Chk_RemoveCurrentBracket.Text = "替换现有括号";
		this.Chk_RemoveCurrentBracket.UseVisualStyleBackColor = true;
		this.Chk_RemoveCurrentBracket.CheckedChanged += new System.EventHandler(Chk_CharSet1_CheckedChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.groupBox3);
		base.Controls.Add(this.groupBox2);
		base.Controls.Add(this.groupBox1);
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		base.Name = "TextFormatUI";
		base.Size = new System.Drawing.Size(261, 306);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		this.groupBox3.ResumeLayout(false);
		this.groupBox3.PerformLayout();
		base.ResumeLayout(false);
	}
}
}