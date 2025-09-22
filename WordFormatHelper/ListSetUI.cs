using System;
using System.ComponentModel;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class ListSetUI : UserControl
{
	private IContainer components;

	private NumericUpDownWithUnit NumUpDownNumberIndent;

	private Label label8;

	private RadioButton RdoDefaultListFormat;

	private RadioButton RdoUseNewListFormat;

	private NumericUpDownWithUnit NumUpDownTextIndent;

	private Label label1;

	private NumericUpDownWithUnit NumUpDownAfterIndent;

	private Label label2;

	private Button BtnApplyListFormat;

	private CheckBox ChkBulletList;

	private CheckBox ChkNumberList;

	private GroupBox groupBox1;

	private CheckBox ChkApplyToAllList;

	private Button BtnSaveAsDefault;

	private CheckBox ChkSameNumberStyle;

	private CheckBox ChkSameNumberFormat;

	private ComboBox CmbNumberStyle;

	private TextBox TxtNumberFormat;

	private Button Btn_GetCurrentListIndent;

	public ListSetUI()
	{
		InitializeComponent();
	}

	private void RdoUseNewListFormat_CheckedChanged(object sender, EventArgs e)
	{
		ChkSameNumberFormat.Enabled = RdoUseNewListFormat.Checked;
		ChkSameNumberStyle.Enabled = RdoUseNewListFormat.Checked;
		NumUpDownNumberIndent.Enabled = RdoUseNewListFormat.Checked;
		NumUpDownTextIndent.Enabled = RdoUseNewListFormat.Checked;
		NumUpDownAfterIndent.Enabled = RdoUseNewListFormat.Checked;
	}

	private void BtnApplyListFormat_Click(object sender, EventArgs e)
	{
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		int numberStyle;
		string numberForamt;
		float numIndent;
		float textIndent;
		float afterNumIndent;
		if (RdoDefaultListFormat.Checked)
		{
			numberStyle = -1;
			numberForamt = null;
			numIndent = defaultValue.ListNumIndent;
			textIndent = defaultValue.ListTextIndent;
			afterNumIndent = defaultValue.ListAfterIndent;
		}
		else
		{
			numberStyle = (ChkSameNumberStyle.Checked ? CmbNumberStyle.SelectedIndex : (-1));
			numberForamt = (ChkSameNumberFormat.Checked ? TxtNumberFormat.Text : null);
			numIndent = (float)NumUpDownNumberIndent.Value;
			textIndent = (float)NumUpDownTextIndent.Value;
			afterNumIndent = (float)NumUpDownAfterIndent.Value;
		}
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (ChkApplyToAllList.Checked)
			{
				foreach (List list in Globals.ThisAddIn.Application.ActiveDocument.Lists)
				{
					int count = list.ListParagraphs.Count;
					if ((list.ListParagraphs[count].Range.ListFormat.ListType == WdListType.wdListSimpleNumbering && ChkNumberList.Checked) || (list.ListParagraphs[count].Range.ListFormat.ListType == WdListType.wdListBullet && ChkBulletList.Checked))
					{
						Globals.ThisAddIn.ListFormat(list, numberStyle, numberForamt, numIndent, textIndent, afterNumIndent);
					}
				}
				return;
			}
			Globals.ThisAddIn.ListFormat(application.Selection.Range.ListFormat.List, numberStyle, numberForamt, numIndent, textIndent, afterNumIndent);
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void BtnSaveAsDefault_Click(object sender, EventArgs e)
	{
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		Selection selection = Globals.ThisAddIn.Application.Selection;
		object Direction = Type.Missing;
		selection.Collapse(ref Direction);
		ListFormat listFormat = application.Selection.Range.ListFormat;
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		if (listFormat.ListType == WdListType.wdListSimpleNumbering || listFormat.ListType == WdListType.wdListBullet)
		{
			defaultValue.ListNumIndent = application.PointsToCentimeters(listFormat.ListTemplate.ListLevels[1].NumberPosition);
			defaultValue.ListTextIndent = application.PointsToCentimeters(listFormat.ListTemplate.ListLevels[1].TextPosition);
			if (listFormat.ListTemplate.ListLevels[1].TabPosition != 9999999f)
			{
				defaultValue.ListAfterIndent = application.PointsToCentimeters(listFormat.ListTemplate.ListLevels[1].TabPosition);
			}
			else
			{
				defaultValue.ListAfterIndent = 0f;
			}
		}
		defaultValue.Save();
	}

	private void ChkSameNumberStyle_CheckedChanged(object sender, EventArgs e)
	{
		CmbNumberStyle.Enabled = ChkSameNumberStyle.Checked;
	}

	private void ChkSameNumberFormat_CheckedChanged(object sender, EventArgs e)
	{
		TxtNumberFormat.Enabled = ChkSameNumberFormat.Checked;
		if (TxtNumberFormat.Enabled && TxtNumberFormat.Text == "")
		{
			TxtNumberFormat.Text = "%1";
		}
	}

	private void TxtNumberFormat_Validating(object sender, CancelEventArgs e)
	{
		if (TxtNumberFormat.Text != "" && !Regex.IsMatch(TxtNumberFormat.Text, ".*%1.*"))
		{
			MessageBox.Show("格式必须包含%1标题编号!", "提醒");
			TxtNumberFormat.Focus();
		}
	}

	private void Btn_GetCurrentListIndent_Click(object sender, EventArgs e)
	{
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		ListFormat listFormat = application.Selection.Range.ListFormat;
		if (listFormat.ListType == WdListType.wdListSimpleNumbering || listFormat.ListType == WdListType.wdListBullet)
		{
			CmbNumberStyle.SelectedIndex = Globals.ThisAddIn.LevelNumStyle.IndexOf(listFormat.ListTemplate.ListLevels[1].NumberStyle);
			TxtNumberFormat.Text = listFormat.ListTemplate.ListLevels[1].NumberFormat;
			if (listFormat.ListTemplate.ListLevels[1].NumberPosition != 9999999f)
			{
				NumUpDownNumberIndent.Value = (decimal)application.PointsToCentimeters(listFormat.ListTemplate.ListLevels[1].NumberPosition);
			}
			if (listFormat.ListTemplate.ListLevels[1].TextPosition != 9999999f)
			{
				NumUpDownTextIndent.Value = (decimal)application.PointsToCentimeters(listFormat.ListTemplate.ListLevels[1].TextPosition);
			}
			if (listFormat.ListTemplate.ListLevels[1].TabPosition != 9999999f)
			{
				NumUpDownAfterIndent.Value = (decimal)application.PointsToCentimeters(listFormat.ListTemplate.ListLevels[1].TabPosition);
			}
			else
			{
				NumUpDownAfterIndent.Value = 0m;
			}
		}
		else
		{
			NumUpDownNumberIndent.Value = 0m;
			NumUpDownTextIndent.Value = 0m;
			NumUpDownAfterIndent.Value = 0m;
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
		this.label8 = new System.Windows.Forms.Label();
		this.RdoDefaultListFormat = new System.Windows.Forms.RadioButton();
		this.RdoUseNewListFormat = new System.Windows.Forms.RadioButton();
		this.label1 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.BtnApplyListFormat = new System.Windows.Forms.Button();
		this.ChkBulletList = new System.Windows.Forms.CheckBox();
		this.ChkNumberList = new System.Windows.Forms.CheckBox();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.ChkApplyToAllList = new System.Windows.Forms.CheckBox();
		this.BtnSaveAsDefault = new System.Windows.Forms.Button();
		this.ChkSameNumberStyle = new System.Windows.Forms.CheckBox();
		this.ChkSameNumberFormat = new System.Windows.Forms.CheckBox();
		this.CmbNumberStyle = new System.Windows.Forms.ComboBox();
		this.TxtNumberFormat = new System.Windows.Forms.TextBox();
		this.Btn_GetCurrentListIndent = new System.Windows.Forms.Button();
		this.NumUpDownAfterIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.NumUpDownTextIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.NumUpDownNumberIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.groupBox1.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownAfterIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownTextIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownNumberIndent).BeginInit();
		base.SuspendLayout();
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(36, 147);
		this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(65, 20);
		this.label8.TabIndex = 74;
		this.label8.Text = "编号缩进";
		this.RdoDefaultListFormat.AutoSize = true;
		this.RdoDefaultListFormat.Checked = true;
		this.RdoDefaultListFormat.Location = new System.Drawing.Point(14, 13);
		this.RdoDefaultListFormat.Name = "RdoDefaultListFormat";
		this.RdoDefaultListFormat.Size = new System.Drawing.Size(139, 24);
		this.RdoDefaultListFormat.TabIndex = 76;
		this.RdoDefaultListFormat.TabStop = true;
		this.RdoDefaultListFormat.Text = "按默认值设置列表";
		this.RdoDefaultListFormat.UseVisualStyleBackColor = true;
		this.RdoUseNewListFormat.AutoSize = true;
		this.RdoUseNewListFormat.Location = new System.Drawing.Point(14, 48);
		this.RdoUseNewListFormat.Name = "RdoUseNewListFormat";
		this.RdoUseNewListFormat.Size = new System.Drawing.Size(139, 24);
		this.RdoUseNewListFormat.TabIndex = 77;
		this.RdoUseNewListFormat.Text = "按下列值设置列表";
		this.RdoUseNewListFormat.UseVisualStyleBackColor = true;
		this.RdoUseNewListFormat.CheckedChanged += new System.EventHandler(RdoUseNewListFormat_CheckedChanged);
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(36, 176);
		this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(65, 20);
		this.label1.TabIndex = 79;
		this.label1.Text = "文本缩进";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(36, 205);
		this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(107, 20);
		this.label2.TabIndex = 81;
		this.label2.Text = "编号后文本缩进";
		this.BtnApplyListFormat.Location = new System.Drawing.Point(153, 294);
		this.BtnApplyListFormat.Name = "BtnApplyListFormat";
		this.BtnApplyListFormat.Size = new System.Drawing.Size(114, 30);
		this.BtnApplyListFormat.TabIndex = 82;
		this.BtnApplyListFormat.Text = "应用设置";
		this.BtnApplyListFormat.UseVisualStyleBackColor = true;
		this.BtnApplyListFormat.Click += new System.EventHandler(BtnApplyListFormat_Click);
		this.ChkBulletList.AutoSize = true;
		this.ChkBulletList.Checked = true;
		this.ChkBulletList.CheckState = System.Windows.Forms.CheckState.Checked;
		this.ChkBulletList.Location = new System.Drawing.Point(11, 25);
		this.ChkBulletList.Name = "ChkBulletList";
		this.ChkBulletList.Size = new System.Drawing.Size(84, 24);
		this.ChkBulletList.TabIndex = 83;
		this.ChkBulletList.Text = "项目符号";
		this.ChkBulletList.UseVisualStyleBackColor = true;
		this.ChkNumberList.AutoSize = true;
		this.ChkNumberList.Checked = true;
		this.ChkNumberList.CheckState = System.Windows.Forms.CheckState.Checked;
		this.ChkNumberList.Location = new System.Drawing.Point(131, 25);
		this.ChkNumberList.Name = "ChkNumberList";
		this.ChkNumberList.Size = new System.Drawing.Size(84, 24);
		this.ChkNumberList.TabIndex = 84;
		this.ChkNumberList.Text = "序号编号";
		this.ChkNumberList.UseVisualStyleBackColor = true;
		this.groupBox1.Controls.Add(this.ChkBulletList);
		this.groupBox1.Controls.Add(this.ChkNumberList);
		this.groupBox1.Location = new System.Drawing.Point(14, 232);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(253, 56);
		this.groupBox1.TabIndex = 85;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "列表类型包括（非大纲列表）";
		this.ChkApplyToAllList.AutoSize = true;
		this.ChkApplyToAllList.Location = new System.Drawing.Point(14, 298);
		this.ChkApplyToAllList.Name = "ChkApplyToAllList";
		this.ChkApplyToAllList.Size = new System.Drawing.Size(126, 24);
		this.ChkApplyToAllList.TabIndex = 86;
		this.ChkApplyToAllList.Text = "应用到全文列表";
		this.ChkApplyToAllList.UseVisualStyleBackColor = true;
		this.BtnSaveAsDefault.Font = new System.Drawing.Font("微软雅黑 Light", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.BtnSaveAsDefault.Location = new System.Drawing.Point(149, 10);
		this.BtnSaveAsDefault.Name = "BtnSaveAsDefault";
		this.BtnSaveAsDefault.Size = new System.Drawing.Size(118, 30);
		this.BtnSaveAsDefault.TabIndex = 87;
		this.BtnSaveAsDefault.Text = "当前缩进设为默认";
		this.BtnSaveAsDefault.UseVisualStyleBackColor = true;
		this.BtnSaveAsDefault.Click += new System.EventHandler(BtnSaveAsDefault_Click);
		this.ChkSameNumberStyle.AutoSize = true;
		this.ChkSameNumberStyle.Enabled = false;
		this.ChkSameNumberStyle.Location = new System.Drawing.Point(36, 84);
		this.ChkSameNumberStyle.Name = "ChkSameNumberStyle";
		this.ChkSameNumberStyle.Size = new System.Drawing.Size(84, 24);
		this.ChkSameNumberStyle.TabIndex = 88;
		this.ChkSameNumberStyle.Text = "编号样式";
		this.ChkSameNumberStyle.UseVisualStyleBackColor = true;
		this.ChkSameNumberStyle.CheckedChanged += new System.EventHandler(ChkSameNumberStyle_CheckedChanged);
		this.ChkSameNumberFormat.AutoSize = true;
		this.ChkSameNumberFormat.Enabled = false;
		this.ChkSameNumberFormat.Location = new System.Drawing.Point(36, 115);
		this.ChkSameNumberFormat.Name = "ChkSameNumberFormat";
		this.ChkSameNumberFormat.Size = new System.Drawing.Size(84, 24);
		this.ChkSameNumberFormat.TabIndex = 89;
		this.ChkSameNumberFormat.Text = "编号格式";
		this.ChkSameNumberFormat.UseVisualStyleBackColor = true;
		this.ChkSameNumberFormat.CheckedChanged += new System.EventHandler(ChkSameNumberFormat_CheckedChanged);
		this.CmbNumberStyle.Enabled = false;
		this.CmbNumberStyle.FormattingEnabled = true;
		this.CmbNumberStyle.Items.AddRange(new object[10] { "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,貳,叁...", "甲,乙,丙...", "正规编号" });
		this.CmbNumberStyle.Location = new System.Drawing.Point(149, 82);
		this.CmbNumberStyle.Name = "CmbNumberStyle";
		this.CmbNumberStyle.Size = new System.Drawing.Size(118, 28);
		this.CmbNumberStyle.TabIndex = 90;
		this.TxtNumberFormat.Enabled = false;
		this.TxtNumberFormat.Location = new System.Drawing.Point(149, 114);
		this.TxtNumberFormat.Name = "TxtNumberFormat";
		this.TxtNumberFormat.Size = new System.Drawing.Size(118, 26);
		this.TxtNumberFormat.TabIndex = 91;
		this.TxtNumberFormat.Validating += new System.ComponentModel.CancelEventHandler(TxtNumberFormat_Validating);
		this.Btn_GetCurrentListIndent.Font = new System.Drawing.Font("微软雅黑 Light", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Btn_GetCurrentListIndent.Location = new System.Drawing.Point(149, 46);
		this.Btn_GetCurrentListIndent.Name = "Btn_GetCurrentListIndent";
		this.Btn_GetCurrentListIndent.Size = new System.Drawing.Size(118, 30);
		this.Btn_GetCurrentListIndent.TabIndex = 92;
		this.Btn_GetCurrentListIndent.Text = "读取当前列表缩进";
		this.Btn_GetCurrentListIndent.UseVisualStyleBackColor = true;
		this.Btn_GetCurrentListIndent.Click += new System.EventHandler(Btn_GetCurrentListIndent_Click);
		this.NumUpDownAfterIndent.DecimalPlaces = 2;
		this.NumUpDownAfterIndent.Enabled = false;
		this.NumUpDownAfterIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownAfterIndent.Label = "厘米";
		this.NumUpDownAfterIndent.Location = new System.Drawing.Point(149, 202);
		this.NumUpDownAfterIndent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownAfterIndent.Name = "NumUpDownAfterIndent";
		this.NumUpDownAfterIndent.Size = new System.Drawing.Size(118, 26);
		this.NumUpDownAfterIndent.TabIndex = 80;
		this.NumUpDownTextIndent.DecimalPlaces = 2;
		this.NumUpDownTextIndent.Enabled = false;
		this.NumUpDownTextIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownTextIndent.Label = "厘米";
		this.NumUpDownTextIndent.Location = new System.Drawing.Point(149, 173);
		this.NumUpDownTextIndent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownTextIndent.Name = "NumUpDownTextIndent";
		this.NumUpDownTextIndent.Size = new System.Drawing.Size(118, 26);
		this.NumUpDownTextIndent.TabIndex = 78;
		this.NumUpDownNumberIndent.DecimalPlaces = 2;
		this.NumUpDownNumberIndent.Enabled = false;
		this.NumUpDownNumberIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownNumberIndent.Label = "厘米";
		this.NumUpDownNumberIndent.Location = new System.Drawing.Point(149, 144);
		this.NumUpDownNumberIndent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownNumberIndent.Name = "NumUpDownNumberIndent";
		this.NumUpDownNumberIndent.Size = new System.Drawing.Size(118, 26);
		this.NumUpDownNumberIndent.TabIndex = 73;
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Btn_GetCurrentListIndent);
		base.Controls.Add(this.TxtNumberFormat);
		base.Controls.Add(this.CmbNumberStyle);
		base.Controls.Add(this.ChkSameNumberFormat);
		base.Controls.Add(this.ChkSameNumberStyle);
		base.Controls.Add(this.BtnSaveAsDefault);
		base.Controls.Add(this.ChkApplyToAllList);
		base.Controls.Add(this.groupBox1);
		base.Controls.Add(this.BtnApplyListFormat);
		base.Controls.Add(this.NumUpDownAfterIndent);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.NumUpDownTextIndent);
		base.Controls.Add(this.label1);
		base.Controls.Add(this.RdoUseNewListFormat);
		base.Controls.Add(this.RdoDefaultListFormat);
		base.Controls.Add(this.NumUpDownNumberIndent);
		base.Controls.Add(this.label8);
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Name = "ListSetUI";
		base.Size = new System.Drawing.Size(280, 330);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownAfterIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownTextIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownNumberIndent).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}