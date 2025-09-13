using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class TableRowHeightUI : UserControl
{
	private float top;

	private float bottom;

	private float left;

	private float right;

	private float lineSpacing = 1f;

	private bool spacingByLine = true;

	private bool cmUnit;

	private IContainer components;

	private GroupBox groupBox1;

	private ComboBox Cmb_InnerMarginSet;

	private Label label3;

	private NumericUpDownWithUnit Num_RightMargin;

	private Label label4;

	private NumericUpDownWithUnit Num_LeftMargin;

	private Label label2;

	private NumericUpDownWithUnit Num_BottomMargin;

	private Label label1;

	private NumericUpDownWithUnit Num_TopMargin;

	private Label label5;

	private ComboBox Cmb_TextLineSpacing;

	private Label label6;

	private Button Btn_Action;

	private CheckBox Chk_ApplyToAllTables;

	public TableRowHeightUI()
	{
		InitializeComponent();
		Cmb_InnerMarginSet.SelectedIndex = 0;
		Cmb_TextLineSpacing.SelectedIndex = 0;
	}

	private void Cmb_InnerMarginSet_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (Cmb_InnerMarginSet.SelectedIndex == 0)
		{
			NumericUpDownWithUnit num_BottomMargin = Num_BottomMargin;
			NumericUpDownWithUnit num_LeftMargin = Num_LeftMargin;
			bool flag = (Num_RightMargin.Enabled = false);
			bool enabled = (num_LeftMargin.Enabled = flag);
			num_BottomMargin.Enabled = enabled;
			NumericUpDownWithUnit num_BottomMargin2 = Num_BottomMargin;
			NumericUpDownWithUnit num_LeftMargin2 = Num_LeftMargin;
			decimal num = (Num_RightMargin.Value = Num_TopMargin.Value);
			decimal value2 = (num_LeftMargin2.Value = num);
			num_BottomMargin2.Value = value2;
		}
		else if (Cmb_InnerMarginSet.SelectedIndex == 1)
		{
			Num_LeftMargin.Enabled = true;
			NumericUpDownWithUnit num_BottomMargin3 = Num_BottomMargin;
			bool enabled = (Num_RightMargin.Enabled = false);
			num_BottomMargin3.Enabled = enabled;
			Num_BottomMargin.Value = Num_TopMargin.Value;
			Num_RightMargin.Value = Num_LeftMargin.Value;
		}
		else
		{
			NumericUpDownWithUnit num_BottomMargin4 = Num_BottomMargin;
			NumericUpDownWithUnit num_LeftMargin3 = Num_LeftMargin;
			bool flag = (Num_RightMargin.Enabled = true);
			bool enabled = (num_LeftMargin3.Enabled = flag);
			num_BottomMargin4.Enabled = enabled;
		}
	}

	private void Num_TopMargin_ValueChanged(object sender, EventArgs e)
	{
		if (Cmb_InnerMarginSet.SelectedIndex == 0)
		{
			NumericUpDownWithUnit num_BottomMargin = Num_BottomMargin;
			NumericUpDownWithUnit num_LeftMargin = Num_LeftMargin;
			decimal num = (Num_RightMargin.Value = Num_TopMargin.Value);
			decimal value2 = (num_LeftMargin.Value = num);
			num_BottomMargin.Value = value2;
		}
		else if (Cmb_InnerMarginSet.SelectedIndex == 1)
		{
			Num_BottomMargin.Value = Num_TopMargin.Value;
		}
	}

	private void Num_LeftMargin_ValueChanged(object sender, EventArgs e)
	{
		if (Cmb_InnerMarginSet.SelectedIndex == 1)
		{
			Num_RightMargin.Value = Num_LeftMargin.Value;
		}
	}

	private void Btn_Action_Click(object sender, EventArgs e)
	{
		top = Globals.ThisAddIn.Application.CentimetersToPoints((float)Num_TopMargin.Value);
		bottom = Globals.ThisAddIn.Application.CentimetersToPoints((float)Num_BottomMargin.Value);
		left = Globals.ThisAddIn.Application.CentimetersToPoints((float)Num_LeftMargin.Value);
		right = Globals.ThisAddIn.Application.CentimetersToPoints((float)Num_RightMargin.Value);
		string text = Cmb_TextLineSpacing.Text;
		if (text.EndsWith("行"))
		{
			text = text.Replace("行", "").Trim();
			spacingByLine = true;
		}
		else if (text.EndsWith("厘米"))
		{
			text = text.Replace("厘米", "").Trim();
			spacingByLine = false;
			cmUnit = true;
		}
		else if (text.EndsWith("cm"))
		{
			text = text.Replace("cm", "").Trim();
			spacingByLine = false;
			cmUnit = true;
		}
		else if (text.EndsWith("磅"))
		{
			text = text.Replace("磅", "").Trim();
			spacingByLine = false;
		}
		else if (text.EndsWith("pt"))
		{
			text = text.Replace("pt", "").Trim();
			spacingByLine = false;
		}
		else
		{
			spacingByLine = false;
		}
		try
		{
			lineSpacing = Convert.ToSingle(text);
		}
		catch
		{
			MessageBox.Show("行距设置不正确！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			return;
		}
		if (lineSpacing <= 0f)
		{
			MessageBox.Show("行距设置应大于0！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			return;
		}
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (Chk_ApplyToAllTables.Checked)
			{
				foreach (Table table in Globals.ThisAddIn.Application.ActiveDocument.Tables)
				{
					SetTableRowHeight(table);
				}
				return;
			}
			if (Globals.ThisAddIn.Application.Selection.Tables.Count < 1)
			{
				return;
			}
			foreach (Table table2 in Globals.ThisAddIn.Application.Selection.Tables)
			{
				SetTableRowHeight(table2);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void SetTableRowHeight(Table thisTable)
	{
		if (thisTable.NestingLevel != 1)
		{
			return;
		}
		thisTable.AllowAutoFit = false;
		thisTable.Rows.Height = 0f;
		if (spacingByLine)
		{
			thisTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
			thisTable.Range.ParagraphFormat.LineSpacing = Globals.ThisAddIn.Application.LinesToPoints(lineSpacing);
		}
		else
		{
			if (cmUnit)
			{
				lineSpacing = Globals.ThisAddIn.Application.CentimetersToPoints(lineSpacing);
			}
			thisTable.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
			thisTable.Range.ParagraphFormat.BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignCenter;
			thisTable.Range.ParagraphFormat.LineSpacing = lineSpacing;
		}
		thisTable.TopPadding = top;
		thisTable.BottomPadding = bottom;
		thisTable.LeftPadding = left;
		thisTable.RightPadding = right;
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
		this.label1 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label4 = new System.Windows.Forms.Label();
		this.Cmb_InnerMarginSet = new System.Windows.Forms.ComboBox();
		this.label5 = new System.Windows.Forms.Label();
		this.Cmb_TextLineSpacing = new System.Windows.Forms.ComboBox();
		this.label6 = new System.Windows.Forms.Label();
		this.Btn_Action = new System.Windows.Forms.Button();
		this.Num_RightMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.Num_LeftMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.Num_BottomMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.Num_TopMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.Chk_ApplyToAllTables = new System.Windows.Forms.CheckBox();
		this.groupBox1.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_RightMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_LeftMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BottomMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_TopMargin).BeginInit();
		base.SuspendLayout();
		this.groupBox1.Controls.Add(this.Cmb_InnerMarginSet);
		this.groupBox1.Controls.Add(this.Num_RightMargin);
		this.groupBox1.Controls.Add(this.Num_LeftMargin);
		this.groupBox1.Controls.Add(this.Num_BottomMargin);
		this.groupBox1.Controls.Add(this.Num_TopMargin);
		this.groupBox1.Controls.Add(this.label5);
		this.groupBox1.Controls.Add(this.label3);
		this.groupBox1.Controls.Add(this.label4);
		this.groupBox1.Controls.Add(this.label2);
		this.groupBox1.Controls.Add(this.label1);
		this.groupBox1.Location = new System.Drawing.Point(3, 3);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(244, 122);
		this.groupBox1.TabIndex = 0;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "内边距";
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(6, 59);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(22, 18);
		this.label1.TabIndex = 1;
		this.label1.Text = "上";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(130, 59);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(22, 18);
		this.label2.TabIndex = 3;
		this.label2.Text = "下";
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(130, 88);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(22, 18);
		this.label3.TabIndex = 7;
		this.label3.Text = "右";
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(6, 88);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(22, 18);
		this.label4.TabIndex = 5;
		this.label4.Text = "左";
		this.Cmb_InnerMarginSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_InnerMarginSet.FormattingEnabled = true;
		this.Cmb_InnerMarginSet.Items.AddRange(new object[3] { "统一内边距", "上下相等、左右相等", "自由设置" });
		this.Cmb_InnerMarginSet.Location = new System.Drawing.Point(93, 24);
		this.Cmb_InnerMarginSet.Name = "Cmb_InnerMarginSet";
		this.Cmb_InnerMarginSet.Size = new System.Drawing.Size(140, 26);
		this.Cmb_InnerMarginSet.TabIndex = 8;
		this.Cmb_InnerMarginSet.SelectedIndexChanged += new System.EventHandler(Cmb_InnerMarginSet_SelectedIndexChanged);
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(6, 28);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(78, 18);
		this.label5.TabIndex = 9;
		this.label5.Text = "内边距设定";
		this.Cmb_TextLineSpacing.FormattingEnabled = true;
		this.Cmb_TextLineSpacing.Items.AddRange(new object[9] { "1.0 行", "1.25 行", "1.5 行", "1.75 行", "2.0 行", "2.25 行", "2.5 行", "2.75 行", "3.0 行" });
		this.Cmb_TextLineSpacing.Location = new System.Drawing.Point(96, 130);
		this.Cmb_TextLineSpacing.Name = "Cmb_TextLineSpacing";
		this.Cmb_TextLineSpacing.Size = new System.Drawing.Size(151, 26);
		this.Cmb_TextLineSpacing.TabIndex = 10;
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(3, 134);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(64, 18);
		this.label6.TabIndex = 11;
		this.label6.Text = "内容行距";
		this.Btn_Action.Location = new System.Drawing.Point(127, 172);
		this.Btn_Action.Name = "Btn_Action";
		this.Btn_Action.Size = new System.Drawing.Size(120, 30);
		this.Btn_Action.TabIndex = 12;
		this.Btn_Action.Text = "确定";
		this.Btn_Action.UseVisualStyleBackColor = true;
		this.Btn_Action.Click += new System.EventHandler(Btn_Action_Click);
		this.Num_RightMargin.DecimalPlaces = 2;
		this.Num_RightMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Num_RightMargin.Label = "cm";
		this.Num_RightMargin.Location = new System.Drawing.Point(153, 85);
		this.Num_RightMargin.Maximum = new decimal(new int[4] { 5000, 0, 0, 0 });
		this.Num_RightMargin.Name = "Num_RightMargin";
		this.Num_RightMargin.Size = new System.Drawing.Size(80, 25);
		this.Num_RightMargin.TabIndex = 6;
		this.Num_LeftMargin.DecimalPlaces = 2;
		this.Num_LeftMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Num_LeftMargin.Label = "cm";
		this.Num_LeftMargin.Location = new System.Drawing.Point(31, 85);
		this.Num_LeftMargin.Maximum = new decimal(new int[4] { 5000, 0, 0, 0 });
		this.Num_LeftMargin.Name = "Num_LeftMargin";
		this.Num_LeftMargin.Size = new System.Drawing.Size(80, 25);
		this.Num_LeftMargin.TabIndex = 4;
		this.Num_LeftMargin.ValueChanged += new System.EventHandler(Num_LeftMargin_ValueChanged);
		this.Num_BottomMargin.DecimalPlaces = 2;
		this.Num_BottomMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Num_BottomMargin.Label = "cm";
		this.Num_BottomMargin.Location = new System.Drawing.Point(153, 56);
		this.Num_BottomMargin.Maximum = new decimal(new int[4] { 5000, 0, 0, 0 });
		this.Num_BottomMargin.Name = "Num_BottomMargin";
		this.Num_BottomMargin.Size = new System.Drawing.Size(80, 25);
		this.Num_BottomMargin.TabIndex = 2;
		this.Num_TopMargin.DecimalPlaces = 2;
		this.Num_TopMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Num_TopMargin.Label = "cm";
		this.Num_TopMargin.Location = new System.Drawing.Point(31, 56);
		this.Num_TopMargin.Maximum = new decimal(new int[4] { 5000, 0, 0, 0 });
		this.Num_TopMargin.Name = "Num_TopMargin";
		this.Num_TopMargin.Size = new System.Drawing.Size(80, 25);
		this.Num_TopMargin.TabIndex = 0;
		this.Num_TopMargin.ValueChanged += new System.EventHandler(Num_TopMargin_ValueChanged);
		this.Chk_ApplyToAllTables.AutoSize = true;
		this.Chk_ApplyToAllTables.Location = new System.Drawing.Point(7, 176);
		this.Chk_ApplyToAllTables.Name = "Chk_ApplyToAllTables";
		this.Chk_ApplyToAllTables.Size = new System.Drawing.Size(97, 22);
		this.Chk_ApplyToAllTables.TabIndex = 13;
		this.Chk_ApplyToAllTables.Text = "应用于全文";
		this.Chk_ApplyToAllTables.UseVisualStyleBackColor = true;
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 18f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Chk_ApplyToAllTables);
		base.Controls.Add(this.Btn_Action);
		base.Controls.Add(this.Cmb_TextLineSpacing);
		base.Controls.Add(this.label6);
		base.Controls.Add(this.groupBox1);
		this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		base.Margin = new System.Windows.Forms.Padding(4);
		base.Name = "TableRowHeightUI";
		base.Size = new System.Drawing.Size(251, 209);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_RightMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_LeftMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BottomMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_TopMargin).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}