using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class TablePictureSetUI : UserControl
{
	private IContainer components;

	private GroupBox groupBox6;

	private NumericUpDownWithUnit NumUpDownPictureHeight;

	private NumericUpDownWithUnit NumUpDownPictureWidth;

	private CheckBox ChkPictureSameHeight;

	private CheckBox ChkPictureSameWidth;

	private RadioButton RdoApplyToAllPicture;

	private RadioButton RdoApplyToCurrentPicture;

	private Button BtnApplyPictureFormat;

	private CheckBox ChkPictureSigleSpace;

	private GroupBox groupBox5;

	private LineTypeSelectComboBox CmbTableLineType;

	private ComboBox CmbTableLineWidth;

	private CheckBox ChkTableBorderLine;

	private CheckBox ChkDeleteUselessLine;

	private NumericUpDownWithUnit NumUpDownCellMargin;

	private ComboBox CmbCellMarginType;

	private CheckBox ChkCellMargin;

	private CheckBox ChkTableSingleSpace;

	private CheckBox ChkDeleteLeftMargin;

	private RadioButton RdoApplyToAllTable;

	private RadioButton RdoApplyToCurrentTable;

	private CheckBox ChkFirstColumnBold;

	private CheckBox ChkFirstRowBold;

	private Button BtnApplyTableFormat;

	private GroupBox groupBox4;

	private RadioButton RdoApplyTOAll;

	private RadioButton RdoApplyToCurrent;

	private ComboBox CmbPictureAlignmentType;

	private CheckBox ChkPictureAlignment;

	private ComboBox CmbTableAlignmentType;

	private CheckBox ChkTableAlignment;

	private NumericUpDownWithUnit NumUpDownTableLeftIndent;

	private Button BtnApplyTableAlignment;

	private NumericUpDownWithUnit NumUpDownPictureLeftIndent;

	public TablePictureSetUI()
	{
		InitializeComponent();
		CmbTableAlignmentType.SelectedIndex = 0;
		CmbPictureAlignmentType.SelectedIndex = 0;
		CmbCellMarginType.SelectedIndex = 6;
		CmbTableLineType.SelectedIndex = 0;
		CmbTableLineWidth.SelectedIndex = 3;
	}

	private void BtnApplyTableAlignment_Click(object sender, EventArgs e)
	{
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		if (application.ActiveDocument.Tables.Count != 0 && ChkTableAlignment.Checked)
		{
			foreach (Table item in RdoApplyTOAll.Checked ? application.ActiveDocument.Tables : application.Selection.Tables)
			{
				if (item.Tables.NestingLevel <= 2)
				{
					Globals.ThisAddIn.SetTableShapeAlignment(item, CmbTableAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownTableLeftIndent.Value), 1);
				}
			}
		}
		if (application.ActiveDocument.InlineShapes.Count != 0 && ChkPictureAlignment.Checked)
		{
			foreach (InlineShape item2 in RdoApplyTOAll.Checked ? application.ActiveDocument.InlineShapes : application.Selection.InlineShapes)
			{
				if ((item2.Type == WdInlineShapeType.wdInlineShapePicture && defaultValue.ApplyToInlinePicture) || (item2.Type == WdInlineShapeType.wdInlineShapeSmartArt && defaultValue.ApplyToInlineSmartArt) || (item2.Type == WdInlineShapeType.wdInlineShapeChart && defaultValue.ApplyToInlineChart) || (item2.Type == WdInlineShapeType.wdInlineShapeLinkedPicture && defaultValue.ApplyToInlineLinkedPicture))
				{
					Globals.ThisAddIn.SetTableShapeAlignment(item2, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 2);
				}
			}
		}
		if (application.ActiveDocument.Shapes.Count == 0 || !ChkPictureAlignment.Checked)
		{
			return;
		}
		if (RdoApplyTOAll.Checked)
		{
			foreach (Shape shape3 in application.ActiveDocument.Shapes)
			{
				if ((shape3.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (shape3.Type == MsoShapeType.msoSmartArt && defaultValue.ApplyToShapeSmartArt) || (shape3.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart) || (shape3.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToInlineLinkedPicture))
				{
					Globals.ThisAddIn.SetTableShapeAlignment(shape3, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 3);
				}
				if (shape3.Type == MsoShapeType.msoGroup)
				{
					if (shape3.Anchor.Text != null && defaultValue.ApplyToInlineGroup)
					{
						Globals.ThisAddIn.SetTableShapeAlignment(shape3, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 4);
					}
					else if (shape3.Anchor.Text == null && defaultValue.ApplyToShapeGroup)
					{
						Globals.ThisAddIn.SetTableShapeAlignment(shape3, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 3);
					}
				}
			}
			return;
		}
		foreach (Shape item3 in application.Selection.ShapeRange)
		{
			if ((item3.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (item3.Type == MsoShapeType.msoSmartArt && defaultValue.ApplyToShapeSmartArt) || (item3.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart) || (item3.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToInlineLinkedPicture))
			{
				Globals.ThisAddIn.SetTableShapeAlignment(item3, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 3);
			}
			if (item3.Type == MsoShapeType.msoGroup)
			{
				if (item3.Anchor.Text != null && defaultValue.ApplyToInlineGroup)
				{
					Globals.ThisAddIn.SetTableShapeAlignment(item3, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 4);
				}
				else if (item3.Anchor.Text == null && defaultValue.ApplyToShapeGroup)
				{
					Globals.ThisAddIn.SetTableShapeAlignment(item3, CmbPictureAlignmentType.SelectedIndex, application.CentimetersToPoints((float)NumUpDownPictureLeftIndent.Value), 3);
				}
			}
		}
	}

	private void CmbTableAlignmentType_SelectedIndexChanged(object sender, EventArgs e)
	{
		NumUpDownTableLeftIndent.Enabled = CmbTableAlignmentType.SelectedIndex == 0 && CmbTableAlignmentType.Enabled;
	}

	private void CmbPictureAlignmentType_SelectedIndexChanged(object sender, EventArgs e)
	{
		NumUpDownPictureLeftIndent.Enabled = CmbPictureAlignmentType.SelectedIndex == 0 && CmbPictureAlignmentType.Enabled;
	}

	private void ChkTableAlignment_CheckedChanged(object sender, EventArgs e)
	{
		CmbTableAlignmentType.Enabled = ChkTableAlignment.Checked;
		NumUpDownTableLeftIndent.Enabled = CmbTableAlignmentType.SelectedIndex == 0 && ChkTableAlignment.Checked;
	}

	private void ChkPictureAlignment_CheckedChanged(object sender, EventArgs e)
	{
		CmbPictureAlignmentType.Enabled = ChkPictureAlignment.Checked;
		NumUpDownPictureLeftIndent.Enabled = CmbPictureAlignmentType.SelectedIndex == 0 && ChkPictureAlignment.Checked;
	}

	private void ChkCellMargin_CheckedChanged(object sender, EventArgs e)
	{
		CmbCellMarginType.Enabled = ChkCellMargin.Checked;
		NumUpDownCellMargin.Enabled = ChkCellMargin.Checked;
	}

	private void BtnApplyTableFormat_Click(object sender, EventArgs e)
	{
		Tables tables = (RdoApplyToAllTable.Checked ? Globals.ThisAddIn.Application.ActiveDocument.Tables : Globals.ThisAddIn.Application.Selection.Tables);
		if (tables.Count == 0)
		{
			return;
		}
		int borderLineWidth = ((CmbTableLineType.SelectedIndex == 1 && (CmbTableLineType.SelectedIndex != 1 || CmbTableLineWidth.SelectedIndex >= 3)) ? (CmbTableLineWidth.SelectedIndex + 1) : CmbTableLineWidth.SelectedIndex);
		foreach (Table item in tables)
		{
			if (item.NestingLevel == 1)
			{
				Globals.ThisAddIn.SetTableFormat(item, ChkFirstRowBold.Checked, ChkFirstColumnBold.Checked, ChkCellMargin.Checked, CmbCellMarginType.SelectedIndex, (float)NumUpDownCellMargin.Value, ChkDeleteUselessLine.Checked, ChkTableBorderLine.Checked, CmbTableLineType.SelectedIndex, borderLineWidth);
				if (ChkTableSingleSpace.Checked)
				{
					item.Range.ParagraphFormat.SpaceBefore = 0f;
					item.Range.ParagraphFormat.SpaceAfter = 0f;
					item.Range.ParagraphFormat.Space1();
				}
				if (ChkDeleteLeftMargin.Checked)
				{
					item.Range.ParagraphFormat.LeftIndent = 0f;
					item.Range.ParagraphFormat.FirstLineIndent = 0f;
					item.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
				}
			}
		}
	}

	private void ChkTableBorderLine_CheckedChanged(object sender, EventArgs e)
	{
		CmbTableLineType.Enabled = ChkTableBorderLine.Checked;
		CmbTableLineWidth.Enabled = ChkTableBorderLine.Checked;
	}

	private void BtnApplyPictureFormat_Click(object sender, EventArgs e)
	{
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		foreach (InlineShape item in RdoApplyToCurrentPicture.Checked ? Globals.ThisAddIn.Application.Selection.InlineShapes : Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
		{
			if ((item.Type == WdInlineShapeType.wdInlineShapeLinkedPicture && defaultValue.ApplyToInlineLinkedPicture) || (item.Type == WdInlineShapeType.wdInlineShapePicture && defaultValue.ApplyToInlinePicture) || (item.Type == WdInlineShapeType.wdInlineShapeChart && defaultValue.ApplyToInlineChart) || (item.Type == WdInlineShapeType.wdInlineShapeSmartArt && defaultValue.ApplyToInlineSmartArt))
			{
				Globals.ThisAddIn.SetPictureFormat(item, 0, ChkPictureSigleSpace.Checked, ChkPictureSameWidth.Checked, (float)NumUpDownPictureWidth.Value, ChkPictureSameHeight.Checked, (float)NumUpDownPictureHeight.Value);
			}
		}
		if (!ChkPictureSameWidth.Checked && !ChkPictureSameHeight.Checked)
		{
			return;
		}
		if (RdoApplyToAllPicture.Checked)
		{
			foreach (Shape shape3 in Globals.ThisAddIn.Application.ActiveDocument.Shapes)
			{
				if ((shape3.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToShapeLinkedPicture) || (shape3.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (shape3.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart) || (shape3.Type == MsoShapeType.msoSmartArt && defaultValue.ApplyToShapeSmartArt) || (shape3.Type == MsoShapeType.msoGroup && defaultValue.ApplyToShapeGroup))
				{
					Globals.ThisAddIn.SetPictureFormat(shape3, 1, setSingleSpace: false, ChkPictureSameWidth.Checked, (float)NumUpDownPictureWidth.Value, ChkPictureSameHeight.Checked, (float)NumUpDownPictureHeight.Value);
				}
			}
			return;
		}
		foreach (Shape item2 in Globals.ThisAddIn.Application.Selection.ShapeRange)
		{
			if ((item2.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToShapeLinkedPicture) || (item2.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (item2.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart) || (item2.Type == MsoShapeType.msoSmartArt && defaultValue.ApplyToShapeSmartArt) || (item2.Type == MsoShapeType.msoGroup && defaultValue.ApplyToShapeGroup))
			{
				Globals.ThisAddIn.SetPictureFormat(item2, 1, setSingleSpace: false, ChkPictureSameWidth.Checked, (float)NumUpDownPictureWidth.Value, ChkPictureSameHeight.Checked, (float)NumUpDownPictureHeight.Value);
			}
		}
	}

	private void ChkPictureSameWidth_CheckedChanged(object sender, EventArgs e)
	{
		NumUpDownPictureWidth.Enabled = ChkPictureSameWidth.Checked;
	}

	private void ChkPictureSameHeight_CheckedChanged(object sender, EventArgs e)
	{
		NumUpDownPictureHeight.Enabled = ChkPictureSameHeight.Checked;
	}

	private void CmbTableLineWidth_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (CmbTableLineType.SelectedIndex == 1)
		{
			CmbTableLineWidth.Items.Clear();
			CmbTableLineWidth.Items.AddRange(new object[6] { "0.25", "0.5", "0.75", "1.5", "2.25", "3.0" });
			CmbTableLineWidth.SelectedIndex = 3;
		}
		else
		{
			CmbTableLineWidth.Items.Clear();
			CmbTableLineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
			CmbTableLineWidth.SelectedIndex = 3;
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
		this.groupBox6 = new System.Windows.Forms.GroupBox();
		this.NumUpDownPictureHeight = new WordFormatHelper.NumericUpDownWithUnit();
		this.NumUpDownPictureWidth = new WordFormatHelper.NumericUpDownWithUnit();
		this.ChkPictureSameHeight = new System.Windows.Forms.CheckBox();
		this.ChkPictureSameWidth = new System.Windows.Forms.CheckBox();
		this.RdoApplyToAllPicture = new System.Windows.Forms.RadioButton();
		this.RdoApplyToCurrentPicture = new System.Windows.Forms.RadioButton();
		this.BtnApplyPictureFormat = new System.Windows.Forms.Button();
		this.ChkPictureSigleSpace = new System.Windows.Forms.CheckBox();
		this.groupBox5 = new System.Windows.Forms.GroupBox();
		this.CmbTableLineType = new WordFormatHelper.LineTypeSelectComboBox();
		this.CmbTableLineWidth = new System.Windows.Forms.ComboBox();
		this.ChkTableBorderLine = new System.Windows.Forms.CheckBox();
		this.ChkDeleteUselessLine = new System.Windows.Forms.CheckBox();
		this.NumUpDownCellMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.CmbCellMarginType = new System.Windows.Forms.ComboBox();
		this.ChkCellMargin = new System.Windows.Forms.CheckBox();
		this.ChkTableSingleSpace = new System.Windows.Forms.CheckBox();
		this.ChkDeleteLeftMargin = new System.Windows.Forms.CheckBox();
		this.RdoApplyToAllTable = new System.Windows.Forms.RadioButton();
		this.RdoApplyToCurrentTable = new System.Windows.Forms.RadioButton();
		this.ChkFirstColumnBold = new System.Windows.Forms.CheckBox();
		this.ChkFirstRowBold = new System.Windows.Forms.CheckBox();
		this.BtnApplyTableFormat = new System.Windows.Forms.Button();
		this.groupBox4 = new System.Windows.Forms.GroupBox();
		this.RdoApplyTOAll = new System.Windows.Forms.RadioButton();
		this.RdoApplyToCurrent = new System.Windows.Forms.RadioButton();
		this.CmbPictureAlignmentType = new System.Windows.Forms.ComboBox();
		this.ChkPictureAlignment = new System.Windows.Forms.CheckBox();
		this.CmbTableAlignmentType = new System.Windows.Forms.ComboBox();
		this.ChkTableAlignment = new System.Windows.Forms.CheckBox();
		this.NumUpDownTableLeftIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.BtnApplyTableAlignment = new System.Windows.Forms.Button();
		this.NumUpDownPictureLeftIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.groupBox6.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPictureHeight).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPictureWidth).BeginInit();
		this.groupBox5.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownCellMargin).BeginInit();
		this.groupBox4.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownTableLeftIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPictureLeftIndent).BeginInit();
		base.SuspendLayout();
		this.groupBox6.Controls.Add(this.NumUpDownPictureHeight);
		this.groupBox6.Controls.Add(this.NumUpDownPictureWidth);
		this.groupBox6.Controls.Add(this.ChkPictureSameHeight);
		this.groupBox6.Controls.Add(this.ChkPictureSameWidth);
		this.groupBox6.Controls.Add(this.RdoApplyToAllPicture);
		this.groupBox6.Controls.Add(this.RdoApplyToCurrentPicture);
		this.groupBox6.Controls.Add(this.BtnApplyPictureFormat);
		this.groupBox6.Controls.Add(this.ChkPictureSigleSpace);
		this.groupBox6.Location = new System.Drawing.Point(3, 254);
		this.groupBox6.Name = "groupBox6";
		this.groupBox6.Size = new System.Drawing.Size(674, 140);
		this.groupBox6.TabIndex = 78;
		this.groupBox6.TabStop = false;
		this.groupBox6.Text = "图片工具";
		this.NumUpDownPictureHeight.DecimalPlaces = 2;
		this.NumUpDownPictureHeight.Enabled = false;
		this.NumUpDownPictureHeight.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownPictureHeight.Label = "厘米";
		this.NumUpDownPictureHeight.Location = new System.Drawing.Point(326, 67);
		this.NumUpDownPictureHeight.Minimum = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownPictureHeight.Name = "NumUpDownPictureHeight";
		this.NumUpDownPictureHeight.Size = new System.Drawing.Size(86, 26);
		this.NumUpDownPictureHeight.TabIndex = 15;
		this.NumUpDownPictureHeight.Value = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownPictureWidth.DecimalPlaces = 2;
		this.NumUpDownPictureWidth.Enabled = false;
		this.NumUpDownPictureWidth.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownPictureWidth.Label = "厘米";
		this.NumUpDownPictureWidth.Location = new System.Drawing.Point(101, 67);
		this.NumUpDownPictureWidth.Minimum = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownPictureWidth.Name = "NumUpDownPictureWidth";
		this.NumUpDownPictureWidth.Size = new System.Drawing.Size(86, 26);
		this.NumUpDownPictureWidth.TabIndex = 14;
		this.NumUpDownPictureWidth.Value = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.ChkPictureSameHeight.AutoSize = true;
		this.ChkPictureSameHeight.Location = new System.Drawing.Point(236, 68);
		this.ChkPictureSameHeight.Name = "ChkPictureSameHeight";
		this.ChkPictureSameHeight.Size = new System.Drawing.Size(84, 24);
		this.ChkPictureSameHeight.TabIndex = 2;
		this.ChkPictureSameHeight.Text = "高度相等";
		this.ChkPictureSameHeight.UseVisualStyleBackColor = true;
		this.ChkPictureSameHeight.CheckedChanged += new System.EventHandler(ChkPictureSameHeight_CheckedChanged);
		this.ChkPictureSameWidth.AutoSize = true;
		this.ChkPictureSameWidth.Location = new System.Drawing.Point(11, 68);
		this.ChkPictureSameWidth.Name = "ChkPictureSameWidth";
		this.ChkPictureSameWidth.Size = new System.Drawing.Size(84, 24);
		this.ChkPictureSameWidth.TabIndex = 1;
		this.ChkPictureSameWidth.Text = "宽度相等";
		this.ChkPictureSameWidth.UseVisualStyleBackColor = true;
		this.ChkPictureSameWidth.CheckedChanged += new System.EventHandler(ChkPictureSameWidth_CheckedChanged);
		this.RdoApplyToAllPicture.AutoSize = true;
		this.RdoApplyToAllPicture.Location = new System.Drawing.Point(154, 107);
		this.RdoApplyToAllPicture.Name = "RdoApplyToAllPicture";
		this.RdoApplyToAllPicture.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyToAllPicture.TabIndex = 7;
		this.RdoApplyToAllPicture.Text = "应用于所有图片";
		this.RdoApplyToAllPicture.UseVisualStyleBackColor = true;
		this.RdoApplyToCurrentPicture.AutoSize = true;
		this.RdoApplyToCurrentPicture.Checked = true;
		this.RdoApplyToCurrentPicture.Location = new System.Drawing.Point(11, 107);
		this.RdoApplyToCurrentPicture.Name = "RdoApplyToCurrentPicture";
		this.RdoApplyToCurrentPicture.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyToCurrentPicture.TabIndex = 6;
		this.RdoApplyToCurrentPicture.TabStop = true;
		this.RdoApplyToCurrentPicture.Text = "应用于选定图片";
		this.RdoApplyToCurrentPicture.UseVisualStyleBackColor = true;
		this.BtnApplyPictureFormat.Location = new System.Drawing.Point(575, 104);
		this.BtnApplyPictureFormat.Name = "BtnApplyPictureFormat";
		this.BtnApplyPictureFormat.Size = new System.Drawing.Size(90, 30);
		this.BtnApplyPictureFormat.TabIndex = 8;
		this.BtnApplyPictureFormat.Text = "应用设置";
		this.BtnApplyPictureFormat.UseVisualStyleBackColor = true;
		this.BtnApplyPictureFormat.Click += new System.EventHandler(BtnApplyPictureFormat_Click);
		this.ChkPictureSigleSpace.AutoSize = true;
		this.ChkPictureSigleSpace.Location = new System.Drawing.Point(11, 29);
		this.ChkPictureSigleSpace.Name = "ChkPictureSigleSpace";
		this.ChkPictureSigleSpace.Size = new System.Drawing.Size(322, 24);
		this.ChkPictureSigleSpace.TabIndex = 0;
		this.ChkPictureSigleSpace.Text = "设置单倍行距（仅对嵌入行内的图形对象有效）";
		this.ChkPictureSigleSpace.UseVisualStyleBackColor = true;
		this.groupBox5.Controls.Add(this.CmbTableLineType);
		this.groupBox5.Controls.Add(this.CmbTableLineWidth);
		this.groupBox5.Controls.Add(this.ChkTableBorderLine);
		this.groupBox5.Controls.Add(this.ChkDeleteUselessLine);
		this.groupBox5.Controls.Add(this.NumUpDownCellMargin);
		this.groupBox5.Controls.Add(this.CmbCellMarginType);
		this.groupBox5.Controls.Add(this.ChkCellMargin);
		this.groupBox5.Controls.Add(this.ChkTableSingleSpace);
		this.groupBox5.Controls.Add(this.ChkDeleteLeftMargin);
		this.groupBox5.Controls.Add(this.RdoApplyToAllTable);
		this.groupBox5.Controls.Add(this.RdoApplyToCurrentTable);
		this.groupBox5.Controls.Add(this.ChkFirstColumnBold);
		this.groupBox5.Controls.Add(this.ChkFirstRowBold);
		this.groupBox5.Controls.Add(this.BtnApplyTableFormat);
		this.groupBox5.Location = new System.Drawing.Point(3, 108);
		this.groupBox5.Name = "groupBox5";
		this.groupBox5.Size = new System.Drawing.Size(674, 140);
		this.groupBox5.TabIndex = 77;
		this.groupBox5.TabStop = false;
		this.groupBox5.Text = "表格工具";
		this.CmbTableLineType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		this.CmbTableLineType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbTableLineType.Enabled = false;
		this.CmbTableLineType.FormattingEnabled = true;
		this.CmbTableLineType.Items.AddRange(new object[4] { "单实线", "双实线", "细粗实线", "粗细实线" });
		this.CmbTableLineType.Location = new System.Drawing.Point(115, 65);
		this.CmbTableLineType.Name = "CmbTableLineType";
		this.CmbTableLineType.Size = new System.Drawing.Size(115, 27);
		this.CmbTableLineType.TabIndex = 13;
		this.CmbTableLineType.SelectedIndexChanged += new System.EventHandler(CmbTableLineWidth_SelectedIndexChanged);
		this.CmbTableLineWidth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbTableLineWidth.Enabled = false;
		this.CmbTableLineWidth.FormattingEnabled = true;
		this.CmbTableLineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
		this.CmbTableLineWidth.Location = new System.Drawing.Point(234, 64);
		this.CmbTableLineWidth.Name = "CmbTableLineWidth";
		this.CmbTableLineWidth.Size = new System.Drawing.Size(90, 28);
		this.CmbTableLineWidth.TabIndex = 12;
		this.ChkTableBorderLine.AutoSize = true;
		this.ChkTableBorderLine.Location = new System.Drawing.Point(11, 66);
		this.ChkTableBorderLine.Name = "ChkTableBorderLine";
		this.ChkTableBorderLine.Size = new System.Drawing.Size(98, 24);
		this.ChkTableBorderLine.TabIndex = 8;
		this.ChkTableBorderLine.Text = "表外框线型";
		this.ChkTableBorderLine.UseVisualStyleBackColor = true;
		this.ChkTableBorderLine.CheckedChanged += new System.EventHandler(ChkTableBorderLine_CheckedChanged);
		this.ChkDeleteUselessLine.AutoSize = true;
		this.ChkDeleteUselessLine.Location = new System.Drawing.Point(517, 29);
		this.ChkDeleteUselessLine.Name = "ChkDeleteUselessLine";
		this.ChkDeleteUselessLine.Size = new System.Drawing.Size(146, 24);
		this.ChkDeleteUselessLine.TabIndex = 7;
		this.ChkDeleteUselessLine.Text = "删除前导/段后空行";
		this.ChkDeleteUselessLine.UseVisualStyleBackColor = true;
		this.NumUpDownCellMargin.DecimalPlaces = 2;
		this.NumUpDownCellMargin.Enabled = false;
		this.NumUpDownCellMargin.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownCellMargin.Label = "厘米";
		this.NumUpDownCellMargin.Location = new System.Drawing.Point(575, 65);
		this.NumUpDownCellMargin.Name = "NumUpDownCellMargin";
		this.NumUpDownCellMargin.Size = new System.Drawing.Size(86, 26);
		this.NumUpDownCellMargin.TabIndex = 6;
		this.CmbCellMarginType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbCellMarginType.Enabled = false;
		this.CmbCellMarginType.FormattingEnabled = true;
		this.CmbCellMarginType.Items.AddRange(new object[7] { "上", "下", "左", "右", "上下", "左右", "全部" });
		this.CmbCellMarginType.Location = new System.Drawing.Point(465, 64);
		this.CmbCellMarginType.Name = "CmbCellMarginType";
		this.CmbCellMarginType.Size = new System.Drawing.Size(100, 28);
		this.CmbCellMarginType.TabIndex = 5;
		this.ChkCellMargin.AutoSize = true;
		this.ChkCellMargin.Location = new System.Drawing.Point(351, 66);
		this.ChkCellMargin.Name = "ChkCellMargin";
		this.ChkCellMargin.Size = new System.Drawing.Size(112, 24);
		this.ChkCellMargin.TabIndex = 4;
		this.ChkCellMargin.Text = "单元格内边距";
		this.ChkCellMargin.UseVisualStyleBackColor = true;
		this.ChkCellMargin.CheckedChanged += new System.EventHandler(ChkCellMargin_CheckedChanged);
		this.ChkTableSingleSpace.AutoSize = true;
		this.ChkTableSingleSpace.Location = new System.Drawing.Point(387, 29);
		this.ChkTableSingleSpace.Name = "ChkTableSingleSpace";
		this.ChkTableSingleSpace.Size = new System.Drawing.Size(112, 24);
		this.ChkTableSingleSpace.TabIndex = 3;
		this.ChkTableSingleSpace.Text = "设置单倍行距";
		this.ChkTableSingleSpace.UseVisualStyleBackColor = true;
		this.ChkDeleteLeftMargin.AutoSize = true;
		this.ChkDeleteLeftMargin.Location = new System.Drawing.Point(271, 29);
		this.ChkDeleteLeftMargin.Name = "ChkDeleteLeftMargin";
		this.ChkDeleteLeftMargin.Size = new System.Drawing.Size(98, 24);
		this.ChkDeleteLeftMargin.TabIndex = 2;
		this.ChkDeleteLeftMargin.Text = "消除左缩进";
		this.ChkDeleteLeftMargin.UseVisualStyleBackColor = true;
		this.RdoApplyToAllTable.AutoSize = true;
		this.RdoApplyToAllTable.Location = new System.Drawing.Point(154, 104);
		this.RdoApplyToAllTable.Name = "RdoApplyToAllTable";
		this.RdoApplyToAllTable.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyToAllTable.TabIndex = 10;
		this.RdoApplyToAllTable.Text = "应用于所有表格";
		this.RdoApplyToAllTable.UseVisualStyleBackColor = true;
		this.RdoApplyToCurrentTable.AutoSize = true;
		this.RdoApplyToCurrentTable.Checked = true;
		this.RdoApplyToCurrentTable.Location = new System.Drawing.Point(11, 104);
		this.RdoApplyToCurrentTable.Name = "RdoApplyToCurrentTable";
		this.RdoApplyToCurrentTable.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyToCurrentTable.TabIndex = 9;
		this.RdoApplyToCurrentTable.TabStop = true;
		this.RdoApplyToCurrentTable.Text = "应用于选定表格";
		this.RdoApplyToCurrentTable.UseVisualStyleBackColor = true;
		this.ChkFirstColumnBold.AutoSize = true;
		this.ChkFirstColumnBold.Location = new System.Drawing.Point(141, 29);
		this.ChkFirstColumnBold.Name = "ChkFirstColumnBold";
		this.ChkFirstColumnBold.Size = new System.Drawing.Size(112, 24);
		this.ChkFirstColumnBold.TabIndex = 1;
		this.ChkFirstColumnBold.Text = "首列字体加粗";
		this.ChkFirstColumnBold.UseVisualStyleBackColor = true;
		this.ChkFirstRowBold.AutoSize = true;
		this.ChkFirstRowBold.Location = new System.Drawing.Point(11, 29);
		this.ChkFirstRowBold.Name = "ChkFirstRowBold";
		this.ChkFirstRowBold.Size = new System.Drawing.Size(112, 24);
		this.ChkFirstRowBold.TabIndex = 0;
		this.ChkFirstRowBold.Text = "首行字体加粗";
		this.ChkFirstRowBold.UseVisualStyleBackColor = true;
		this.BtnApplyTableFormat.Location = new System.Drawing.Point(575, 101);
		this.BtnApplyTableFormat.Name = "BtnApplyTableFormat";
		this.BtnApplyTableFormat.Size = new System.Drawing.Size(90, 30);
		this.BtnApplyTableFormat.TabIndex = 11;
		this.BtnApplyTableFormat.Text = "应用设置";
		this.BtnApplyTableFormat.UseVisualStyleBackColor = true;
		this.BtnApplyTableFormat.Click += new System.EventHandler(BtnApplyTableFormat_Click);
		this.groupBox4.Controls.Add(this.RdoApplyTOAll);
		this.groupBox4.Controls.Add(this.RdoApplyToCurrent);
		this.groupBox4.Controls.Add(this.CmbPictureAlignmentType);
		this.groupBox4.Controls.Add(this.ChkPictureAlignment);
		this.groupBox4.Controls.Add(this.CmbTableAlignmentType);
		this.groupBox4.Controls.Add(this.ChkTableAlignment);
		this.groupBox4.Controls.Add(this.NumUpDownTableLeftIndent);
		this.groupBox4.Controls.Add(this.BtnApplyTableAlignment);
		this.groupBox4.Controls.Add(this.NumUpDownPictureLeftIndent);
		this.groupBox4.Location = new System.Drawing.Point(3, 3);
		this.groupBox4.Name = "groupBox4";
		this.groupBox4.Size = new System.Drawing.Size(674, 99);
		this.groupBox4.TabIndex = 76;
		this.groupBox4.TabStop = false;
		this.groupBox4.Text = "对齐设置";
		this.RdoApplyTOAll.AutoSize = true;
		this.RdoApplyTOAll.Location = new System.Drawing.Point(154, 68);
		this.RdoApplyTOAll.Name = "RdoApplyTOAll";
		this.RdoApplyTOAll.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyTOAll.TabIndex = 12;
		this.RdoApplyTOAll.Text = "应用于所有图表";
		this.RdoApplyTOAll.UseVisualStyleBackColor = true;
		this.RdoApplyToCurrent.AutoSize = true;
		this.RdoApplyToCurrent.Checked = true;
		this.RdoApplyToCurrent.Location = new System.Drawing.Point(12, 68);
		this.RdoApplyToCurrent.Name = "RdoApplyToCurrent";
		this.RdoApplyToCurrent.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyToCurrent.TabIndex = 11;
		this.RdoApplyToCurrent.TabStop = true;
		this.RdoApplyToCurrent.Text = "应用于选定图表";
		this.RdoApplyToCurrent.UseVisualStyleBackColor = true;
		this.CmbPictureAlignmentType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbPictureAlignmentType.Enabled = false;
		this.CmbPictureAlignmentType.FormattingEnabled = true;
		this.CmbPictureAlignmentType.Items.AddRange(new object[3] { "左对齐", "右对齐", "居中对齐" });
		this.CmbPictureAlignmentType.Location = new System.Drawing.Point(469, 28);
		this.CmbPictureAlignmentType.Name = "CmbPictureAlignmentType";
		this.CmbPictureAlignmentType.Size = new System.Drawing.Size(100, 28);
		this.CmbPictureAlignmentType.TabIndex = 5;
		this.CmbPictureAlignmentType.SelectedIndexChanged += new System.EventHandler(CmbPictureAlignmentType_SelectedIndexChanged);
		this.ChkPictureAlignment.AutoSize = true;
		this.ChkPictureAlignment.Location = new System.Drawing.Point(351, 30);
		this.ChkPictureAlignment.Name = "ChkPictureAlignment";
		this.ChkPictureAlignment.Size = new System.Drawing.Size(112, 24);
		this.ChkPictureAlignment.TabIndex = 4;
		this.ChkPictureAlignment.Text = "图形对齐方式";
		this.ChkPictureAlignment.UseVisualStyleBackColor = true;
		this.ChkPictureAlignment.CheckedChanged += new System.EventHandler(ChkPictureAlignment_CheckedChanged);
		this.CmbTableAlignmentType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbTableAlignmentType.Enabled = false;
		this.CmbTableAlignmentType.FormattingEnabled = true;
		this.CmbTableAlignmentType.Items.AddRange(new object[3] { "左对齐", "右对齐", "居中对齐" });
		this.CmbTableAlignmentType.Location = new System.Drawing.Point(133, 28);
		this.CmbTableAlignmentType.Name = "CmbTableAlignmentType";
		this.CmbTableAlignmentType.Size = new System.Drawing.Size(100, 28);
		this.CmbTableAlignmentType.TabIndex = 1;
		this.CmbTableAlignmentType.SelectedIndexChanged += new System.EventHandler(CmbTableAlignmentType_SelectedIndexChanged);
		this.ChkTableAlignment.AutoSize = true;
		this.ChkTableAlignment.Location = new System.Drawing.Point(11, 30);
		this.ChkTableAlignment.Name = "ChkTableAlignment";
		this.ChkTableAlignment.Size = new System.Drawing.Size(112, 24);
		this.ChkTableAlignment.TabIndex = 0;
		this.ChkTableAlignment.Text = "表格对齐方式";
		this.ChkTableAlignment.UseVisualStyleBackColor = true;
		this.ChkTableAlignment.CheckedChanged += new System.EventHandler(ChkTableAlignment_CheckedChanged);
		this.NumUpDownTableLeftIndent.DecimalPlaces = 2;
		this.NumUpDownTableLeftIndent.Enabled = false;
		this.NumUpDownTableLeftIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownTableLeftIndent.Label = "厘米";
		this.NumUpDownTableLeftIndent.Location = new System.Drawing.Point(236, 29);
		this.NumUpDownTableLeftIndent.Name = "NumUpDownTableLeftIndent";
		this.NumUpDownTableLeftIndent.Size = new System.Drawing.Size(90, 26);
		this.NumUpDownTableLeftIndent.TabIndex = 2;
		this.BtnApplyTableAlignment.Location = new System.Drawing.Point(575, 65);
		this.BtnApplyTableAlignment.Name = "BtnApplyTableAlignment";
		this.BtnApplyTableAlignment.Size = new System.Drawing.Size(90, 30);
		this.BtnApplyTableAlignment.TabIndex = 9;
		this.BtnApplyTableAlignment.Text = "应用设置";
		this.BtnApplyTableAlignment.UseVisualStyleBackColor = true;
		this.BtnApplyTableAlignment.Click += new System.EventHandler(BtnApplyTableAlignment_Click);
		this.NumUpDownPictureLeftIndent.DecimalPlaces = 2;
		this.NumUpDownPictureLeftIndent.Enabled = false;
		this.NumUpDownPictureLeftIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownPictureLeftIndent.Label = "厘米";
		this.NumUpDownPictureLeftIndent.Location = new System.Drawing.Point(575, 29);
		this.NumUpDownPictureLeftIndent.Name = "NumUpDownPictureLeftIndent";
		this.NumUpDownPictureLeftIndent.Size = new System.Drawing.Size(90, 26);
		this.NumUpDownPictureLeftIndent.TabIndex = 6;
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.groupBox6);
		base.Controls.Add(this.groupBox5);
		base.Controls.Add(this.groupBox4);
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Name = "TablePictureSetUI";
		base.Size = new System.Drawing.Size(680, 400);
		this.groupBox6.ResumeLayout(false);
		this.groupBox6.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPictureHeight).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPictureWidth).EndInit();
		this.groupBox5.ResumeLayout(false);
		this.groupBox5.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownCellMargin).EndInit();
		this.groupBox4.ResumeLayout(false);
		this.groupBox4.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownTableLeftIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPictureLeftIndent).EndInit();
		base.ResumeLayout(false);
	}
}
}