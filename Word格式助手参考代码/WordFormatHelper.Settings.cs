// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.Settings
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WordFormatHelper;

public class Settings : UserControl
{
	private readonly WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;

	private static bool isLoading;

	private IContainer components;

	private GroupBox groupBox1;

	private ComboBox Cmb_TextLineGap;

	private Label label5;

	private GroupBox groupBox2;

	private LineTypeSelectComboBox Cmb_TableLineType;

	private Label label6;

	private Label label8;

	private ComboBox Cmb_TableLineWidth;

	private Button Btn_SetToDefault;

	private GroupBox groupBox3;

	private CheckBox Chk_InlineGroup;

	private CheckBox Chk_InlineSmartArt;

	private CheckBox Chk_InlineChart;

	private CheckBox Chk_InlineLinkedPicture;

	private CheckBox Chk_InlinePicture;

	private GroupBox groupBox4;

	private CheckBox Chk_ShapeGroup;

	private CheckBox Chk_ShapeSmartArt;

	private CheckBox Chk_ShapeChart;

	private CheckBox Chk_ShapeLinkedPicture;

	private CheckBox Chk_ShapePicture;

	private GroupBox groupBox5;

	private NumericUpDownWithUnit NumUpDown_AfterIndent;

	private Label label7;

	private NumericUpDownWithUnit NumUpDown_TextIndent;

	private Label label9;

	private NumericUpDownWithUnit NumUpDown_NumberIndent;

	private Label label10;

	private GroupBox groupBox6;

	private NumericUpDownWithUnit NumUpDown_PageTopMargin;

	private Label label11;

	private NumericUpDownWithUnit NumUpDown_PageRightMargin;

	private Label label14;

	private NumericUpDownWithUnit NumUpDown_PageLeftMargin;

	private Label label13;

	private NumericUpDownWithUnit NumUpDown_PageBottomMargin;

	private Label label12;

	private GroupBox groupBox7;

	private ComboBox Cmb_QRECCLevel;

	private Button Btn_QRLightColor;

	private Button Btn_QRDarkColor;

	private NumericUpDownWithUnit Nud_QRModulePixel;

	private Label label4;

	private Label label3;

	private Label label2;

	private Label label1;

	private CheckBox Chk_QRQuitZone;

	public Settings()
	{
		InitializeComponent();
	}

	private void Settings_Load(object sender, EventArgs e)
	{
		ReadSettings();
	}

	private void Cmb_TextLineGap_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!isLoading)
		{
			defaultValue.HeaderFooter_TextLineGap = Cmb_TextLineGap.SelectedIndex;
			defaultValue.Save();
		}
	}

	private void LineTypeChanged(object sender, EventArgs e)
	{
		if (!isLoading)
		{
			if (Cmb_TableLineType.SelectedIndex == 1 && Cmb_TableLineWidth.Items.Count == 9)
			{
				Cmb_TableLineWidth.Items.Clear();
				Cmb_TableLineWidth.Items.AddRange(new object[6] { "0.25", "0.5", "0.75", "1.5", "2.25", "3.0" });
			}
			else if (Cmb_TableLineType.SelectedIndex != 1 && Cmb_TableLineWidth.Items.Count == 6)
			{
				Cmb_TableLineWidth.Items.Clear();
				Cmb_TableLineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
			}
			if (Cmb_TableLineWidth.SelectedIndex == -1)
			{
				Cmb_TableLineWidth.SelectedIndex = 3;
			}
			defaultValue.TableOutside_LineType = Cmb_TableLineType.SelectedIndex;
			defaultValue.Save();
		}
	}

	private void LineWidthChanged(object sender, EventArgs e)
	{
		if (!isLoading)
		{
			int tableOutside_LineWidth = ((Cmb_TableLineType.SelectedIndex == 1 && (Cmb_TableLineType.SelectedIndex != 1 || Cmb_TableLineWidth.SelectedIndex >= 3)) ? (Cmb_TableLineWidth.SelectedIndex + 1) : Cmb_TableLineWidth.SelectedIndex);
			defaultValue.TableOutside_LineWidth = tableOutside_LineWidth;
			defaultValue.Save();
		}
	}

	private void ListIndentChanged(object sender, EventArgs e)
	{
		float num = (float)(sender as NumericUpDownWithUnit).Value;
		if (!isLoading)
		{
			switch ((sender as NumericUpDownWithUnit).Name)
			{
			case "NumUpDown_NumberIndent":
				defaultValue.ListNumIndent = num;
				break;
			case "NumUpDown_TextIndent":
				defaultValue.ListTextIndent = num;
				break;
			case "NumUpDown_AfterIndent":
				defaultValue.ListAfterIndent = num;
				break;
			}
			defaultValue.Save();
		}
	}

	private void PageMarginChanged(object sender, EventArgs e)
	{
		float num = (float)(sender as NumericUpDownWithUnit).Value;
		if (!isLoading)
		{
			switch ((sender as NumericUpDownWithUnit).Name)
			{
			case "NumUpDown_PageTopMargin":
				defaultValue.PageTopMargin = num;
				break;
			case "NumUpDown_PageBottomMargin":
				defaultValue.PageBottomMargin = num;
				break;
			case "NumUpDown_PageLeftMargin":
				defaultValue.PageLeftMargin = num;
				break;
			case "NumUpDown_PageRightMargin":
				defaultValue.PageRightMargin = num;
				break;
			case "Nud_QRModulePixel":
				defaultValue.QRModulePixel = (int)(sender as NumericUpDownWithUnit).Value;
				break;
			}
			defaultValue.Save();
		}
	}

	private void Btn_SetToDefault_Click(object sender, EventArgs e)
	{
		if (MessageBox.Show("恢复默认将覆盖之前所有用户设置！", "提醒", MessageBoxButtons.OKCancel) != DialogResult.Cancel)
		{
			defaultValue.PageTopMargin = 3f;
			defaultValue.PageBottomMargin = 3f;
			defaultValue.PageLeftMargin = 2.5f;
			defaultValue.PageRightMargin = 2.5f;
			defaultValue.HeaderFooter_TextLineGap = 0;
			defaultValue.TableOutside_LineType = 0;
			defaultValue.ListNumIndent = 0f;
			defaultValue.ListTextIndent = 2f;
			defaultValue.ListAfterIndent = 2f;
			if (Cmb_TableLineWidth.Items.Count != 9)
			{
				Cmb_TableLineWidth.Items.Clear();
				Cmb_TableLineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
			}
			defaultValue.TableOutside_LineWidth = 5;
			defaultValue.ApplyToInlinePicture = true;
			defaultValue.ApplyToInlineLinkedPicture = true;
			defaultValue.ApplyToInlineChart = false;
			defaultValue.ApplyToInlineSmartArt = false;
			defaultValue.ApplyToInlineGroup = false;
			defaultValue.ApplyToShapePicture = true;
			defaultValue.ApplyToShapeLinkedPicture = true;
			defaultValue.ApplyToShapeChart = false;
			defaultValue.ApplyToShapeSmartArt = false;
			defaultValue.ApplyToShapeGroup = false;
			defaultValue.QRCodeDarkColor = Color.Black;
			defaultValue.QRCodeLightColor = Color.White;
			defaultValue.QRCodeQuitZone = true;
			defaultValue.QRModulePixel = 5;
			defaultValue.QRECCLevel = 2;
			defaultValue.Save();
			ReadSettings();
		}
	}

	public void ReadSettings()
	{
		isLoading = true;
		NumUpDown_PageTopMargin.Value = (decimal)defaultValue.PageTopMargin;
		NumUpDown_PageBottomMargin.Value = (decimal)defaultValue.PageBottomMargin;
		NumUpDown_PageLeftMargin.Value = (decimal)defaultValue.PageLeftMargin;
		NumUpDown_PageRightMargin.Value = (decimal)defaultValue.PageRightMargin;
		Cmb_TextLineGap.SelectedIndex = defaultValue.HeaderFooter_TextLineGap;
		Cmb_TableLineType.SelectedIndex = defaultValue.TableOutside_LineType;
		NumUpDown_NumberIndent.Value = (decimal)defaultValue.ListNumIndent;
		NumUpDown_TextIndent.Value = (decimal)defaultValue.ListTextIndent;
		NumUpDown_AfterIndent.Value = (decimal)defaultValue.ListAfterIndent;
		if (Cmb_TableLineType.SelectedIndex == 1)
		{
			Cmb_TableLineWidth.Items.Clear();
			Cmb_TableLineWidth.Items.AddRange(new object[6] { "0.25", "0.5", "0.75", "1.5", "2.25", "3.0" });
			Cmb_TableLineWidth.SelectedIndex = ((defaultValue.TableOutside_LineWidth < 3) ? defaultValue.TableOutside_LineWidth : (defaultValue.TableOutside_LineWidth - 1));
		}
		else
		{
			Cmb_TableLineWidth.Items.Clear();
			Cmb_TableLineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
			Cmb_TableLineWidth.SelectedIndex = defaultValue.TableOutside_LineWidth;
		}
		Chk_InlinePicture.Checked = defaultValue.ApplyToInlinePicture;
		Chk_InlineLinkedPicture.Checked = defaultValue.ApplyToInlineLinkedPicture;
		Chk_InlineChart.Checked = defaultValue.ApplyToInlineChart;
		Chk_InlineSmartArt.Checked = defaultValue.ApplyToInlineSmartArt;
		Chk_InlineGroup.Checked = defaultValue.ApplyToInlineGroup;
		Chk_ShapePicture.Checked = defaultValue.ApplyToShapePicture;
		Chk_ShapeLinkedPicture.Checked = defaultValue.ApplyToShapeLinkedPicture;
		Chk_ShapeChart.Checked = defaultValue.ApplyToShapeChart;
		Chk_ShapeSmartArt.Checked = defaultValue.ApplyToShapeSmartArt;
		Chk_ShapeGroup.Checked = defaultValue.ApplyToShapeGroup;
		Btn_QRDarkColor.BackColor = defaultValue.QRCodeDarkColor;
		Btn_QRLightColor.BackColor = defaultValue.QRCodeLightColor;
		Chk_QRQuitZone.Checked = defaultValue.QRCodeQuitZone;
		Nud_QRModulePixel.Value = defaultValue.QRModulePixel;
		Cmb_QRECCLevel.SelectedIndex = defaultValue.QRECCLevel;
		isLoading = false;
	}

	private void Chk_InlinePicture_CheckedChanged(object sender, EventArgs e)
	{
		if (!isLoading)
		{
			switch ((sender as CheckBox).Name)
			{
			case "Chk_InlinePicture":
				defaultValue.ApplyToInlinePicture = (sender as CheckBox).Checked;
				break;
			case "Chk_InlineLinkedPicture":
				defaultValue.ApplyToInlineLinkedPicture = (sender as CheckBox).Checked;
				break;
			case "Chk_InlineChart":
				defaultValue.ApplyToInlineChart = (sender as CheckBox).Checked;
				break;
			case "Chk_InlineSmartArt":
				defaultValue.ApplyToInlineSmartArt = (sender as CheckBox).Checked;
				break;
			case "Chk_InlineGroup":
				defaultValue.ApplyToInlineGroup = (sender as CheckBox).Checked;
				break;
			case "Chk_ShapePicture":
				defaultValue.ApplyToShapePicture = (sender as CheckBox).Checked;
				break;
			case "Chk_ShapeLinkedPicture":
				defaultValue.ApplyToShapeLinkedPicture = (sender as CheckBox).Checked;
				break;
			case "Chk_ShapeChart":
				defaultValue.ApplyToShapeChart = (sender as CheckBox).Checked;
				break;
			case "Chk_ShapeSmartArt":
				defaultValue.ApplyToShapeSmartArt = (sender as CheckBox).Checked;
				break;
			case "Chk_ShapeGroup":
				defaultValue.ApplyToShapeGroup = (sender as CheckBox).Checked;
				break;
			case "Chk_QRQuitZone":
				defaultValue.QRCodeQuitZone = (sender as CheckBox).Checked;
				break;
			}
			defaultValue.Save();
		}
	}

	private void Btn_QRDarkColor_Click(object sender, EventArgs e)
	{
		ColorDialog colorDialog = new ColorDialog();
		if (colorDialog.ShowDialog() == DialogResult.OK)
		{
			(sender as Button).BackColor = colorDialog.Color;
			if ((sender as Button).Name.Contains("Dark"))
			{
				defaultValue.QRCodeDarkColor = colorDialog.Color;
			}
			else
			{
				defaultValue.QRCodeLightColor = colorDialog.Color;
			}
			defaultValue.Save();
		}
	}

	private void Cmb_QRECCLevel_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!isLoading)
		{
			defaultValue.QRECCLevel = Cmb_QRECCLevel.SelectedIndex;
			defaultValue.Save();
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
		this.Cmb_TextLineGap = new System.Windows.Forms.ComboBox();
		this.label5 = new System.Windows.Forms.Label();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.Cmb_TableLineWidth = new System.Windows.Forms.ComboBox();
		this.label8 = new System.Windows.Forms.Label();
		this.Cmb_TableLineType = new WordFormatHelper.LineTypeSelectComboBox();
		this.label6 = new System.Windows.Forms.Label();
		this.Btn_SetToDefault = new System.Windows.Forms.Button();
		this.groupBox3 = new System.Windows.Forms.GroupBox();
		this.Chk_InlineGroup = new System.Windows.Forms.CheckBox();
		this.Chk_InlineSmartArt = new System.Windows.Forms.CheckBox();
		this.Chk_InlineChart = new System.Windows.Forms.CheckBox();
		this.Chk_InlineLinkedPicture = new System.Windows.Forms.CheckBox();
		this.Chk_InlinePicture = new System.Windows.Forms.CheckBox();
		this.groupBox4 = new System.Windows.Forms.GroupBox();
		this.Chk_ShapeGroup = new System.Windows.Forms.CheckBox();
		this.Chk_ShapeSmartArt = new System.Windows.Forms.CheckBox();
		this.Chk_ShapeChart = new System.Windows.Forms.CheckBox();
		this.Chk_ShapeLinkedPicture = new System.Windows.Forms.CheckBox();
		this.Chk_ShapePicture = new System.Windows.Forms.CheckBox();
		this.groupBox5 = new System.Windows.Forms.GroupBox();
		this.NumUpDown_AfterIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.label7 = new System.Windows.Forms.Label();
		this.NumUpDown_TextIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.label9 = new System.Windows.Forms.Label();
		this.NumUpDown_NumberIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.label10 = new System.Windows.Forms.Label();
		this.groupBox6 = new System.Windows.Forms.GroupBox();
		this.NumUpDown_PageRightMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label14 = new System.Windows.Forms.Label();
		this.NumUpDown_PageLeftMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label13 = new System.Windows.Forms.Label();
		this.NumUpDown_PageBottomMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label12 = new System.Windows.Forms.Label();
		this.NumUpDown_PageTopMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label11 = new System.Windows.Forms.Label();
		this.groupBox7 = new System.Windows.Forms.GroupBox();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.label1 = new System.Windows.Forms.Label();
		this.Chk_QRQuitZone = new System.Windows.Forms.CheckBox();
		this.Cmb_QRECCLevel = new System.Windows.Forms.ComboBox();
		this.Btn_QRLightColor = new System.Windows.Forms.Button();
		this.Btn_QRDarkColor = new System.Windows.Forms.Button();
		this.Nud_QRModulePixel = new WordFormatHelper.NumericUpDownWithUnit();
		this.groupBox1.SuspendLayout();
		this.groupBox2.SuspendLayout();
		this.groupBox3.SuspendLayout();
		this.groupBox4.SuspendLayout();
		this.groupBox5.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_AfterIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_TextIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_NumberIndent).BeginInit();
		this.groupBox6.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageRightMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageLeftMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageBottomMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageTopMargin).BeginInit();
		this.groupBox7.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRModulePixel).BeginInit();
		base.SuspendLayout();
		this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox1.BackColor = System.Drawing.Color.Transparent;
		this.groupBox1.Controls.Add(this.Cmb_TextLineGap);
		this.groupBox1.Controls.Add(this.label5);
		this.groupBox1.Location = new System.Drawing.Point(4, 95);
		this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
		this.groupBox1.Size = new System.Drawing.Size(252, 57);
		this.groupBox1.TabIndex = 0;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "页眉页脚";
		this.Cmb_TextLineGap.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TextLineGap.FormattingEnabled = true;
		this.Cmb_TextLineGap.Items.AddRange(new object[3] { "适中", "紧凑", "宽松" });
		this.Cmb_TextLineGap.Location = new System.Drawing.Point(130, 24);
		this.Cmb_TextLineGap.Name = "Cmb_TextLineGap";
		this.Cmb_TextLineGap.Size = new System.Drawing.Size(114, 22);
		this.Cmb_TextLineGap.TabIndex = 9;
		this.Cmb_TextLineGap.SelectedIndexChanged += new System.EventHandler(Cmb_TextLineGap_SelectedIndexChanged);
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(8, 28);
		this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(105, 14);
		this.label5.TabIndex = 8;
		this.label5.Text = "文字距离分隔线";
		this.groupBox2.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox2.BackColor = System.Drawing.Color.Transparent;
		this.groupBox2.Controls.Add(this.Cmb_TableLineWidth);
		this.groupBox2.Controls.Add(this.label8);
		this.groupBox2.Controls.Add(this.Cmb_TableLineType);
		this.groupBox2.Controls.Add(this.label6);
		this.groupBox2.Location = new System.Drawing.Point(4, 285);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.Size = new System.Drawing.Size(252, 96);
		this.groupBox2.TabIndex = 1;
		this.groupBox2.TabStop = false;
		this.groupBox2.Text = "表格外框";
		this.Cmb_TableLineWidth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TableLineWidth.FormattingEnabled = true;
		this.Cmb_TableLineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
		this.Cmb_TableLineWidth.Location = new System.Drawing.Point(129, 61);
		this.Cmb_TableLineWidth.Name = "Cmb_TableLineWidth";
		this.Cmb_TableLineWidth.Size = new System.Drawing.Size(114, 22);
		this.Cmb_TableLineWidth.TabIndex = 14;
		this.Cmb_TableLineWidth.SelectedIndexChanged += new System.EventHandler(LineWidthChanged);
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(7, 65);
		this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(91, 14);
		this.label8.TabIndex = 11;
		this.label8.Text = "表格外框线宽";
		this.Cmb_TableLineType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		this.Cmb_TableLineType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TableLineType.FormattingEnabled = true;
		this.Cmb_TableLineType.Items.AddRange(new object[4] { "单实线", "双实线", "细粗实线", "粗细实线" });
		this.Cmb_TableLineType.Location = new System.Drawing.Point(129, 29);
		this.Cmb_TableLineType.Name = "Cmb_TableLineType";
		this.Cmb_TableLineType.Size = new System.Drawing.Size(114, 23);
		this.Cmb_TableLineType.TabIndex = 8;
		this.Cmb_TableLineType.SelectedIndexChanged += new System.EventHandler(LineTypeChanged);
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(7, 33);
		this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(91, 14);
		this.label6.TabIndex = 7;
		this.label6.Text = "表格外框线型";
		this.Btn_SetToDefault.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.Btn_SetToDefault.Location = new System.Drawing.Point(3, 762);
		this.Btn_SetToDefault.Name = "Btn_SetToDefault";
		this.Btn_SetToDefault.Size = new System.Drawing.Size(252, 35);
		this.Btn_SetToDefault.TabIndex = 2;
		this.Btn_SetToDefault.Text = "恢复默认设置";
		this.Btn_SetToDefault.UseVisualStyleBackColor = true;
		this.Btn_SetToDefault.Click += new System.EventHandler(Btn_SetToDefault_Click);
		this.groupBox3.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox3.BackColor = System.Drawing.Color.Transparent;
		this.groupBox3.Controls.Add(this.Chk_InlineGroup);
		this.groupBox3.Controls.Add(this.Chk_InlineSmartArt);
		this.groupBox3.Controls.Add(this.Chk_InlineChart);
		this.groupBox3.Controls.Add(this.Chk_InlineLinkedPicture);
		this.groupBox3.Controls.Add(this.Chk_InlinePicture);
		this.groupBox3.Location = new System.Drawing.Point(4, 387);
		this.groupBox3.Name = "groupBox3";
		this.groupBox3.Size = new System.Drawing.Size(252, 111);
		this.groupBox3.TabIndex = 3;
		this.groupBox3.TabStop = false;
		this.groupBox3.Text = "图片格式 - 内嵌图形";
		this.Chk_InlineGroup.AutoSize = true;
		this.Chk_InlineGroup.Location = new System.Drawing.Point(10, 79);
		this.Chk_InlineGroup.Name = "Chk_InlineGroup";
		this.Chk_InlineGroup.Size = new System.Drawing.Size(82, 18);
		this.Chk_InlineGroup.TabIndex = 4;
		this.Chk_InlineGroup.Text = "编组图形";
		this.Chk_InlineGroup.UseVisualStyleBackColor = true;
		this.Chk_InlineGroup.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_InlineSmartArt.AutoSize = true;
		this.Chk_InlineSmartArt.Location = new System.Drawing.Point(129, 57);
		this.Chk_InlineSmartArt.Name = "Chk_InlineSmartArt";
		this.Chk_InlineSmartArt.Size = new System.Drawing.Size(93, 18);
		this.Chk_InlineSmartArt.TabIndex = 3;
		this.Chk_InlineSmartArt.Text = "SmartArt图";
		this.Chk_InlineSmartArt.UseVisualStyleBackColor = true;
		this.Chk_InlineSmartArt.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_InlineChart.AutoSize = true;
		this.Chk_InlineChart.Location = new System.Drawing.Point(10, 55);
		this.Chk_InlineChart.Name = "Chk_InlineChart";
		this.Chk_InlineChart.Size = new System.Drawing.Size(54, 18);
		this.Chk_InlineChart.TabIndex = 2;
		this.Chk_InlineChart.Text = "图表";
		this.Chk_InlineChart.UseVisualStyleBackColor = true;
		this.Chk_InlineChart.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_InlineLinkedPicture.AutoSize = true;
		this.Chk_InlineLinkedPicture.Location = new System.Drawing.Point(129, 33);
		this.Chk_InlineLinkedPicture.Name = "Chk_InlineLinkedPicture";
		this.Chk_InlineLinkedPicture.Size = new System.Drawing.Size(96, 18);
		this.Chk_InlineLinkedPicture.TabIndex = 1;
		this.Chk_InlineLinkedPicture.Text = "链接的图片";
		this.Chk_InlineLinkedPicture.UseVisualStyleBackColor = true;
		this.Chk_InlineLinkedPicture.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_InlinePicture.AutoSize = true;
		this.Chk_InlinePicture.Location = new System.Drawing.Point(10, 31);
		this.Chk_InlinePicture.Name = "Chk_InlinePicture";
		this.Chk_InlinePicture.Size = new System.Drawing.Size(54, 18);
		this.Chk_InlinePicture.TabIndex = 0;
		this.Chk_InlinePicture.Text = "图片";
		this.Chk_InlinePicture.UseVisualStyleBackColor = true;
		this.Chk_InlinePicture.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.groupBox4.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox4.BackColor = System.Drawing.Color.Transparent;
		this.groupBox4.Controls.Add(this.Chk_ShapeGroup);
		this.groupBox4.Controls.Add(this.Chk_ShapeSmartArt);
		this.groupBox4.Controls.Add(this.Chk_ShapeChart);
		this.groupBox4.Controls.Add(this.Chk_ShapeLinkedPicture);
		this.groupBox4.Controls.Add(this.Chk_ShapePicture);
		this.groupBox4.Font = new System.Drawing.Font("等线", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.groupBox4.Location = new System.Drawing.Point(4, 504);
		this.groupBox4.Name = "groupBox4";
		this.groupBox4.Size = new System.Drawing.Size(252, 111);
		this.groupBox4.TabIndex = 5;
		this.groupBox4.TabStop = false;
		this.groupBox4.Text = "图片格式 - 浮置图形";
		this.Chk_ShapeGroup.AutoSize = true;
		this.Chk_ShapeGroup.Location = new System.Drawing.Point(10, 79);
		this.Chk_ShapeGroup.Name = "Chk_ShapeGroup";
		this.Chk_ShapeGroup.Size = new System.Drawing.Size(82, 18);
		this.Chk_ShapeGroup.TabIndex = 4;
		this.Chk_ShapeGroup.Text = "编组图形";
		this.Chk_ShapeGroup.UseVisualStyleBackColor = true;
		this.Chk_ShapeGroup.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_ShapeSmartArt.AutoSize = true;
		this.Chk_ShapeSmartArt.Location = new System.Drawing.Point(129, 57);
		this.Chk_ShapeSmartArt.Name = "Chk_ShapeSmartArt";
		this.Chk_ShapeSmartArt.Size = new System.Drawing.Size(93, 18);
		this.Chk_ShapeSmartArt.TabIndex = 3;
		this.Chk_ShapeSmartArt.Text = "SmartArt图";
		this.Chk_ShapeSmartArt.UseVisualStyleBackColor = true;
		this.Chk_ShapeSmartArt.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_ShapeChart.AutoSize = true;
		this.Chk_ShapeChart.Location = new System.Drawing.Point(10, 55);
		this.Chk_ShapeChart.Name = "Chk_ShapeChart";
		this.Chk_ShapeChart.Size = new System.Drawing.Size(54, 18);
		this.Chk_ShapeChart.TabIndex = 2;
		this.Chk_ShapeChart.Text = "图表";
		this.Chk_ShapeChart.UseVisualStyleBackColor = true;
		this.Chk_ShapeChart.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_ShapeLinkedPicture.AutoSize = true;
		this.Chk_ShapeLinkedPicture.Location = new System.Drawing.Point(129, 33);
		this.Chk_ShapeLinkedPicture.Name = "Chk_ShapeLinkedPicture";
		this.Chk_ShapeLinkedPicture.Size = new System.Drawing.Size(96, 18);
		this.Chk_ShapeLinkedPicture.TabIndex = 1;
		this.Chk_ShapeLinkedPicture.Text = "链接的图片";
		this.Chk_ShapeLinkedPicture.UseVisualStyleBackColor = true;
		this.Chk_ShapeLinkedPicture.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Chk_ShapePicture.AutoSize = true;
		this.Chk_ShapePicture.Location = new System.Drawing.Point(10, 31);
		this.Chk_ShapePicture.Name = "Chk_ShapePicture";
		this.Chk_ShapePicture.Size = new System.Drawing.Size(54, 18);
		this.Chk_ShapePicture.TabIndex = 0;
		this.Chk_ShapePicture.Text = "图片";
		this.Chk_ShapePicture.UseVisualStyleBackColor = true;
		this.Chk_ShapePicture.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.groupBox5.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox5.Controls.Add(this.NumUpDown_AfterIndent);
		this.groupBox5.Controls.Add(this.label7);
		this.groupBox5.Controls.Add(this.NumUpDown_TextIndent);
		this.groupBox5.Controls.Add(this.label9);
		this.groupBox5.Controls.Add(this.NumUpDown_NumberIndent);
		this.groupBox5.Controls.Add(this.label10);
		this.groupBox5.Location = new System.Drawing.Point(4, 159);
		this.groupBox5.Name = "groupBox5";
		this.groupBox5.Size = new System.Drawing.Size(252, 120);
		this.groupBox5.TabIndex = 6;
		this.groupBox5.TabStop = false;
		this.groupBox5.Text = "列表格式";
		this.NumUpDown_AfterIndent.DecimalPlaces = 1;
		this.NumUpDown_AfterIndent.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_AfterIndent.Label = "厘米";
		this.NumUpDown_AfterIndent.Location = new System.Drawing.Point(128, 86);
		this.NumUpDown_AfterIndent.Name = "NumUpDown_AfterIndent";
		this.NumUpDown_AfterIndent.Size = new System.Drawing.Size(114, 22);
		this.NumUpDown_AfterIndent.TabIndex = 11;
		this.NumUpDown_AfterIndent.ValueChanged += new System.EventHandler(ListIndentChanged);
		this.label7.AutoSize = true;
		this.label7.Location = new System.Drawing.Point(6, 90);
		this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(105, 14);
		this.label7.TabIndex = 10;
		this.label7.Text = "编号后文本缩进";
		this.NumUpDown_TextIndent.DecimalPlaces = 1;
		this.NumUpDown_TextIndent.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_TextIndent.Label = "厘米";
		this.NumUpDown_TextIndent.Location = new System.Drawing.Point(128, 54);
		this.NumUpDown_TextIndent.Name = "NumUpDown_TextIndent";
		this.NumUpDown_TextIndent.Size = new System.Drawing.Size(114, 22);
		this.NumUpDown_TextIndent.TabIndex = 9;
		this.NumUpDown_TextIndent.ValueChanged += new System.EventHandler(ListIndentChanged);
		this.label9.AutoSize = true;
		this.label9.Location = new System.Drawing.Point(6, 58);
		this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(63, 14);
		this.label9.TabIndex = 8;
		this.label9.Text = "文本缩进";
		this.NumUpDown_NumberIndent.DecimalPlaces = 1;
		this.NumUpDown_NumberIndent.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_NumberIndent.Label = "厘米";
		this.NumUpDown_NumberIndent.Location = new System.Drawing.Point(128, 22);
		this.NumUpDown_NumberIndent.Name = "NumUpDown_NumberIndent";
		this.NumUpDown_NumberIndent.Size = new System.Drawing.Size(114, 22);
		this.NumUpDown_NumberIndent.TabIndex = 7;
		this.NumUpDown_NumberIndent.ValueChanged += new System.EventHandler(ListIndentChanged);
		this.label10.AutoSize = true;
		this.label10.Location = new System.Drawing.Point(6, 26);
		this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label10.Name = "label10";
		this.label10.Size = new System.Drawing.Size(63, 14);
		this.label10.TabIndex = 6;
		this.label10.Text = "编号缩进";
		this.groupBox6.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox6.Controls.Add(this.NumUpDown_PageRightMargin);
		this.groupBox6.Controls.Add(this.label14);
		this.groupBox6.Controls.Add(this.NumUpDown_PageLeftMargin);
		this.groupBox6.Controls.Add(this.label13);
		this.groupBox6.Controls.Add(this.NumUpDown_PageBottomMargin);
		this.groupBox6.Controls.Add(this.label12);
		this.groupBox6.Controls.Add(this.NumUpDown_PageTopMargin);
		this.groupBox6.Controls.Add(this.label11);
		this.groupBox6.Location = new System.Drawing.Point(4, 3);
		this.groupBox6.Name = "groupBox6";
		this.groupBox6.Size = new System.Drawing.Size(252, 85);
		this.groupBox6.TabIndex = 7;
		this.groupBox6.TabStop = false;
		this.groupBox6.Text = "页边距";
		this.NumUpDown_PageRightMargin.DecimalPlaces = 1;
		this.NumUpDown_PageRightMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_PageRightMargin.Label = "厘米";
		this.NumUpDown_PageRightMargin.Location = new System.Drawing.Point(154, 49);
		this.NumUpDown_PageRightMargin.Name = "NumUpDown_PageRightMargin";
		this.NumUpDown_PageRightMargin.Size = new System.Drawing.Size(90, 22);
		this.NumUpDown_PageRightMargin.TabIndex = 9;
		this.NumUpDown_PageRightMargin.ValueChanged += new System.EventHandler(PageMarginChanged);
		this.label14.AutoSize = true;
		this.label14.Location = new System.Drawing.Point(126, 53);
		this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label14.Name = "label14";
		this.label14.Size = new System.Drawing.Size(21, 14);
		this.label14.TabIndex = 8;
		this.label14.Text = "右";
		this.NumUpDown_PageLeftMargin.DecimalPlaces = 1;
		this.NumUpDown_PageLeftMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_PageLeftMargin.Label = "厘米";
		this.NumUpDown_PageLeftMargin.Location = new System.Drawing.Point(33, 49);
		this.NumUpDown_PageLeftMargin.Name = "NumUpDown_PageLeftMargin";
		this.NumUpDown_PageLeftMargin.Size = new System.Drawing.Size(90, 22);
		this.NumUpDown_PageLeftMargin.TabIndex = 7;
		this.NumUpDown_PageLeftMargin.ValueChanged += new System.EventHandler(PageMarginChanged);
		this.label13.AutoSize = true;
		this.label13.Location = new System.Drawing.Point(8, 53);
		this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label13.Name = "label13";
		this.label13.Size = new System.Drawing.Size(21, 14);
		this.label13.TabIndex = 6;
		this.label13.Text = "左";
		this.NumUpDown_PageBottomMargin.DecimalPlaces = 1;
		this.NumUpDown_PageBottomMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_PageBottomMargin.Label = "厘米";
		this.NumUpDown_PageBottomMargin.Location = new System.Drawing.Point(154, 21);
		this.NumUpDown_PageBottomMargin.Name = "NumUpDown_PageBottomMargin";
		this.NumUpDown_PageBottomMargin.Size = new System.Drawing.Size(90, 22);
		this.NumUpDown_PageBottomMargin.TabIndex = 5;
		this.NumUpDown_PageBottomMargin.ValueChanged += new System.EventHandler(PageMarginChanged);
		this.label12.AutoSize = true;
		this.label12.Location = new System.Drawing.Point(126, 25);
		this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label12.Name = "label12";
		this.label12.Size = new System.Drawing.Size(21, 14);
		this.label12.TabIndex = 4;
		this.label12.Text = "下";
		this.NumUpDown_PageTopMargin.DecimalPlaces = 1;
		this.NumUpDown_PageTopMargin.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.NumUpDown_PageTopMargin.Label = "厘米";
		this.NumUpDown_PageTopMargin.Location = new System.Drawing.Point(33, 21);
		this.NumUpDown_PageTopMargin.Name = "NumUpDown_PageTopMargin";
		this.NumUpDown_PageTopMargin.Size = new System.Drawing.Size(90, 22);
		this.NumUpDown_PageTopMargin.TabIndex = 3;
		this.NumUpDown_PageTopMargin.ValueChanged += new System.EventHandler(PageMarginChanged);
		this.label11.AutoSize = true;
		this.label11.Location = new System.Drawing.Point(8, 25);
		this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label11.Name = "label11";
		this.label11.Size = new System.Drawing.Size(21, 14);
		this.label11.TabIndex = 2;
		this.label11.Text = "上";
		this.groupBox7.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.groupBox7.Controls.Add(this.label4);
		this.groupBox7.Controls.Add(this.label3);
		this.groupBox7.Controls.Add(this.label2);
		this.groupBox7.Controls.Add(this.label1);
		this.groupBox7.Controls.Add(this.Chk_QRQuitZone);
		this.groupBox7.Controls.Add(this.Cmb_QRECCLevel);
		this.groupBox7.Controls.Add(this.Btn_QRLightColor);
		this.groupBox7.Controls.Add(this.Btn_QRDarkColor);
		this.groupBox7.Controls.Add(this.Nud_QRModulePixel);
		this.groupBox7.Location = new System.Drawing.Point(4, 621);
		this.groupBox7.Name = "groupBox7";
		this.groupBox7.Size = new System.Drawing.Size(252, 116);
		this.groupBox7.TabIndex = 8;
		this.groupBox7.TabStop = false;
		this.groupBox7.Text = "二维码设置";
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(8, 59);
		this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(63, 14);
		this.label4.TabIndex = 11;
		this.label4.Text = "模块像素";
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(8, 87);
		this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(63, 14);
		this.label3.TabIndex = 10;
		this.label3.Text = "纠错等级";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(83, 28);
		this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(35, 14);
		this.label2.TabIndex = 9;
		this.label2.Text = "浅色";
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(8, 28);
		this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(35, 14);
		this.label1.TabIndex = 8;
		this.label1.Text = "深色";
		this.Chk_QRQuitZone.AutoSize = true;
		this.Chk_QRQuitZone.Location = new System.Drawing.Point(164, 26);
		this.Chk_QRQuitZone.Name = "Chk_QRQuitZone";
		this.Chk_QRQuitZone.Size = new System.Drawing.Size(82, 18);
		this.Chk_QRQuitZone.TabIndex = 5;
		this.Chk_QRQuitZone.Text = "创建边框";
		this.Chk_QRQuitZone.UseVisualStyleBackColor = true;
		this.Chk_QRQuitZone.CheckedChanged += new System.EventHandler(Chk_InlinePicture_CheckedChanged);
		this.Cmb_QRECCLevel.FormattingEnabled = true;
		this.Cmb_QRECCLevel.Items.AddRange(new object[4] { "L级-低", "M级-中", "Q级-较高", "H级-高" });
		this.Cmb_QRECCLevel.Location = new System.Drawing.Point(78, 83);
		this.Cmb_QRECCLevel.Name = "Cmb_QRECCLevel";
		this.Cmb_QRECCLevel.Size = new System.Drawing.Size(90, 22);
		this.Cmb_QRECCLevel.TabIndex = 3;
		this.Cmb_QRECCLevel.SelectedIndexChanged += new System.EventHandler(Cmb_QRECCLevel_SelectedIndexChanged);
		this.Btn_QRLightColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
		this.Btn_QRLightColor.Location = new System.Drawing.Point(121, 25);
		this.Btn_QRLightColor.Name = "Btn_QRLightColor";
		this.Btn_QRLightColor.Size = new System.Drawing.Size(20, 20);
		this.Btn_QRLightColor.TabIndex = 2;
		this.Btn_QRLightColor.UseVisualStyleBackColor = true;
		this.Btn_QRLightColor.Click += new System.EventHandler(Btn_QRDarkColor_Click);
		this.Btn_QRDarkColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
		this.Btn_QRDarkColor.Location = new System.Drawing.Point(45, 25);
		this.Btn_QRDarkColor.Name = "Btn_QRDarkColor";
		this.Btn_QRDarkColor.Size = new System.Drawing.Size(20, 20);
		this.Btn_QRDarkColor.TabIndex = 1;
		this.Btn_QRDarkColor.UseVisualStyleBackColor = true;
		this.Btn_QRDarkColor.Click += new System.EventHandler(Btn_QRDarkColor_Click);
		this.Nud_QRModulePixel.Label = "像素";
		this.Nud_QRModulePixel.Location = new System.Drawing.Point(78, 55);
		this.Nud_QRModulePixel.Maximum = new decimal(new int[4] { 1000, 0, 0, 0 });
		this.Nud_QRModulePixel.Name = "Nud_QRModulePixel";
		this.Nud_QRModulePixel.Size = new System.Drawing.Size(90, 22);
		this.Nud_QRModulePixel.TabIndex = 0;
		this.Nud_QRModulePixel.ValueChanged += new System.EventHandler(PageMarginChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(96f, 96f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
		this.AutoSize = true;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.groupBox7);
		base.Controls.Add(this.groupBox6);
		base.Controls.Add(this.groupBox5);
		base.Controls.Add(this.groupBox4);
		base.Controls.Add(this.groupBox3);
		base.Controls.Add(this.Btn_SetToDefault);
		base.Controls.Add(this.groupBox2);
		base.Controls.Add(this.groupBox1);
		this.Font = new System.Drawing.Font("等线", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Margin = new System.Windows.Forms.Padding(4);
		base.Name = "Settings";
		base.Size = new System.Drawing.Size(260, 800);
		base.Load += new System.EventHandler(Settings_Load);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		this.groupBox3.ResumeLayout(false);
		this.groupBox3.PerformLayout();
		this.groupBox4.ResumeLayout(false);
		this.groupBox4.PerformLayout();
		this.groupBox5.ResumeLayout(false);
		this.groupBox5.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_AfterIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_TextIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_NumberIndent).EndInit();
		this.groupBox6.ResumeLayout(false);
		this.groupBox6.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageRightMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageLeftMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageBottomMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDown_PageTopMargin).EndInit();
		this.groupBox7.ResumeLayout(false);
		this.groupBox7.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRModulePixel).EndInit();
		base.ResumeLayout(false);
	}
}
