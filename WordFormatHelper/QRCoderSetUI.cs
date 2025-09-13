using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EnocheastyBarCode.BarCode;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class QRCoderSetUI : UserControl
{
	private IContainer components;

	private PictureBox Pic_ImageView;

	private TextBox Txt_QRString;

	private GroupBox groupBox1;

	private Button Btn_QRLightColor;

	private Button Btn_QRDarkColor;

	private ComboBox Cmb_QRECCLevel;

	private NumericUpDownWithUnit Nud_QRModulePixel;

	private NumericUpDownWithUnit Nud_QRIconSize;

	private CheckBox Chk_QRQuitZone;

	private NumericUpDownWithUnit Nud_QRIconBorderWidth;

	private Label label7;

	private Label label6;

	private Label label5;

	private Label label4;

	private Label label3;

	private Label label2;

	private Label label1;

	private Button Btn_GetIconPath;

	private TextBox Txt_QRIconPath;

	private Button Btn_QRCodeInsert;

	private Button Btn_QRCodeView;

	private Label Lab_ImageView;

	private CheckBox Chk_IsolatedModule;

	private Label label9;

	private NumericUpDownWithUnit Num_RectangleRadius;

	private RadioButton Rdo_RectangleIcon;

	private RadioButton Rdo_RoundIcon;

	private CheckBox Chk_RoundModule;

	private Label label10;

	private Button Btn_IconBackGround;

	private TabControl Tab_Codes;

	private TabPage Tab_QRCode;

	private TabPage Tab_BarCode;

	private TextBox Txt_BarCodeString;

	private GroupBox groupBox2;

	private Label label8;

	private Label label11;

	private Label label12;

	private Label label13;

	private ComboBox Cmb_BarCodeType;

	private NumericUpDownWithUnit Num_BarUnitWidth;

	private Button Btn_BarCodeLightColor;

	private Button Btn_BarCodeDarkColor;

	private Label label15;

	private ComboBox Cmb_TextFont;

	private Label label14;

	private NumericUpDownWithUnit Num_BarCodeHeight;

	private CheckBox Chk_ShowText;

	private Label label17;

	private NumericUpDownWithUnit Num_BarCodeTextHeight;

	private Label label16;

	private Button Btn_BarCodeTextColor;

	private Label label18;

	private NumericUpDownWithUnit Num_QuietZoneWidth;

	public QRCoderSetUI()
	{
		InitializeComponent();
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionIP)
		{
			if (Globals.ThisAddIn.Application.Selection.Hyperlinks.Count > 0)
			{
				TextBox txt_QRString = Txt_QRString;
				Hyperlinks hyperlinks = Globals.ThisAddIn.Application.Selection.Hyperlinks;
				object Index = 1;
				txt_QRString.Text = hyperlinks[ref Index].Address;
			}
			else
			{
				Txt_QRString.Text = Globals.ThisAddIn.Application.Selection.Range.Text;
			}
		}
		Btn_QRDarkColor.BackColor = defaultValue.QRCodeDarkColor;
		Btn_QRLightColor.BackColor = defaultValue.QRCodeLightColor;
		Btn_IconBackGround.BackColor = defaultValue.QRCodeLightColor;
		Nud_QRModulePixel.Value = defaultValue.QRModulePixel;
		Cmb_QRECCLevel.SelectedIndex = defaultValue.QRECCLevel;
		Chk_QRQuitZone.Checked = defaultValue.QRCodeQuitZone;
		InstalledFontCollection installedFontCollection = new InstalledFontCollection();
		Cmb_TextFont.Items.Clear();
		FontFamily[] families = installedFontCollection.Families;
		foreach (FontFamily fontFamily in families)
		{
			Cmb_TextFont.Items.Add(fontFamily.Name);
			if (fontFamily.Name == "宋体")
			{
				Cmb_TextFont.SelectedItem = "宋体";
			}
		}
		Cmb_BarCodeType.SelectedIndex = 2;
	}

	private void Btn_ColorSelect_Click(object sender, EventArgs e)
	{
		ColorDialog colorDialog = new ColorDialog();
		if (colorDialog.ShowDialog() == DialogResult.OK)
		{
			(sender as Button).BackColor = colorDialog.Color;
		}
	}

	private void Btn_GetIconPath_Click(object sender, EventArgs e)
	{
		OpenFileDialog openFileDialog = new OpenFileDialog
		{
			Filter = "图片文件（BMP、JPG、GIF、PNG、ICO）|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.ico",
			Title = "选择一个图形文件",
			FileName = "IconFile"
		};
		if (openFileDialog.ShowDialog() == DialogResult.OK)
		{
			Txt_QRIconPath.Text = openFileDialog.FileName;
		}
	}

	private void Btn_CreateImage(object sender, EventArgs e)
	{
		Image image;
		if (Tab_Codes.SelectedTab.Equals(Tab_QRCode))
		{
			if (string.IsNullOrEmpty(Txt_QRString.Text))
			{
				return;
			}
			image = CreateQRCode();
		}
		else
		{
			if (string.IsNullOrEmpty(Txt_BarCodeString.Text))
			{
				return;
			}
			image = CreateBarCode();
		}
		if ((sender as Button).Name == "Btn_QRCodeInsert")
		{
			Selection selection = Globals.ThisAddIn.Application.Selection;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			selection.Collapse(ref Direction);
			Clipboard.SetDataObject(image);
			Globals.ThisAddIn.Application.Selection.Range.Paste();
			return;
		}
		Lab_ImageView.Parent = Pic_ImageView;
		Lab_ImageView.BackColor = Color.Transparent;
		Pic_ImageView.Image = image;
		if (image.Width <= 360)
		{
			Pic_ImageView.SizeMode = PictureBoxSizeMode.CenterImage;
			Lab_ImageView.Text = "图形尺寸：" + image.Width + "x" + image.Height;
			return;
		}
		Pic_ImageView.SizeMode = PictureBoxSizeMode.Zoom;
		Lab_ImageView.Text = "图形尺寸：" + image.Width + "x" + image.Height + "," + ((float)image.Width / 350f).ToString("0.00%");
	}

	private void Rdo_RectangleIcon_CheckedChanged(object sender, EventArgs e)
	{
		Num_RectangleRadius.Enabled = Rdo_RectangleIcon.Checked;
	}

	private Image CreateQRCode()
	{
		string text = Txt_QRString.Text.Trim(" \r\n".ToCharArray());
		if (text == "")
		{
			Txt_QRString.Text = "";
			return null;
		}
		Bitmap icon = null;
		if (Txt_QRIconPath.Text != "" && File.Exists(Txt_QRIconPath.Text))
		{
			icon = (Bitmap)Image.FromFile(Txt_QRIconPath.Text);
		}
		return new QRCodeCreator(new QRCodeConfig
		{
			ErrorCorrectionLevel = (QRErrorCorrectionLevel)Cmb_QRECCLevel.SelectedIndex,
			DrawQuietZone = Chk_QRQuitZone.Checked,
			Icon = icon,
			DarkColor = Btn_QRDarkColor.BackColor,
			LightColor = Btn_QRLightColor.BackColor,
			IconSize = (int)Nud_QRIconSize.Value,
			ModuleSize = (int)Nud_QRModulePixel.Value,
			IconMargin = (int)Nud_QRIconBorderWidth.Value,
			ModuleSeparate = Chk_IsolatedModule.Checked,
			RoundModule = Chk_RoundModule.Checked,
			RoundIcon = Rdo_RoundIcon.Checked,
			IconRectangleRadius = (int)Num_RectangleRadius.Value,
			IconBorderColor = Btn_IconBackGround.BackColor
		}).GenerateQrCode(text);
	}

	private Image CreateBarCode()
	{
		string text = Txt_BarCodeString.Text;
		switch (Cmb_BarCodeType.SelectedIndex)
		{
		case 0:
			if (!Regex.IsMatch(text, "^[0-9]{12}$"))
			{
				text = "000000000000";
			}
			break;
		case 1:
			if (!Regex.IsMatch(text, "^[0-9]{7}$"))
			{
				text = "0000000";
			}
			break;
		case 2:
			if (!Regex.IsMatch(text, "^[\\x20-\\x7F]+$"))
			{
				text = "ABC12345678";
			}
			break;
		}
		BarCodeConfig config = new BarCodeConfig
		{
			ModuleWidth = (int)Num_BarUnitWidth.Value,
			Height = (int)Num_BarCodeHeight.Value,
			Type = (BarCodeEncodingType)Cmb_BarCodeType.SelectedIndex,
			ShowText = Chk_ShowText.Checked,
			DarkColor = Btn_BarCodeDarkColor.BackColor,
			LightColor = Btn_BarCodeLightColor.BackColor,
			QuietZoneWidth = (int)Num_QuietZoneWidth.Value,
			FontName = Cmb_TextFont.SelectedItem.ToString(),
			TextColor = Btn_BarCodeTextColor.BackColor,
			TextHeight = (int)Num_BarCodeTextHeight.Value
		};
		return BarCodeCreator.GenerateBarCode(text, config);
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
		this.Pic_ImageView = new System.Windows.Forms.PictureBox();
		this.Txt_QRString = new System.Windows.Forms.TextBox();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.label10 = new System.Windows.Forms.Label();
		this.Btn_IconBackGround = new System.Windows.Forms.Button();
		this.Chk_RoundModule = new System.Windows.Forms.CheckBox();
		this.label9 = new System.Windows.Forms.Label();
		this.Num_RectangleRadius = new WordFormatHelper.NumericUpDownWithUnit();
		this.Rdo_RectangleIcon = new System.Windows.Forms.RadioButton();
		this.Rdo_RoundIcon = new System.Windows.Forms.RadioButton();
		this.Chk_IsolatedModule = new System.Windows.Forms.CheckBox();
		this.label7 = new System.Windows.Forms.Label();
		this.label6 = new System.Windows.Forms.Label();
		this.label5 = new System.Windows.Forms.Label();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.label1 = new System.Windows.Forms.Label();
		this.Btn_GetIconPath = new System.Windows.Forms.Button();
		this.Txt_QRIconPath = new System.Windows.Forms.TextBox();
		this.Nud_QRIconBorderWidth = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_QRIconSize = new WordFormatHelper.NumericUpDownWithUnit();
		this.Chk_QRQuitZone = new System.Windows.Forms.CheckBox();
		this.Cmb_QRECCLevel = new System.Windows.Forms.ComboBox();
		this.Nud_QRModulePixel = new WordFormatHelper.NumericUpDownWithUnit();
		this.Btn_QRLightColor = new System.Windows.Forms.Button();
		this.Btn_QRDarkColor = new System.Windows.Forms.Button();
		this.Btn_QRCodeInsert = new System.Windows.Forms.Button();
		this.Btn_QRCodeView = new System.Windows.Forms.Button();
		this.Lab_ImageView = new System.Windows.Forms.Label();
		this.Tab_Codes = new System.Windows.Forms.TabControl();
		this.Tab_QRCode = new System.Windows.Forms.TabPage();
		this.Tab_BarCode = new System.Windows.Forms.TabPage();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.label18 = new System.Windows.Forms.Label();
		this.Num_QuietZoneWidth = new WordFormatHelper.NumericUpDownWithUnit();
		this.label17 = new System.Windows.Forms.Label();
		this.Num_BarCodeTextHeight = new WordFormatHelper.NumericUpDownWithUnit();
		this.label16 = new System.Windows.Forms.Label();
		this.Btn_BarCodeTextColor = new System.Windows.Forms.Button();
		this.label15 = new System.Windows.Forms.Label();
		this.Cmb_TextFont = new System.Windows.Forms.ComboBox();
		this.label14 = new System.Windows.Forms.Label();
		this.Num_BarCodeHeight = new WordFormatHelper.NumericUpDownWithUnit();
		this.Chk_ShowText = new System.Windows.Forms.CheckBox();
		this.label8 = new System.Windows.Forms.Label();
		this.label11 = new System.Windows.Forms.Label();
		this.label12 = new System.Windows.Forms.Label();
		this.label13 = new System.Windows.Forms.Label();
		this.Cmb_BarCodeType = new System.Windows.Forms.ComboBox();
		this.Num_BarUnitWidth = new WordFormatHelper.NumericUpDownWithUnit();
		this.Btn_BarCodeLightColor = new System.Windows.Forms.Button();
		this.Btn_BarCodeDarkColor = new System.Windows.Forms.Button();
		this.Txt_BarCodeString = new System.Windows.Forms.TextBox();
		((System.ComponentModel.ISupportInitialize)this.Pic_ImageView).BeginInit();
		this.groupBox1.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_RectangleRadius).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRIconBorderWidth).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRIconSize).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRModulePixel).BeginInit();
		this.Tab_Codes.SuspendLayout();
		this.Tab_QRCode.SuspendLayout();
		this.Tab_BarCode.SuspendLayout();
		this.groupBox2.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_QuietZoneWidth).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BarCodeTextHeight).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BarCodeHeight).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BarUnitWidth).BeginInit();
		base.SuspendLayout();
		this.Pic_ImageView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
		this.Pic_ImageView.Location = new System.Drawing.Point(8, 10);
		this.Pic_ImageView.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Pic_ImageView.Name = "Pic_ImageView";
		this.Pic_ImageView.Size = new System.Drawing.Size(360, 360);
		this.Pic_ImageView.TabIndex = 0;
		this.Pic_ImageView.TabStop = false;
		this.Txt_QRString.Location = new System.Drawing.Point(6, 6);
		this.Txt_QRString.Multiline = true;
		this.Txt_QRString.Name = "Txt_QRString";
		this.Txt_QRString.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
		this.Txt_QRString.Size = new System.Drawing.Size(350, 91);
		this.Txt_QRString.TabIndex = 1;
		this.groupBox1.Controls.Add(this.label10);
		this.groupBox1.Controls.Add(this.Btn_IconBackGround);
		this.groupBox1.Controls.Add(this.Chk_RoundModule);
		this.groupBox1.Controls.Add(this.label9);
		this.groupBox1.Controls.Add(this.Num_RectangleRadius);
		this.groupBox1.Controls.Add(this.Rdo_RectangleIcon);
		this.groupBox1.Controls.Add(this.Rdo_RoundIcon);
		this.groupBox1.Controls.Add(this.Chk_IsolatedModule);
		this.groupBox1.Controls.Add(this.label7);
		this.groupBox1.Controls.Add(this.label6);
		this.groupBox1.Controls.Add(this.label5);
		this.groupBox1.Controls.Add(this.label4);
		this.groupBox1.Controls.Add(this.label3);
		this.groupBox1.Controls.Add(this.label2);
		this.groupBox1.Controls.Add(this.label1);
		this.groupBox1.Controls.Add(this.Btn_GetIconPath);
		this.groupBox1.Controls.Add(this.Txt_QRIconPath);
		this.groupBox1.Controls.Add(this.Nud_QRIconBorderWidth);
		this.groupBox1.Controls.Add(this.Nud_QRIconSize);
		this.groupBox1.Controls.Add(this.Chk_QRQuitZone);
		this.groupBox1.Controls.Add(this.Cmb_QRECCLevel);
		this.groupBox1.Controls.Add(this.Nud_QRModulePixel);
		this.groupBox1.Controls.Add(this.Btn_QRLightColor);
		this.groupBox1.Controls.Add(this.Btn_QRDarkColor);
		this.groupBox1.Location = new System.Drawing.Point(6, 103);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(350, 192);
		this.groupBox1.TabIndex = 2;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "二维码设置";
		this.label10.AutoSize = true;
		this.label10.Location = new System.Drawing.Point(278, 94);
		this.label10.Name = "label10";
		this.label10.Size = new System.Drawing.Size(37, 20);
		this.label10.TabIndex = 23;
		this.label10.Text = "背景";
		this.Btn_IconBackGround.Location = new System.Drawing.Point(314, 92);
		this.Btn_IconBackGround.Name = "Btn_IconBackGround";
		this.Btn_IconBackGround.Size = new System.Drawing.Size(25, 25);
		this.Btn_IconBackGround.TabIndex = 22;
		this.Btn_IconBackGround.UseVisualStyleBackColor = true;
		this.Btn_IconBackGround.Click += new System.EventHandler(Btn_ColorSelect_Click);
		this.Chk_RoundModule.AutoSize = true;
		this.Chk_RoundModule.Location = new System.Drawing.Point(255, 58);
		this.Chk_RoundModule.Name = "Chk_RoundModule";
		this.Chk_RoundModule.Size = new System.Drawing.Size(84, 24);
		this.Chk_RoundModule.TabIndex = 21;
		this.Chk_RoundModule.Text = "圆形码元";
		this.Chk_RoundModule.UseVisualStyleBackColor = true;
		this.label9.AutoSize = true;
		this.label9.Location = new System.Drawing.Point(183, 160);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(65, 20);
		this.label9.TabIndex = 20;
		this.label9.Text = "圆角半径";
		this.Num_RectangleRadius.Label = "像素";
		this.Num_RectangleRadius.Location = new System.Drawing.Point(251, 157);
		this.Num_RectangleRadius.Name = "Num_RectangleRadius";
		this.Num_RectangleRadius.Size = new System.Drawing.Size(87, 26);
		this.Num_RectangleRadius.TabIndex = 19;
		this.Rdo_RectangleIcon.AutoSize = true;
		this.Rdo_RectangleIcon.Checked = true;
		this.Rdo_RectangleIcon.Location = new System.Drawing.Point(102, 158);
		this.Rdo_RectangleIcon.Name = "Rdo_RectangleIcon";
		this.Rdo_RectangleIcon.Size = new System.Drawing.Size(83, 24);
		this.Rdo_RectangleIcon.TabIndex = 18;
		this.Rdo_RectangleIcon.TabStop = true;
		this.Rdo_RectangleIcon.Text = "矩形图标";
		this.Rdo_RectangleIcon.UseVisualStyleBackColor = true;
		this.Rdo_RectangleIcon.CheckedChanged += new System.EventHandler(Rdo_RectangleIcon_CheckedChanged);
		this.Rdo_RoundIcon.AutoSize = true;
		this.Rdo_RoundIcon.Location = new System.Drawing.Point(10, 158);
		this.Rdo_RoundIcon.Name = "Rdo_RoundIcon";
		this.Rdo_RoundIcon.Size = new System.Drawing.Size(83, 24);
		this.Rdo_RoundIcon.TabIndex = 17;
		this.Rdo_RoundIcon.Text = "圆形图标";
		this.Rdo_RoundIcon.UseVisualStyleBackColor = true;
		this.Chk_IsolatedModule.AutoSize = true;
		this.Chk_IsolatedModule.Location = new System.Drawing.Point(167, 58);
		this.Chk_IsolatedModule.Name = "Chk_IsolatedModule";
		this.Chk_IsolatedModule.Size = new System.Drawing.Size(84, 24);
		this.Chk_IsolatedModule.TabIndex = 16;
		this.Chk_IsolatedModule.Text = "分离码元";
		this.Chk_IsolatedModule.UseVisualStyleBackColor = true;
		this.label7.AutoSize = true;
		this.label7.Location = new System.Drawing.Point(183, 126);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(65, 20);
		this.label7.TabIndex = 15;
		this.label7.Text = "图标边框";
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(6, 126);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(65, 20);
		this.label6.TabIndex = 14;
		this.label6.Text = "图标大小";
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(6, 94);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(65, 20);
		this.label5.TabIndex = 13;
		this.label5.Text = "图标文件";
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(76, 26);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(37, 20);
		this.label4.TabIndex = 12;
		this.label4.Text = "纠错";
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(213, 26);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(37, 20);
		this.label3.TabIndex = 11;
		this.label3.Text = "大小";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(6, 60);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(37, 20);
		this.label2.TabIndex = 10;
		this.label2.Text = "浅色";
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(6, 26);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(37, 20);
		this.label1.TabIndex = 9;
		this.label1.Text = "深色";
		this.Btn_GetIconPath.Location = new System.Drawing.Point(246, 91);
		this.Btn_GetIconPath.Name = "Btn_GetIconPath";
		this.Btn_GetIconPath.Size = new System.Drawing.Size(26, 26);
		this.Btn_GetIconPath.TabIndex = 8;
		this.Btn_GetIconPath.Text = "...";
		this.Btn_GetIconPath.UseVisualStyleBackColor = true;
		this.Btn_GetIconPath.Click += new System.EventHandler(Btn_GetIconPath_Click);
		this.Txt_QRIconPath.Location = new System.Drawing.Point(77, 91);
		this.Txt_QRIconPath.Name = "Txt_QRIconPath";
		this.Txt_QRIconPath.Size = new System.Drawing.Size(171, 26);
		this.Txt_QRIconPath.TabIndex = 7;
		this.Nud_QRIconBorderWidth.Label = "像素";
		this.Nud_QRIconBorderWidth.Location = new System.Drawing.Point(251, 123);
		this.Nud_QRIconBorderWidth.Minimum = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Nud_QRIconBorderWidth.Name = "Nud_QRIconBorderWidth";
		this.Nud_QRIconBorderWidth.Size = new System.Drawing.Size(87, 26);
		this.Nud_QRIconBorderWidth.TabIndex = 6;
		this.Nud_QRIconBorderWidth.Value = new decimal(new int[4] { 6, 0, 0, 0 });
		this.Nud_QRIconSize.Label = "%";
		this.Nud_QRIconSize.Location = new System.Drawing.Point(77, 123);
		this.Nud_QRIconSize.Maximum = new decimal(new int[4] { 50, 0, 0, 0 });
		this.Nud_QRIconSize.Minimum = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Nud_QRIconSize.Name = "Nud_QRIconSize";
		this.Nud_QRIconSize.Size = new System.Drawing.Size(88, 26);
		this.Nud_QRIconSize.TabIndex = 5;
		this.Nud_QRIconSize.Value = new decimal(new int[4] { 15, 0, 0, 0 });
		this.Chk_QRQuitZone.AutoSize = true;
		this.Chk_QRQuitZone.Location = new System.Drawing.Point(80, 58);
		this.Chk_QRQuitZone.Name = "Chk_QRQuitZone";
		this.Chk_QRQuitZone.Size = new System.Drawing.Size(84, 24);
		this.Chk_QRQuitZone.TabIndex = 4;
		this.Chk_QRQuitZone.Text = "创建边框";
		this.Chk_QRQuitZone.UseVisualStyleBackColor = true;
		this.Cmb_QRECCLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_QRECCLevel.FormattingEnabled = true;
		this.Cmb_QRECCLevel.Items.AddRange(new object[4] { "L-低", "M-中", "Q-较高", "H-高" });
		this.Cmb_QRECCLevel.Location = new System.Drawing.Point(116, 22);
		this.Cmb_QRECCLevel.Name = "Cmb_QRECCLevel";
		this.Cmb_QRECCLevel.Size = new System.Drawing.Size(88, 28);
		this.Cmb_QRECCLevel.TabIndex = 3;
		this.Nud_QRModulePixel.Label = "像素";
		this.Nud_QRModulePixel.Location = new System.Drawing.Point(252, 23);
		this.Nud_QRModulePixel.Minimum = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Nud_QRModulePixel.Name = "Nud_QRModulePixel";
		this.Nud_QRModulePixel.Size = new System.Drawing.Size(87, 26);
		this.Nud_QRModulePixel.TabIndex = 2;
		this.Nud_QRModulePixel.Value = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Btn_QRLightColor.Location = new System.Drawing.Point(42, 58);
		this.Btn_QRLightColor.Name = "Btn_QRLightColor";
		this.Btn_QRLightColor.Size = new System.Drawing.Size(25, 25);
		this.Btn_QRLightColor.TabIndex = 1;
		this.Btn_QRLightColor.UseVisualStyleBackColor = true;
		this.Btn_QRLightColor.Click += new System.EventHandler(Btn_ColorSelect_Click);
		this.Btn_QRDarkColor.Location = new System.Drawing.Point(42, 24);
		this.Btn_QRDarkColor.Name = "Btn_QRDarkColor";
		this.Btn_QRDarkColor.Size = new System.Drawing.Size(25, 25);
		this.Btn_QRDarkColor.TabIndex = 0;
		this.Btn_QRDarkColor.UseVisualStyleBackColor = true;
		this.Btn_QRDarkColor.Click += new System.EventHandler(Btn_ColorSelect_Click);
		this.Btn_QRCodeInsert.Location = new System.Drawing.Point(589, 340);
		this.Btn_QRCodeInsert.Name = "Btn_QRCodeInsert";
		this.Btn_QRCodeInsert.Size = new System.Drawing.Size(152, 30);
		this.Btn_QRCodeInsert.TabIndex = 17;
		this.Btn_QRCodeInsert.Text = "插入文中";
		this.Btn_QRCodeInsert.UseVisualStyleBackColor = true;
		this.Btn_QRCodeInsert.Click += new System.EventHandler(Btn_CreateImage);
		this.Btn_QRCodeView.Location = new System.Drawing.Point(375, 340);
		this.Btn_QRCodeView.Name = "Btn_QRCodeView";
		this.Btn_QRCodeView.Size = new System.Drawing.Size(152, 30);
		this.Btn_QRCodeView.TabIndex = 18;
		this.Btn_QRCodeView.Text = "预览";
		this.Btn_QRCodeView.UseVisualStyleBackColor = true;
		this.Btn_QRCodeView.Click += new System.EventHandler(Btn_CreateImage);
		this.Lab_ImageView.Location = new System.Drawing.Point(144, 340);
		this.Lab_ImageView.Name = "Lab_ImageView";
		this.Lab_ImageView.Size = new System.Drawing.Size(214, 20);
		this.Lab_ImageView.TabIndex = 19;
		this.Lab_ImageView.Text = "图形预览";
		this.Lab_ImageView.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.Tab_Codes.Controls.Add(this.Tab_QRCode);
		this.Tab_Codes.Controls.Add(this.Tab_BarCode);
		this.Tab_Codes.Location = new System.Drawing.Point(374, 5);
		this.Tab_Codes.Name = "Tab_Codes";
		this.Tab_Codes.SelectedIndex = 0;
		this.Tab_Codes.Size = new System.Drawing.Size(371, 333);
		this.Tab_Codes.TabIndex = 20;
		this.Tab_QRCode.BackColor = System.Drawing.Color.AliceBlue;
		this.Tab_QRCode.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
		this.Tab_QRCode.Controls.Add(this.Txt_QRString);
		this.Tab_QRCode.Controls.Add(this.groupBox1);
		this.Tab_QRCode.Location = new System.Drawing.Point(4, 29);
		this.Tab_QRCode.Name = "Tab_QRCode";
		this.Tab_QRCode.Padding = new System.Windows.Forms.Padding(3);
		this.Tab_QRCode.Size = new System.Drawing.Size(363, 300);
		this.Tab_QRCode.TabIndex = 0;
		this.Tab_QRCode.Text = "二维码";
		this.Tab_BarCode.BackColor = System.Drawing.Color.AliceBlue;
		this.Tab_BarCode.Controls.Add(this.groupBox2);
		this.Tab_BarCode.Controls.Add(this.Txt_BarCodeString);
		this.Tab_BarCode.Location = new System.Drawing.Point(4, 29);
		this.Tab_BarCode.Name = "Tab_BarCode";
		this.Tab_BarCode.Padding = new System.Windows.Forms.Padding(3);
		this.Tab_BarCode.Size = new System.Drawing.Size(363, 300);
		this.Tab_BarCode.TabIndex = 1;
		this.Tab_BarCode.Text = "条形码";
		this.groupBox2.Controls.Add(this.label18);
		this.groupBox2.Controls.Add(this.Num_QuietZoneWidth);
		this.groupBox2.Controls.Add(this.label17);
		this.groupBox2.Controls.Add(this.Num_BarCodeTextHeight);
		this.groupBox2.Controls.Add(this.label16);
		this.groupBox2.Controls.Add(this.Btn_BarCodeTextColor);
		this.groupBox2.Controls.Add(this.label15);
		this.groupBox2.Controls.Add(this.Cmb_TextFont);
		this.groupBox2.Controls.Add(this.label14);
		this.groupBox2.Controls.Add(this.Num_BarCodeHeight);
		this.groupBox2.Controls.Add(this.Chk_ShowText);
		this.groupBox2.Controls.Add(this.label8);
		this.groupBox2.Controls.Add(this.label11);
		this.groupBox2.Controls.Add(this.label12);
		this.groupBox2.Controls.Add(this.label13);
		this.groupBox2.Controls.Add(this.Cmb_BarCodeType);
		this.groupBox2.Controls.Add(this.Num_BarUnitWidth);
		this.groupBox2.Controls.Add(this.Btn_BarCodeLightColor);
		this.groupBox2.Controls.Add(this.Btn_BarCodeDarkColor);
		this.groupBox2.Location = new System.Drawing.Point(6, 38);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.Size = new System.Drawing.Size(350, 256);
		this.groupBox2.TabIndex = 3;
		this.groupBox2.TabStop = false;
		this.groupBox2.Text = "条形码设置";
		this.label18.AutoSize = true;
		this.label18.Location = new System.Drawing.Point(14, 104);
		this.label18.Name = "label18";
		this.label18.Size = new System.Drawing.Size(65, 20);
		this.label18.TabIndex = 31;
		this.label18.Text = "两侧白边";
		this.Num_QuietZoneWidth.Label = "倍单位宽度";
		this.Num_QuietZoneWidth.Location = new System.Drawing.Point(85, 101);
		this.Num_QuietZoneWidth.Maximum = new decimal(new int[4] { 1000, 0, 0, 0 });
		this.Num_QuietZoneWidth.Minimum = new decimal(new int[4] { 10, 0, 0, 0 });
		this.Num_QuietZoneWidth.Name = "Num_QuietZoneWidth";
		this.Num_QuietZoneWidth.Size = new System.Drawing.Size(144, 26);
		this.Num_QuietZoneWidth.TabIndex = 30;
		this.Num_QuietZoneWidth.Value = new decimal(new int[4] { 10, 0, 0, 0 });
		this.label17.AutoSize = true;
		this.label17.Location = new System.Drawing.Point(181, 213);
		this.label17.Name = "label17";
		this.label17.Size = new System.Drawing.Size(65, 20);
		this.label17.TabIndex = 29;
		this.label17.Text = "文字高度";
		this.Num_BarCodeTextHeight.Label = "像素";
		this.Num_BarCodeTextHeight.Location = new System.Drawing.Point(252, 210);
		this.Num_BarCodeTextHeight.Minimum = new decimal(new int[4] { 5, 0, 0, 0 });
		this.Num_BarCodeTextHeight.Name = "Num_BarCodeTextHeight";
		this.Num_BarCodeTextHeight.Size = new System.Drawing.Size(87, 26);
		this.Num_BarCodeTextHeight.TabIndex = 28;
		this.Num_BarCodeTextHeight.Value = new decimal(new int[4] { 5, 0, 0, 0 });
		this.label16.AutoSize = true;
		this.label16.Location = new System.Drawing.Point(37, 213);
		this.label16.Name = "label16";
		this.label16.Size = new System.Drawing.Size(65, 20);
		this.label16.TabIndex = 27;
		this.label16.Text = "文字颜色";
		this.Btn_BarCodeTextColor.BackColor = System.Drawing.Color.Black;
		this.Btn_BarCodeTextColor.Location = new System.Drawing.Point(117, 211);
		this.Btn_BarCodeTextColor.Name = "Btn_BarCodeTextColor";
		this.Btn_BarCodeTextColor.Size = new System.Drawing.Size(25, 25);
		this.Btn_BarCodeTextColor.TabIndex = 26;
		this.Btn_BarCodeTextColor.UseVisualStyleBackColor = false;
		this.Btn_BarCodeTextColor.Click += new System.EventHandler(Btn_ColorSelect_Click);
		this.label15.AutoSize = true;
		this.label15.Location = new System.Drawing.Point(37, 178);
		this.label15.Name = "label15";
		this.label15.Size = new System.Drawing.Size(65, 20);
		this.label15.TabIndex = 25;
		this.label15.Text = "文字字体";
		this.Cmb_TextFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TextFont.FormattingEnabled = true;
		this.Cmb_TextFont.Items.AddRange(new object[3] { "EAN13", "EAN8", "Code128B" });
		this.Cmb_TextFont.Location = new System.Drawing.Point(117, 174);
		this.Cmb_TextFont.Name = "Cmb_TextFont";
		this.Cmb_TextFont.Size = new System.Drawing.Size(222, 28);
		this.Cmb_TextFont.TabIndex = 24;
		this.label14.AutoSize = true;
		this.label14.Location = new System.Drawing.Point(184, 69);
		this.label14.Name = "label14";
		this.label14.Size = new System.Drawing.Size(65, 20);
		this.label14.TabIndex = 23;
		this.label14.Text = "条码高度";
		this.Num_BarCodeHeight.Label = "像素";
		this.Num_BarCodeHeight.Location = new System.Drawing.Point(252, 66);
		this.Num_BarCodeHeight.Maximum = new decimal(new int[4] { 1000, 0, 0, 0 });
		this.Num_BarCodeHeight.Minimum = new decimal(new int[4] { 30, 0, 0, 0 });
		this.Num_BarCodeHeight.Name = "Num_BarCodeHeight";
		this.Num_BarCodeHeight.Size = new System.Drawing.Size(87, 26);
		this.Num_BarCodeHeight.TabIndex = 22;
		this.Num_BarCodeHeight.Value = new decimal(new int[4] { 150, 0, 0, 0 });
		this.Chk_ShowText.AutoSize = true;
		this.Chk_ShowText.Location = new System.Drawing.Point(18, 139);
		this.Chk_ShowText.Name = "Chk_ShowText";
		this.Chk_ShowText.Size = new System.Drawing.Size(84, 24);
		this.Chk_ShowText.TabIndex = 21;
		this.Chk_ShowText.Text = "显示文字";
		this.Chk_ShowText.UseVisualStyleBackColor = true;
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(14, 34);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(65, 20);
		this.label8.TabIndex = 20;
		this.label8.Text = "编码类型";
		this.label11.AutoSize = true;
		this.label11.Location = new System.Drawing.Point(184, 34);
		this.label11.Name = "label11";
		this.label11.Size = new System.Drawing.Size(65, 20);
		this.label11.TabIndex = 19;
		this.label11.Text = "单元宽度";
		this.label12.AutoSize = true;
		this.label12.Location = new System.Drawing.Point(81, 69);
		this.label12.Name = "label12";
		this.label12.Size = new System.Drawing.Size(37, 20);
		this.label12.TabIndex = 18;
		this.label12.Text = "浅色";
		this.label13.AutoSize = true;
		this.label13.Location = new System.Drawing.Point(14, 69);
		this.label13.Name = "label13";
		this.label13.Size = new System.Drawing.Size(37, 20);
		this.label13.TabIndex = 17;
		this.label13.Text = "深色";
		this.Cmb_BarCodeType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_BarCodeType.FormattingEnabled = true;
		this.Cmb_BarCodeType.Items.AddRange(new object[3] { "EAN13", "EAN8", "Code128B" });
		this.Cmb_BarCodeType.Location = new System.Drawing.Point(85, 30);
		this.Cmb_BarCodeType.Name = "Cmb_BarCodeType";
		this.Cmb_BarCodeType.Size = new System.Drawing.Size(95, 28);
		this.Cmb_BarCodeType.TabIndex = 16;
		this.Num_BarUnitWidth.Label = "像素";
		this.Num_BarUnitWidth.Location = new System.Drawing.Point(252, 31);
		this.Num_BarUnitWidth.Maximum = new decimal(new int[4] { 10, 0, 0, 0 });
		this.Num_BarUnitWidth.Minimum = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Num_BarUnitWidth.Name = "Num_BarUnitWidth";
		this.Num_BarUnitWidth.Size = new System.Drawing.Size(87, 26);
		this.Num_BarUnitWidth.TabIndex = 15;
		this.Num_BarUnitWidth.Value = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Btn_BarCodeLightColor.BackColor = System.Drawing.Color.White;
		this.Btn_BarCodeLightColor.Location = new System.Drawing.Point(117, 67);
		this.Btn_BarCodeLightColor.Name = "Btn_BarCodeLightColor";
		this.Btn_BarCodeLightColor.Size = new System.Drawing.Size(25, 25);
		this.Btn_BarCodeLightColor.TabIndex = 14;
		this.Btn_BarCodeLightColor.UseVisualStyleBackColor = false;
		this.Btn_BarCodeLightColor.Click += new System.EventHandler(Btn_ColorSelect_Click);
		this.Btn_BarCodeDarkColor.BackColor = System.Drawing.Color.Black;
		this.Btn_BarCodeDarkColor.Location = new System.Drawing.Point(50, 67);
		this.Btn_BarCodeDarkColor.Name = "Btn_BarCodeDarkColor";
		this.Btn_BarCodeDarkColor.Size = new System.Drawing.Size(25, 25);
		this.Btn_BarCodeDarkColor.TabIndex = 13;
		this.Btn_BarCodeDarkColor.UseVisualStyleBackColor = false;
		this.Btn_BarCodeDarkColor.Click += new System.EventHandler(Btn_ColorSelect_Click);
		this.Txt_BarCodeString.Location = new System.Drawing.Point(6, 6);
		this.Txt_BarCodeString.Name = "Txt_BarCodeString";
		this.Txt_BarCodeString.Size = new System.Drawing.Size(350, 26);
		this.Txt_BarCodeString.TabIndex = 2;
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Tab_Codes);
		base.Controls.Add(this.Lab_ImageView);
		base.Controls.Add(this.Btn_QRCodeView);
		base.Controls.Add(this.Btn_QRCodeInsert);
		base.Controls.Add(this.Pic_ImageView);
		this.DoubleBuffered = true;
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		base.Name = "QRCoderSetUI";
		base.Size = new System.Drawing.Size(748, 378);
		((System.ComponentModel.ISupportInitialize)this.Pic_ImageView).EndInit();
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_RectangleRadius).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRIconBorderWidth).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRIconSize).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_QRModulePixel).EndInit();
		this.Tab_Codes.ResumeLayout(false);
		this.Tab_QRCode.ResumeLayout(false);
		this.Tab_QRCode.PerformLayout();
		this.Tab_BarCode.ResumeLayout(false);
		this.Tab_BarCode.PerformLayout();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_QuietZoneWidth).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BarCodeTextHeight).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BarCodeHeight).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BarUnitWidth).EndInit();
		base.ResumeLayout(false);
	}
}
}