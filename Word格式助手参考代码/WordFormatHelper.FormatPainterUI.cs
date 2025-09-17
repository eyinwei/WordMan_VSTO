// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.FormatPainterUI
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
using WordFormatHelper;

public class FormatPainterUI : UserControl
{
	private bool isInitializing;

	private readonly Dictionary<float, string> ChnFontSize = new Dictionary<float, string>();

	private IContainer components;

	private GroupBox Grp_FromatSetting;

	private ListBox Lst_FormatPainterList;

	private TextBox Txt_FormatDiscription;

	private TextBox Txt_FormatPainterName;

	private ComboBox Cmb_EngFontName;

	private Label label4;

	private Label label3;

	private ComboBox Cmb_ChnFontName;

	private Label label2;

	private Label label1;

	private ComboBox Cmb_FontSize;

	private Label label5;

	private ToggleButton Btn_FontUnerline;

	private ToggleButton Btn_FontItalic;

	private ToggleButton Btn_FontBold;

	private Button Btn_ShadingColor;

	private Button Btn_TextColor;

	private CheckBox Chk_FontShading;

	private Label Lab_ChnFontShow;

	private Label Lab_EngFontShow;

	private Button Btn_Apply;

	private Button Btn_New;

	private Button Btn_Modify;

	private CheckBox Chk_TextColor;

	private Button Btn_Delete;

	public FormatPainterUI()
	{
		isInitializing = true;
		InitializeComponent();
		InstalledFontCollection installedFontCollection = new InstalledFontCollection();
		Cmb_ChnFontName.Items.Clear();
		Cmb_EngFontName.Items.Clear();
		Cmb_FontSize.Items.Clear();
		Cmb_ChnFontName.Items.AddRange(((IEnumerable<object>)installedFontCollection.Families.Select((FontFamily item) => item.Name)).ToArray());
		Cmb_EngFontName.Items.AddRange(((IEnumerable<object>)installedFontCollection.Families.Select((FontFamily item) => item.Name)).ToArray());
		ChnFontSize.Add(42f, "初号");
		ChnFontSize.Add(36f, "小初");
		ChnFontSize.Add(26f, "一号");
		ChnFontSize.Add(24f, "小一");
		ChnFontSize.Add(22f, "二号");
		ChnFontSize.Add(18f, "小二");
		ChnFontSize.Add(16f, "三号");
		ChnFontSize.Add(15f, "小三");
		ChnFontSize.Add(14f, "四号");
		ChnFontSize.Add(12f, "小四");
		ChnFontSize.Add(10.5f, "五号");
		ChnFontSize.Add(9f, "小五");
		ChnFontSize.Add(7.5f, "六号");
		ChnFontSize.Add(6.5f, "小六");
		ChnFontSize.Add(5.5f, "七号");
		ChnFontSize.Add(5f, "八号");
		Cmb_FontSize.Items.AddRange(((IEnumerable<object>)ChnFontSize.Select((KeyValuePair<float, string> item) => item.Value)).ToArray());
		Lst_FormatPainterList.Items.Clear();
		Lst_FormatPainterList.Items.AddRange(((IEnumerable<object>)ThisAddIn.formatPainter.StoredFormat.Select((FixFormatPainterSetting.FixFormat item) => item.StyleName)).ToArray());
		Lst_FormatPainterList.SelectedIndex = ThisAddIn.formatPainter.CurrentID;
		Lst_FormatPainterList_SelectedIndexChanged(null, null);
		Btn_Delete.Enabled = Lst_FormatPainterList.Items.Count > 1;
		isInitializing = false;
	}

	private void Lst_FormatPainterList_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (Lst_FormatPainterList.SelectedIndex != -1)
		{
			isInitializing = true;
			FixFormatPainterSetting.FixFormat fixFormat = ThisAddIn.formatPainter.StoredFormat[Lst_FormatPainterList.SelectedIndex];
			Txt_FormatPainterName.Text = fixFormat.StyleName;
			Txt_FormatDiscription.Text = fixFormat.Discription;
			Cmb_ChnFontName.Text = fixFormat.ChnFontName;
			Cmb_EngFontName.Text = fixFormat.EngFontName;
			if (ChnFontSize.ContainsKey(fixFormat.FontSize))
			{
				Cmb_FontSize.Text = ChnFontSize[fixFormat.FontSize];
			}
			else
			{
				Cmb_FontSize.Text = fixFormat.FontSize.ToString();
			}
			Btn_FontBold.Pressed = fixFormat.Bold;
			Btn_FontItalic.Pressed = fixFormat.Italic;
			Btn_FontUnerline.Pressed = fixFormat.Underline;
			Chk_TextColor.Checked = fixFormat.UseColor;
			if (fixFormat.UseColor)
			{
				Btn_TextColor.BackColor = Color.FromArgb(fixFormat.TextColor);
			}
			Chk_FontShading.Checked = fixFormat.Shading;
			if (fixFormat.Shading)
			{
				Btn_ShadingColor.BackColor = Color.FromArgb(fixFormat.ShadingColor);
			}
			isInitializing = false;
			UpdateFontShowLable();
			Btn_Modify.Enabled = false;
			ThisAddIn.formatPainter.CurrentID = Lst_FormatPainterList.SelectedIndex;
		}
	}

	private void UpdateFontShowLable()
	{
		if (!isInitializing)
		{
			float emSize = ((Cmb_FontSize.SelectedIndex == -1) ? Convert.ToSingle(Cmb_FontSize.Text) : ChnFontSize.Where((KeyValuePair<float, string> item) => item.Value == Cmb_FontSize.SelectedItem.ToString()).First().Key);
			FontStyle fontStyle = FontStyle.Regular;
			if (Btn_FontBold.Pressed)
			{
				fontStyle |= FontStyle.Bold;
			}
			if (Btn_FontItalic.Pressed)
			{
				fontStyle |= FontStyle.Italic;
			}
			if (Btn_FontUnerline.Pressed)
			{
				fontStyle |= FontStyle.Underline;
			}
			Font font = new Font(new FontFamily(Cmb_ChnFontName.Text), emSize, fontStyle);
			Lab_ChnFontShow.Font = font;
			font = new Font(new FontFamily(Cmb_EngFontName.Text), emSize, fontStyle);
			Lab_EngFontShow.Font = font;
			if (Chk_TextColor.Checked)
			{
				Lab_ChnFontShow.ForeColor = Btn_TextColor.BackColor;
				Lab_EngFontShow.ForeColor = Btn_TextColor.BackColor;
			}
			else
			{
				Lab_ChnFontShow.ForeColor = Color.FromKnownColor(KnownColor.ControlText);
				Lab_EngFontShow.ForeColor = Color.FromKnownColor(KnownColor.ControlText);
			}
			if (Chk_FontShading.Checked)
			{
				Lab_ChnFontShow.BackColor = Btn_ShadingColor.BackColor;
				Lab_EngFontShow.BackColor = Btn_ShadingColor.BackColor;
			}
			else
			{
				Lab_ChnFontShow.BackColor = BackColor;
				Lab_EngFontShow.BackColor = BackColor;
			}
		}
	}

	private void Btn_Modify_Click(object sender, EventArgs e)
	{
		if (string.IsNullOrEmpty(Txt_FormatPainterName.Text))
		{
			MessageBox.Show("格式名称不能为空！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
		else
		{
			if (Lst_FormatPainterList.SelectedIndex == -1)
			{
				return;
			}
			FixFormatPainterSetting.FixFormat value = new FixFormatPainterSetting.FixFormat
			{
				Id = Lst_FormatPainterList.SelectedIndex,
				StyleName = Txt_FormatPainterName.Text,
				Discription = Txt_FormatDiscription.Text,
				ChnFontName = Cmb_ChnFontName.Text,
				EngFontName = Cmb_EngFontName.Text,
				Bold = Btn_FontBold.Pressed,
				Italic = Btn_FontItalic.Pressed,
				Underline = Btn_FontUnerline.Pressed,
				UseColor = Chk_TextColor.Checked,
				TextColor = Btn_TextColor.BackColor.ToArgb(),
				Shading = Chk_FontShading.Checked,
				ShadingColor = Btn_ShadingColor.BackColor.ToArgb()
			};
			if (Cmb_FontSize.SelectedIndex != -1)
			{
				value.FontSize = ChnFontSize.Where((KeyValuePair<float, string> item) => item.Value == Cmb_FontSize.Text).First().Key;
			}
			else
			{
				value.FontSize = Convert.ToSingle(Cmb_FontSize.Text);
			}
			ThisAddIn.formatPainter.StoredFormat[Lst_FormatPainterList.SelectedIndex] = value;
			Btn_Modify.Enabled = false;
		}
	}

	private void Btn_New_Click(object sender, EventArgs e)
	{
		if (string.IsNullOrEmpty(Txt_FormatPainterName.Text))
		{
			MessageBox.Show("格式名称不能为空！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			return;
		}
		FixFormatPainterSetting.FixFormat item = new FixFormatPainterSetting.FixFormat
		{
			Id = Lst_FormatPainterList.Items.Count,
			StyleName = Txt_FormatPainterName.Text,
			Discription = Txt_FormatDiscription.Text,
			ChnFontName = Cmb_ChnFontName.Text,
			EngFontName = Cmb_EngFontName.Text,
			Bold = Btn_FontBold.Pressed,
			Italic = Btn_FontItalic.Pressed,
			Underline = Btn_FontUnerline.Pressed,
			UseColor = Chk_TextColor.Checked,
			TextColor = Btn_TextColor.BackColor.ToArgb(),
			Shading = Chk_FontShading.Checked,
			ShadingColor = Btn_ShadingColor.BackColor.ToArgb()
		};
		if (Cmb_FontSize.SelectedIndex != -1)
		{
			item.FontSize = ChnFontSize.Where((KeyValuePair<float, string> keyValuePair) => keyValuePair.Value == Cmb_FontSize.Text).First().Key;
		}
		else
		{
			item.FontSize = Convert.ToSingle(Cmb_FontSize.Text);
		}
		ThisAddIn.formatPainter.StoredFormat.Add(item);
		Lst_FormatPainterList.Items.Add(item.StyleName);
		Lst_FormatPainterList.SelectedIndex = Lst_FormatPainterList.Items.Count - 1;
		if (Lst_FormatPainterList.Items.Count > 1)
		{
			Btn_Delete.Enabled = true;
		}
	}

	private void Cmb_ChnFontName_SelectedIndexChanged(object sender, EventArgs e)
	{
		Btn_Modify.Enabled = true;
		UpdateFontShowLable();
	}

	private void Btn_FontBold_Click(object sender, EventArgs e)
	{
		Btn_Modify.Enabled = true;
		UpdateFontShowLable();
	}

	private void Chk_TextColor_CheckedChanged(object sender, EventArgs e)
	{
		if ((sender as CheckBox).Name == "Chk_TextColor")
		{
			Btn_TextColor.Enabled = Chk_TextColor.Checked;
		}
		else
		{
			Btn_ShadingColor.Enabled = Chk_FontShading.Checked;
		}
		Btn_Modify.Enabled = true;
		UpdateFontShowLable();
	}

	private void Btn_TextColor_BackColorChanged(object sender, EventArgs e)
	{
		Btn_Modify.Enabled = true;
		UpdateFontShowLable();
	}

	private void Btn_TextColor_Click(object sender, EventArgs e)
	{
		ColorDialog colorDialog = new ColorDialog
		{
			SolidColorOnly = true,
			AllowFullOpen = true,
			AnyColor = true
		};
		if (colorDialog.ShowDialog() == DialogResult.OK)
		{
			if ((sender as Button).Name == "Btn_TextColor")
			{
				Btn_TextColor.BackColor = colorDialog.Color;
			}
			else
			{
				Btn_ShadingColor.BackColor = colorDialog.Color;
			}
			Btn_Modify.Enabled = true;
			UpdateFontShowLable();
		}
	}

	private void Txt_FormatPainterName_TextChanged(object sender, EventArgs e)
	{
		Btn_Modify.Enabled = true;
	}

	private void Cmb_FontSize_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (Cmb_FontSize.SelectedIndex == -1)
		{
			try
			{
				float num = Convert.ToSingle(Cmb_FontSize.Text);
				if (num < 1f)
				{
					num = 1f;
				}
				else
				{
					num = (float)Math.Round(num * 2f) / 2f;
					if (ChnFontSize.ContainsKey(num))
					{
						Cmb_FontSize.Text = ChnFontSize[num];
					}
					else
					{
						Cmb_FontSize.Text = num.ToString();
					}
				}
			}
			catch
			{
				Cmb_FontSize.SelectedIndex = 10;
			}
		}
		Btn_Modify.Enabled = true;
		UpdateFontShowLable();
	}

	private void Btn_Apply_Click(object sender, EventArgs e)
	{
		Globals.Ribbons.WordFormatHelperRibbon.ApplyFixFormatPainter();
	}

	private void Btn_Delete_Click(object sender, EventArgs e)
	{
		ThisAddIn.formatPainter.StoredFormat.RemoveAt(ThisAddIn.formatPainter.CurrentID);
		if (ThisAddIn.formatPainter.CurrentID <= ThisAddIn.formatPainter.StoredFormat.Count - 1)
		{
			for (int i = ThisAddIn.formatPainter.CurrentID; i <= ThisAddIn.formatPainter.StoredFormat.Count - 1; i++)
			{
				FixFormatPainterSetting.FixFormat value = ThisAddIn.formatPainter.StoredFormat[i];
				value.Id = i;
				ThisAddIn.formatPainter.StoredFormat[i] = value;
			}
			Lst_FormatPainterList.Items.RemoveAt(ThisAddIn.formatPainter.CurrentID);
			Lst_FormatPainterList.SelectedIndex = ThisAddIn.formatPainter.CurrentID;
		}
		else
		{
			Lst_FormatPainterList.Items.RemoveAt(ThisAddIn.formatPainter.CurrentID);
			Lst_FormatPainterList.SelectedIndex = ThisAddIn.formatPainter.StoredFormat.Count - 1;
		}
		Btn_Delete.Enabled = ThisAddIn.formatPainter.StoredFormat.Count > 1;
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
		this.Grp_FromatSetting = new System.Windows.Forms.GroupBox();
		this.Btn_ShadingColor = new System.Windows.Forms.Button();
		this.Btn_TextColor = new System.Windows.Forms.Button();
		this.Chk_TextColor = new System.Windows.Forms.CheckBox();
		this.Btn_New = new System.Windows.Forms.Button();
		this.Btn_Modify = new System.Windows.Forms.Button();
		this.Lab_EngFontShow = new System.Windows.Forms.Label();
		this.Lab_ChnFontShow = new System.Windows.Forms.Label();
		this.Chk_FontShading = new System.Windows.Forms.CheckBox();
		this.Cmb_FontSize = new System.Windows.Forms.ComboBox();
		this.label5 = new System.Windows.Forms.Label();
		this.Txt_FormatDiscription = new System.Windows.Forms.TextBox();
		this.Txt_FormatPainterName = new System.Windows.Forms.TextBox();
		this.Cmb_EngFontName = new System.Windows.Forms.ComboBox();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.Cmb_ChnFontName = new System.Windows.Forms.ComboBox();
		this.label2 = new System.Windows.Forms.Label();
		this.label1 = new System.Windows.Forms.Label();
		this.Lst_FormatPainterList = new System.Windows.Forms.ListBox();
		this.Btn_Apply = new System.Windows.Forms.Button();
		this.Btn_Delete = new System.Windows.Forms.Button();
		this.Btn_FontUnerline = new WordFormatHelper.ToggleButton();
		this.Btn_FontItalic = new WordFormatHelper.ToggleButton();
		this.Btn_FontBold = new WordFormatHelper.ToggleButton();
		this.Grp_FromatSetting.SuspendLayout();
		base.SuspendLayout();
		this.Grp_FromatSetting.Controls.Add(this.Btn_Delete);
		this.Grp_FromatSetting.Controls.Add(this.Btn_ShadingColor);
		this.Grp_FromatSetting.Controls.Add(this.Btn_TextColor);
		this.Grp_FromatSetting.Controls.Add(this.Chk_TextColor);
		this.Grp_FromatSetting.Controls.Add(this.Btn_New);
		this.Grp_FromatSetting.Controls.Add(this.Btn_Modify);
		this.Grp_FromatSetting.Controls.Add(this.Lab_EngFontShow);
		this.Grp_FromatSetting.Controls.Add(this.Lab_ChnFontShow);
		this.Grp_FromatSetting.Controls.Add(this.Chk_FontShading);
		this.Grp_FromatSetting.Controls.Add(this.Cmb_FontSize);
		this.Grp_FromatSetting.Controls.Add(this.label5);
		this.Grp_FromatSetting.Controls.Add(this.Btn_FontUnerline);
		this.Grp_FromatSetting.Controls.Add(this.Btn_FontItalic);
		this.Grp_FromatSetting.Controls.Add(this.Btn_FontBold);
		this.Grp_FromatSetting.Controls.Add(this.Txt_FormatDiscription);
		this.Grp_FromatSetting.Controls.Add(this.Txt_FormatPainterName);
		this.Grp_FromatSetting.Controls.Add(this.Cmb_EngFontName);
		this.Grp_FromatSetting.Controls.Add(this.label4);
		this.Grp_FromatSetting.Controls.Add(this.label3);
		this.Grp_FromatSetting.Controls.Add(this.Cmb_ChnFontName);
		this.Grp_FromatSetting.Controls.Add(this.label2);
		this.Grp_FromatSetting.Controls.Add(this.label1);
		this.Grp_FromatSetting.Location = new System.Drawing.Point(174, 2);
		this.Grp_FromatSetting.Name = "Grp_FromatSetting";
		this.Grp_FromatSetting.Size = new System.Drawing.Size(265, 310);
		this.Grp_FromatSetting.TabIndex = 0;
		this.Grp_FromatSetting.TabStop = false;
		this.Grp_FromatSetting.Text = "格式设置";
		this.Btn_ShadingColor.Location = new System.Drawing.Point(228, 201);
		this.Btn_ShadingColor.Name = "Btn_ShadingColor";
		this.Btn_ShadingColor.Size = new System.Drawing.Size(30, 30);
		this.Btn_ShadingColor.TabIndex = 16;
		this.Btn_ShadingColor.UseVisualStyleBackColor = true;
		this.Btn_ShadingColor.BackColorChanged += new System.EventHandler(Btn_TextColor_BackColorChanged);
		this.Btn_ShadingColor.Click += new System.EventHandler(Btn_TextColor_Click);
		this.Btn_TextColor.Location = new System.Drawing.Point(98, 201);
		this.Btn_TextColor.Name = "Btn_TextColor";
		this.Btn_TextColor.Size = new System.Drawing.Size(30, 30);
		this.Btn_TextColor.TabIndex = 15;
		this.Btn_TextColor.UseVisualStyleBackColor = true;
		this.Btn_TextColor.BackColorChanged += new System.EventHandler(Btn_TextColor_BackColorChanged);
		this.Btn_TextColor.Click += new System.EventHandler(Btn_TextColor_Click);
		this.Chk_TextColor.AutoSize = true;
		this.Chk_TextColor.Location = new System.Drawing.Point(18, 206);
		this.Chk_TextColor.Name = "Chk_TextColor";
		this.Chk_TextColor.Size = new System.Drawing.Size(83, 22);
		this.Chk_TextColor.TabIndex = 21;
		this.Chk_TextColor.Text = "字体颜色";
		this.Chk_TextColor.UseVisualStyleBackColor = true;
		this.Chk_TextColor.CheckedChanged += new System.EventHandler(Chk_TextColor_CheckedChanged);
		this.Btn_New.Location = new System.Drawing.Point(13, 272);
		this.Btn_New.Name = "Btn_New";
		this.Btn_New.Size = new System.Drawing.Size(80, 30);
		this.Btn_New.TabIndex = 20;
		this.Btn_New.Text = "新建格式";
		this.Btn_New.UseVisualStyleBackColor = true;
		this.Btn_New.Click += new System.EventHandler(Btn_New_Click);
		this.Btn_Modify.Enabled = false;
		this.Btn_Modify.Location = new System.Drawing.Point(94, 272);
		this.Btn_Modify.Name = "Btn_Modify";
		this.Btn_Modify.Size = new System.Drawing.Size(80, 30);
		this.Btn_Modify.TabIndex = 19;
		this.Btn_Modify.Text = "修改格式";
		this.Btn_Modify.UseVisualStyleBackColor = true;
		this.Btn_Modify.Click += new System.EventHandler(Btn_Modify_Click);
		this.Lab_EngFontShow.Location = new System.Drawing.Point(135, 234);
		this.Lab_EngFontShow.Name = "Lab_EngFontShow";
		this.Lab_EngFontShow.Size = new System.Drawing.Size(121, 35);
		this.Lab_EngFontShow.TabIndex = 18;
		this.Lab_EngFontShow.Text = "English";
		this.Lab_EngFontShow.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Lab_ChnFontShow.Location = new System.Drawing.Point(17, 234);
		this.Lab_ChnFontShow.Name = "Lab_ChnFontShow";
		this.Lab_ChnFontShow.Size = new System.Drawing.Size(119, 35);
		this.Lab_ChnFontShow.TabIndex = 17;
		this.Lab_ChnFontShow.Text = "中国汉字";
		this.Lab_ChnFontShow.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Chk_FontShading.AutoSize = true;
		this.Chk_FontShading.Location = new System.Drawing.Point(149, 206);
		this.Chk_FontShading.Name = "Chk_FontShading";
		this.Chk_FontShading.Size = new System.Drawing.Size(83, 22);
		this.Chk_FontShading.TabIndex = 14;
		this.Chk_FontShading.Text = "设置底纹";
		this.Chk_FontShading.UseVisualStyleBackColor = true;
		this.Chk_FontShading.CheckedChanged += new System.EventHandler(Chk_TextColor_CheckedChanged);
		this.Cmb_FontSize.FormattingEnabled = true;
		this.Cmb_FontSize.Location = new System.Drawing.Point(85, 169);
		this.Cmb_FontSize.Name = "Cmb_FontSize";
		this.Cmb_FontSize.Size = new System.Drawing.Size(74, 26);
		this.Cmb_FontSize.TabIndex = 12;
		this.Cmb_FontSize.SelectedIndexChanged += new System.EventHandler(Cmb_FontSize_SelectedIndexChanged);
		this.Cmb_FontSize.TextUpdate += new System.EventHandler(Cmb_FontSize_SelectedIndexChanged);
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(15, 173);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(64, 18);
		this.label5.TabIndex = 11;
		this.label5.Text = "字体大小";
		this.Txt_FormatDiscription.Location = new System.Drawing.Point(85, 57);
		this.Txt_FormatDiscription.Multiline = true;
		this.Txt_FormatDiscription.Name = "Txt_FormatDiscription";
		this.Txt_FormatDiscription.Size = new System.Drawing.Size(173, 42);
		this.Txt_FormatDiscription.TabIndex = 7;
		this.Txt_FormatDiscription.TextChanged += new System.EventHandler(Txt_FormatPainterName_TextChanged);
		this.Txt_FormatPainterName.Location = new System.Drawing.Point(85, 28);
		this.Txt_FormatPainterName.Name = "Txt_FormatPainterName";
		this.Txt_FormatPainterName.Size = new System.Drawing.Size(173, 25);
		this.Txt_FormatPainterName.TabIndex = 6;
		this.Txt_FormatPainterName.TextChanged += new System.EventHandler(Txt_FormatPainterName_TextChanged);
		this.Cmb_EngFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_EngFontName.FormattingEnabled = true;
		this.Cmb_EngFontName.Location = new System.Drawing.Point(85, 137);
		this.Cmb_EngFontName.Name = "Cmb_EngFontName";
		this.Cmb_EngFontName.Size = new System.Drawing.Size(173, 26);
		this.Cmb_EngFontName.TabIndex = 5;
		this.Cmb_EngFontName.SelectedIndexChanged += new System.EventHandler(Cmb_ChnFontName_SelectedIndexChanged);
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(15, 60);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(64, 18);
		this.label4.TabIndex = 4;
		this.label4.Text = "格式描述";
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(14, 31);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(64, 18);
		this.label3.TabIndex = 3;
		this.label3.Text = "格式名称";
		this.Cmb_ChnFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_ChnFontName.FormattingEnabled = true;
		this.Cmb_ChnFontName.Location = new System.Drawing.Point(85, 105);
		this.Cmb_ChnFontName.Name = "Cmb_ChnFontName";
		this.Cmb_ChnFontName.Size = new System.Drawing.Size(173, 26);
		this.Cmb_ChnFontName.TabIndex = 2;
		this.Cmb_ChnFontName.SelectedIndexChanged += new System.EventHandler(Cmb_ChnFontName_SelectedIndexChanged);
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(15, 141);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(64, 18);
		this.label2.TabIndex = 1;
		this.label2.Text = "西文字体";
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(15, 109);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(64, 18);
		this.label1.TabIndex = 0;
		this.label1.Text = "中文字体";
		this.Lst_FormatPainterList.FormattingEnabled = true;
		this.Lst_FormatPainterList.ItemHeight = 18;
		this.Lst_FormatPainterList.Location = new System.Drawing.Point(7, 11);
		this.Lst_FormatPainterList.Name = "Lst_FormatPainterList";
		this.Lst_FormatPainterList.Size = new System.Drawing.Size(161, 346);
		this.Lst_FormatPainterList.TabIndex = 1;
		this.Lst_FormatPainterList.SelectedIndexChanged += new System.EventHandler(Lst_FormatPainterList_SelectedIndexChanged);
		this.Btn_Apply.Location = new System.Drawing.Point(174, 318);
		this.Btn_Apply.Name = "Btn_Apply";
		this.Btn_Apply.Size = new System.Drawing.Size(265, 39);
		this.Btn_Apply.TabIndex = 2;
		this.Btn_Apply.Text = "应用格式刷";
		this.Btn_Apply.UseVisualStyleBackColor = true;
		this.Btn_Apply.Click += new System.EventHandler(Btn_Apply_Click);
		this.Btn_Delete.Location = new System.Drawing.Point(175, 272);
		this.Btn_Delete.Name = "Btn_Delete";
		this.Btn_Delete.Size = new System.Drawing.Size(80, 30);
		this.Btn_Delete.TabIndex = 22;
		this.Btn_Delete.Text = "删除格式";
		this.Btn_Delete.UseVisualStyleBackColor = true;
		this.Btn_Delete.Click += new System.EventHandler(Btn_Delete_Click);
		this.Btn_FontUnerline.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_FontUnerline.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, 0);
		this.Btn_FontUnerline.Location = new System.Drawing.Point(228, 167);
		this.Btn_FontUnerline.Name = "Btn_FontUnerline";
		this.Btn_FontUnerline.Pressed = false;
		this.Btn_FontUnerline.Size = new System.Drawing.Size(30, 30);
		this.Btn_FontUnerline.TabIndex = 10;
		this.Btn_FontUnerline.Text = "U";
		this.Btn_FontUnerline.UseVisualStyleBackColor = false;
		this.Btn_FontUnerline.Click += new System.EventHandler(Btn_FontBold_Click);
		this.Btn_FontItalic.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_FontItalic.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, 0);
		this.Btn_FontItalic.Location = new System.Drawing.Point(197, 167);
		this.Btn_FontItalic.Name = "Btn_FontItalic";
		this.Btn_FontItalic.Pressed = false;
		this.Btn_FontItalic.Size = new System.Drawing.Size(30, 30);
		this.Btn_FontItalic.TabIndex = 9;
		this.Btn_FontItalic.Text = "I";
		this.Btn_FontItalic.UseVisualStyleBackColor = false;
		this.Btn_FontItalic.Click += new System.EventHandler(Btn_FontBold_Click);
		this.Btn_FontBold.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_FontBold.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
		this.Btn_FontBold.Location = new System.Drawing.Point(165, 167);
		this.Btn_FontBold.Name = "Btn_FontBold";
		this.Btn_FontBold.Pressed = false;
		this.Btn_FontBold.Size = new System.Drawing.Size(30, 30);
		this.Btn_FontBold.TabIndex = 8;
		this.Btn_FontBold.Text = "B";
		this.Btn_FontBold.UseVisualStyleBackColor = false;
		this.Btn_FontBold.Click += new System.EventHandler(Btn_FontBold_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 18f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Btn_Apply);
		base.Controls.Add(this.Lst_FormatPainterList);
		base.Controls.Add(this.Grp_FromatSetting);
		this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		base.Margin = new System.Windows.Forms.Padding(4);
		base.Name = "FormatPainterUI";
		base.Size = new System.Drawing.Size(445, 363);
		this.Grp_FromatSetting.ResumeLayout(false);
		this.Grp_FromatSetting.PerformLayout();
		base.ResumeLayout(false);
	}
}
