using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class TableStyleSettings : UserControl
{
	private readonly List<float> FontSizePoint = new List<float>(16)
	{
		5f, 5.5f, 6.5f, 7.5f, 9f, 10.5f, 12f, 14f, 15f, 16f,
		18f, 22f, 24f, 26f, 36f, 42f
	};

	private readonly bool LoadingSettings;

	private IContainer components;

	private Label label1;

	private NumericUpDownWithUnit Nud_TableRows;

	private NumericUpDownWithUnit Nud_TableColumns;

	private Label label2;

	private ComboBox Cmb_TableFontName;

	private ComboBox Cmb_TableFontSize;

	private CheckBox Chk_TitleNumber;

	private ComboBox Cmb_TitleName;

	private Label label3;

	private ComboBox Cmb_NumberStyle;

	private Label Lab_NumberdLike;

	private CheckBox Chk_IncludeHeadings;

	private ComboBox Cmb_HeadingsLevel;

	private ComboBox Cmb_LinkChar;

	private Label label5;

	private ComboBox Cmb_FilledType;

	private Label label6;

	private Label label7;

	private Button Btn_BackgroundColor;

	private Button Btn_TableFontColor;

	private Label label8;

	private LineTypeSelectComboBox Lts_LineType;

	private Label label9;

	private ComboBox Cmb_LineWidth;

	private FlowLayoutPanel flowLayoutPanel1;

	private Label label10;

	private ComboBox Cmb_Line;

	private Label label4;

	private CheckBox Chk_FixRowHeight;

	private TextureSelectComboBox Cmb_TextureStyle;

	private WaterMarkTextBoxEx Txt_TableTitle;

	public TableStyleSettings()
	{
		InitializeComponent();
		LoadingSettings = true;
		Nud_TableRows.Value = ThisAddIn.tableSettings.Rows;
		Nud_TableColumns.Value = ThisAddIn.tableSettings.Columns;
		Chk_FixRowHeight.Checked = ThisAddIn.tableSettings.FixRowHeight;
		int num = FontSizePoint.IndexOf(ThisAddIn.tableSettings.FontSize);
		if (num != -1)
		{
			Cmb_TableFontSize.SelectedIndex = num;
		}
		else
		{
			Cmb_TableFontSize.Text = ThisAddIn.tableSettings.FontSize.ToString();
		}
		FontFamily[] families = new InstalledFontCollection().Families;
		foreach (FontFamily fontFamily in families)
		{
			Cmb_TableFontName.Items.Add(fontFamily.Name);
		}
		Cmb_TableFontName.Text = ThisAddIn.tableSettings.FontName;
		Btn_TableFontColor.BackColor = ThisAddIn.tableSettings.FontColor;
		Txt_TableTitle.Text = ThisAddIn.tableSettings.TableTitle;
		Chk_TitleNumber.Checked = ThisAddIn.tableSettings.CaptionLab;
		Cmb_TitleName.Text = ThisAddIn.tableSettings.CaptionTitle;
		Cmb_NumberStyle.SelectedIndex = ThisAddIn.tableSettings.CaptionNumberStyle;
		Chk_IncludeHeadings.Checked = ThisAddIn.tableSettings.CaptionIncludeHeadings;
		Cmb_HeadingsLevel.SelectedIndex = ((ThisAddIn.tableSettings.HeadingsLevel == -1) ? (-1) : (ThisAddIn.tableSettings.HeadingsLevel - 1));
		Cmb_LinkChar.SelectedIndex = ThisAddIn.tableSettings.LinkChar;
		Cmb_Line.SelectedIndex = 0;
		Lts_LineType.SelectedIndex = ThisAddIn.tableSettings.OuterLineType;
		UpdataLineWidthList(ThisAddIn.tableSettings.OuterLineType == 1);
		Cmb_LineWidth.SelectedIndex = ThisAddIn.tableSettings.OuterLineWidth;
		foreach (CaptionLabel captionLabel in Globals.ThisAddIn.Application.CaptionLabels)
		{
			Cmb_TitleName.Items.Add(captionLabel.Name);
		}
		Cmb_FilledType.SelectedIndex = ThisAddIn.tableSettings.FillType;
		Cmb_TextureStyle.SelectedIndex = ThisAddIn.tableSettings.TextureStyle;
		if (ThisAddIn.tableSettings.FillType == 0)
		{
			Cmb_TextureStyle.Enabled = false;
		}
		Btn_BackgroundColor.BackColor = ThisAddIn.tableSettings.BackgrounColor;
		if (ThisAddIn.tableSettings.FillType == 1)
		{
			Btn_BackgroundColor.Enabled = false;
		}
		if (Chk_TitleNumber.Checked)
		{
			Cmb_TitleName.Enabled = (Cmb_NumberStyle.Enabled = (Chk_IncludeHeadings.Enabled = true));
			if (Chk_IncludeHeadings.Checked)
			{
				Cmb_HeadingsLevel.Enabled = (Cmb_LinkChar.Enabled = true);
			}
		}
		else
		{
			Cmb_TitleName.Enabled = (Cmb_NumberStyle.Enabled = (Chk_IncludeHeadings.Enabled = false));
			Chk_IncludeHeadings.Checked = false;
			Cmb_HeadingsLevel.Enabled = (Cmb_LinkChar.Enabled = false);
		}
		ShowNumberedLike();
		LoadingSettings = false;
	}

	private void Chk_TitleNumber_CheckedChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			Cmb_TitleName.Enabled = Chk_TitleNumber.Checked;
			Cmb_NumberStyle.Enabled = Chk_TitleNumber.Checked;
			Chk_IncludeHeadings.Enabled = Chk_TitleNumber.Checked;
			if (Cmb_TitleName.Enabled && Cmb_TitleName.SelectedIndex == -1)
			{
				Cmb_TitleName.Text = "表";
			}
			if (!Chk_IncludeHeadings.Enabled)
			{
				Chk_IncludeHeadings.Checked = false;
			}
			ThisAddIn.tableSettings.CaptionLab = Chk_TitleNumber.Checked;
			ShowNumberedLike();
		}
	}

	private void Chk_IncludeHeadings_CheckedChanged(object sender, EventArgs e)
	{
		if (LoadingSettings)
		{
			return;
		}
		Cmb_HeadingsLevel.Enabled = Chk_IncludeHeadings.Checked;
		Cmb_LinkChar.Enabled = Chk_IncludeHeadings.Checked;
		if (Cmb_HeadingsLevel.Enabled)
		{
			if (Cmb_HeadingsLevel.SelectedIndex == -1)
			{
				Cmb_HeadingsLevel.SelectedIndex = 0;
			}
			if (Cmb_LinkChar.SelectedIndex == -1)
			{
				Cmb_LinkChar.SelectedIndex = 0;
			}
		}
		ThisAddIn.tableSettings.CaptionIncludeHeadings = Chk_IncludeHeadings.Checked;
		ShowNumberedLike();
	}

	private void Cmb_FilledType_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			Cmb_TextureStyle.Enabled = Cmb_FilledType.SelectedIndex != 0;
			Btn_BackgroundColor.Enabled = Cmb_FilledType.SelectedIndex != 1;
			ThisAddIn.tableSettings.FillType = Cmb_FilledType.SelectedIndex;
		}
	}

	private void Btn_TableFontColor_Click(object sender, EventArgs e)
	{
		ColorDialog colorDialog = new ColorDialog();
		if (colorDialog.ShowDialog() == DialogResult.OK)
		{
			(sender as Button).BackColor = colorDialog.Color;
			if ((sender as Button).Name == "Btn_TableFontColor")
			{
				ThisAddIn.tableSettings.FontColor = colorDialog.Color;
			}
			else
			{
				ThisAddIn.tableSettings.BackgrounColor = colorDialog.Color;
			}
		}
	}

	private void Cmb_TitleName_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			switch ((sender as ComboBox).Name)
			{
			case "Cmb_TitleName":
				ThisAddIn.tableSettings.CaptionTitle = Cmb_TitleName.Text;
				break;
			case "Cmb_NumberStyle":
				ThisAddIn.tableSettings.CaptionNumberStyle = Cmb_NumberStyle.SelectedIndex;
				break;
			case "Cmb_HeadingsLevel":
				ThisAddIn.tableSettings.HeadingsLevel = Cmb_HeadingsLevel.SelectedIndex + 1;
				break;
			case "Cmb_LinkChar":
				ThisAddIn.tableSettings.LinkChar = Cmb_LinkChar.SelectedIndex;
				break;
			}
			ShowNumberedLike();
		}
	}

	private void ShowNumberedLike()
	{
		string text = "";
		if (Chk_TitleNumber.Checked)
		{
			text = Cmb_TitleName.Text;
			if (Chk_IncludeHeadings.Checked)
			{
				text = text + " " + "1.1.1.1.1.1.1.1.1".Substring(0, 1 + 2 * Cmb_HeadingsLevel.SelectedIndex);
				switch (Cmb_LinkChar.SelectedIndex)
				{
				case 0:
					text += " -";
					break;
				case 1:
					text += " .";
					break;
				case 2:
					text += " :";
					break;
				case 3:
					text += " —";
					break;
				}
			}
			switch (Cmb_NumberStyle.SelectedIndex)
			{
			case 0:
				text += " 1";
				break;
			case 1:
				text += " １";
				break;
			case 2:
				text += " A";
				break;
			case 3:
				text += " a";
				break;
			case 4:
				text += " I";
				break;
			case 5:
				text += " i";
				break;
			case 6:
				text += " 一";
				break;
			case 7:
				text += " 壹";
				break;
			case 8:
				text += " 甲";
				break;
			}
		}
		Lab_NumberdLike.Text = ((text == "") ? "编号示例" : text);
	}

	private void Txt_TableTitle_TextChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			ThisAddIn.tableSettings.TableTitle = Txt_TableTitle.Text;
		}
	}

	private void Cmb_TableFontName_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			ThisAddIn.tableSettings.FontName = Cmb_TableFontName.Text;
		}
	}

	private void Cmb_TableFontSize_Leave(object sender, EventArgs e)
	{
		if (Cmb_TableFontSize.SelectedIndex != -1)
		{
			return;
		}
		int num = FontSizePoint.IndexOf(Convert.ToSingle((Cmb_TableFontSize.Text == "") ? "0" : Cmb_TableFontSize.Text));
		if (num != -1)
		{
			Cmb_TableFontSize.SelectedIndex = num;
			ThisAddIn.tableSettings.FontSize = FontSizePoint.IndexOf(num);
		}
		else if (Cmb_TableFontSize.Text != "" && Regex.IsMatch(Cmb_TableFontSize.Text, "^[1-9]{1,4}(\\.5|\\.0){0,1}$"))
		{
			float num2 = Convert.ToSingle(Cmb_TableFontSize.Text);
			if (num2 >= 1f || num2 <= 1638f)
			{
				ThisAddIn.tableSettings.FontSize = Convert.ToSingle(Cmb_TableFontSize.Text);
			}
		}
	}

	private void Nud_TableRows_ValueChanged(object sender, EventArgs e)
	{
		if (LoadingSettings)
		{
			return;
		}
		string name = (sender as NumericUpDownWithUnit).Name;
		if (!(name == "Nud_TableRows"))
		{
			if (name == "Nud_TableColumns")
			{
				ThisAddIn.tableSettings.Columns = (int)Nud_TableColumns.Value;
			}
		}
		else
		{
			ThisAddIn.tableSettings.Rows = (int)Nud_TableRows.Value;
		}
	}

	private void Cmb_TitleName_Leave(object sender, EventArgs e)
	{
		if (!Cmb_TitleName.Items.Contains(Cmb_TitleName.Text))
		{
			CaptionLabel captionLabel = Globals.ThisAddIn.Application.CaptionLabels.Add(Cmb_TitleName.Text);
			Cmb_TitleName.SelectedIndex = Cmb_TitleName.Items.Add(captionLabel.Name);
			ThisAddIn.tableSettings.CaptionTitle = Cmb_TitleName.Text;
		}
	}

	private void Lts_LineType_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			switch (Cmb_Line.SelectedIndex)
			{
			case 0:
				ThisAddIn.tableSettings.OuterLineType = Lts_LineType.SelectedIndex;
				break;
			case 1:
				ThisAddIn.tableSettings.InnerLineType = Lts_LineType.SelectedIndex;
				break;
			case 2:
				ThisAddIn.tableSettings.TitleRowLineType = Lts_LineType.SelectedIndex;
				break;
			}
			UpdataLineWidthList(Lts_LineType.SelectedIndex == 1);
			switch (Lts_LineType.SelectedIndex)
			{
			case 0:
			case 1:
				Cmb_LineWidth.SelectedIndex = 3;
				break;
			case 2:
			case 3:
				Cmb_LineWidth.SelectedIndex = 6;
				break;
			}
		}
	}

	private void Cmb_LineWidth_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			switch (Cmb_Line.SelectedIndex)
			{
			case 0:
				ThisAddIn.tableSettings.OuterLineWidth = Cmb_LineWidth.SelectedIndex;
				break;
			case 1:
				ThisAddIn.tableSettings.InnerLineWidth = Cmb_LineWidth.SelectedIndex;
				break;
			case 2:
				ThisAddIn.tableSettings.TitleRowLineWidth = Cmb_LineWidth.SelectedIndex;
				break;
			}
		}
	}

	private void Cmb_Line_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			switch (Cmb_Line.SelectedIndex)
			{
			case 0:
				Lts_LineType.SelectedIndex = ThisAddIn.tableSettings.OuterLineType;
				UpdataLineWidthList(ThisAddIn.tableSettings.OuterLineType == 1);
				Cmb_LineWidth.SelectedIndex = ThisAddIn.tableSettings.OuterLineWidth;
				break;
			case 1:
				Lts_LineType.SelectedIndex = ThisAddIn.tableSettings.InnerLineType;
				UpdataLineWidthList(ThisAddIn.tableSettings.InnerLineType == 1);
				Cmb_LineWidth.SelectedIndex = ThisAddIn.tableSettings.InnerLineWidth;
				break;
			case 2:
				Lts_LineType.SelectedIndex = ThisAddIn.tableSettings.TitleRowLineType;
				UpdataLineWidthList(ThisAddIn.tableSettings.TitleRowLineType == 1);
				Cmb_LineWidth.SelectedIndex = ThisAddIn.tableSettings.TitleRowLineWidth;
				break;
			}
		}
	}

	private void UpdataLineWidthList(bool IsDoubleLine)
	{
		Cmb_LineWidth.Items.Clear();
		if (IsDoubleLine)
		{
			Cmb_LineWidth.Items.AddRange(new object[6] { "0.25", "0.5", "0.75", "1.5", "2.25", "3.0" });
		}
		else
		{
			Cmb_LineWidth.Items.AddRange(new object[9] { "0.25", "0.5", "0.75", "1.0", "1.5", "2.25", "3.0", "4.5", "6.0" });
		}
	}

	private void Chk_FixRowHeight_CheckedChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			ThisAddIn.tableSettings.FixRowHeight = Chk_FixRowHeight.Checked;
		}
	}

	private void Cmb_TableFontSize_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			ThisAddIn.tableSettings.FontSize = FontSizePoint[Cmb_TableFontSize.SelectedIndex];
		}
	}

	private void Cmb_TextureStyle_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!LoadingSettings)
		{
			ThisAddIn.tableSettings.TextureStyle = Cmb_TextureStyle.SelectedIndex;
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
		this.label1 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.Cmb_TableFontName = new System.Windows.Forms.ComboBox();
		this.Cmb_TableFontSize = new System.Windows.Forms.ComboBox();
		this.Chk_TitleNumber = new System.Windows.Forms.CheckBox();
		this.Cmb_TitleName = new System.Windows.Forms.ComboBox();
		this.label3 = new System.Windows.Forms.Label();
		this.Cmb_NumberStyle = new System.Windows.Forms.ComboBox();
		this.Lab_NumberdLike = new System.Windows.Forms.Label();
		this.Chk_IncludeHeadings = new System.Windows.Forms.CheckBox();
		this.Cmb_HeadingsLevel = new System.Windows.Forms.ComboBox();
		this.Cmb_LinkChar = new System.Windows.Forms.ComboBox();
		this.label5 = new System.Windows.Forms.Label();
		this.Cmb_FilledType = new System.Windows.Forms.ComboBox();
		this.label6 = new System.Windows.Forms.Label();
		this.label7 = new System.Windows.Forms.Label();
		this.Btn_BackgroundColor = new System.Windows.Forms.Button();
		this.Btn_TableFontColor = new System.Windows.Forms.Button();
		this.label8 = new System.Windows.Forms.Label();
		this.label9 = new System.Windows.Forms.Label();
		this.Cmb_LineWidth = new System.Windows.Forms.ComboBox();
		this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
		this.label4 = new System.Windows.Forms.Label();
		this.label10 = new System.Windows.Forms.Label();
		this.Cmb_Line = new System.Windows.Forms.ComboBox();
		this.Chk_FixRowHeight = new System.Windows.Forms.CheckBox();
		this.Cmb_TextureStyle = new WordFormatHelper.TextureSelectComboBox();
		this.Lts_LineType = new WordFormatHelper.LineTypeSelectComboBox();
		this.Nud_TableColumns = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_TableRows = new WordFormatHelper.NumericUpDownWithUnit();
		this.Txt_TableTitle = new WordFormatHelper.WaterMarkTextBoxEx();
		this.flowLayoutPanel1.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_TableColumns).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_TableRows).BeginInit();
		base.SuspendLayout();
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(11, 11);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(65, 20);
		this.label1.TabIndex = 0;
		this.label1.Text = "表格大小";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(11, 44);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(65, 20);
		this.label2.TabIndex = 3;
		this.label2.Text = "表格字体";
		this.Cmb_TableFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TableFontName.FormattingEnabled = true;
		this.Cmb_TableFontName.Location = new System.Drawing.Point(82, 40);
		this.Cmb_TableFontName.Name = "Cmb_TableFontName";
		this.Cmb_TableFontName.Size = new System.Drawing.Size(164, 28);
		this.Cmb_TableFontName.TabIndex = 4;
		this.Cmb_TableFontName.SelectedIndexChanged += new System.EventHandler(Cmb_TableFontName_SelectedIndexChanged);
		this.Cmb_TableFontSize.FormattingEnabled = true;
		this.Cmb_TableFontSize.Items.AddRange(new object[16]
		{
			"八号", "七号", "小六", "六号", "小五", "五号", "小四", "四号", "小三", "三号",
			"小二", "二号", "小一", "一号", "小初", "初号"
		});
		this.Cmb_TableFontSize.Location = new System.Drawing.Point(256, 40);
		this.Cmb_TableFontSize.Name = "Cmb_TableFontSize";
		this.Cmb_TableFontSize.Size = new System.Drawing.Size(73, 28);
		this.Cmb_TableFontSize.TabIndex = 5;
		this.Cmb_TableFontSize.SelectedIndexChanged += new System.EventHandler(Cmb_TableFontSize_SelectedIndexChanged);
		this.Cmb_TableFontSize.Leave += new System.EventHandler(Cmb_TableFontSize_Leave);
		this.Chk_TitleNumber.Location = new System.Drawing.Point(3, 35);
		this.Chk_TitleNumber.Name = "Chk_TitleNumber";
		this.Chk_TitleNumber.Size = new System.Drawing.Size(112, 28);
		this.Chk_TitleNumber.TabIndex = 6;
		this.Chk_TitleNumber.Text = "使用题注标签";
		this.Chk_TitleNumber.UseVisualStyleBackColor = true;
		this.Chk_TitleNumber.CheckedChanged += new System.EventHandler(Chk_TitleNumber_CheckedChanged);
		this.Cmb_TitleName.Enabled = false;
		this.Cmb_TitleName.FormattingEnabled = true;
		this.Cmb_TitleName.Location = new System.Drawing.Point(121, 35);
		this.Cmb_TitleName.Name = "Cmb_TitleName";
		this.Cmb_TitleName.Size = new System.Drawing.Size(100, 28);
		this.Cmb_TitleName.TabIndex = 7;
		this.Cmb_TitleName.SelectedIndexChanged += new System.EventHandler(Cmb_TitleName_SelectedIndexChanged);
		this.Cmb_TitleName.Leave += new System.EventHandler(Cmb_TitleName_Leave);
		this.label3.Location = new System.Drawing.Point(227, 32);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(70, 28);
		this.label3.TabIndex = 8;
		this.label3.Text = "编号样式";
		this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_NumberStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_NumberStyle.Enabled = false;
		this.Cmb_NumberStyle.FormattingEnabled = true;
		this.Cmb_NumberStyle.Items.AddRange(new object[9] { "1,2,3...", "全角数字", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,貳,叁...", "甲,乙,丙..." });
		this.Cmb_NumberStyle.Location = new System.Drawing.Point(303, 35);
		this.Cmb_NumberStyle.Name = "Cmb_NumberStyle";
		this.Cmb_NumberStyle.Size = new System.Drawing.Size(100, 28);
		this.Cmb_NumberStyle.TabIndex = 9;
		this.Cmb_NumberStyle.SelectedIndexChanged += new System.EventHandler(Cmb_TitleName_SelectedIndexChanged);
		this.Lab_NumberdLike.Location = new System.Drawing.Point(3, 0);
		this.Lab_NumberdLike.Name = "Lab_NumberdLike";
		this.Lab_NumberdLike.Size = new System.Drawing.Size(112, 26);
		this.Lab_NumberdLike.TabIndex = 10;
		this.Lab_NumberdLike.Text = "编号示例";
		this.Lab_NumberdLike.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
		this.Chk_IncludeHeadings.Enabled = false;
		this.Chk_IncludeHeadings.Location = new System.Drawing.Point(3, 69);
		this.Chk_IncludeHeadings.Name = "Chk_IncludeHeadings";
		this.Chk_IncludeHeadings.Size = new System.Drawing.Size(112, 28);
		this.Chk_IncludeHeadings.TabIndex = 11;
		this.Chk_IncludeHeadings.Text = "包含章节编号";
		this.Chk_IncludeHeadings.UseVisualStyleBackColor = true;
		this.Chk_IncludeHeadings.CheckedChanged += new System.EventHandler(Chk_IncludeHeadings_CheckedChanged);
		this.Cmb_HeadingsLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_HeadingsLevel.Enabled = false;
		this.Cmb_HeadingsLevel.FormattingEnabled = true;
		this.Cmb_HeadingsLevel.Items.AddRange(new object[9] { "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9" });
		this.Cmb_HeadingsLevel.Location = new System.Drawing.Point(121, 69);
		this.Cmb_HeadingsLevel.Name = "Cmb_HeadingsLevel";
		this.Cmb_HeadingsLevel.Size = new System.Drawing.Size(100, 28);
		this.Cmb_HeadingsLevel.TabIndex = 12;
		this.Cmb_HeadingsLevel.SelectedIndexChanged += new System.EventHandler(Cmb_TitleName_SelectedIndexChanged);
		this.Cmb_LinkChar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_LinkChar.Enabled = false;
		this.Cmb_LinkChar.FormattingEnabled = true;
		this.Cmb_LinkChar.Items.AddRange(new object[4] { "- （连字符）", ". （句点）", ": （冒号）", "—— （长划线）" });
		this.Cmb_LinkChar.Location = new System.Drawing.Point(303, 69);
		this.Cmb_LinkChar.Name = "Cmb_LinkChar";
		this.Cmb_LinkChar.Size = new System.Drawing.Size(100, 28);
		this.Cmb_LinkChar.TabIndex = 14;
		this.Cmb_LinkChar.SelectedIndexChanged += new System.EventHandler(Cmb_TitleName_SelectedIndexChanged);
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(11, 244);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(65, 20);
		this.label5.TabIndex = 15;
		this.label5.Text = "填充类型";
		this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_FilledType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_FilledType.FormattingEnabled = true;
		this.Cmb_FilledType.Items.AddRange(new object[3] { "纯色", "纹理", "纯色+纹理" });
		this.Cmb_FilledType.Location = new System.Drawing.Point(82, 240);
		this.Cmb_FilledType.Name = "Cmb_FilledType";
		this.Cmb_FilledType.Size = new System.Drawing.Size(109, 28);
		this.Cmb_FilledType.TabIndex = 16;
		this.Cmb_FilledType.SelectedIndexChanged += new System.EventHandler(Cmb_FilledType_SelectedIndexChanged);
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(197, 244);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(65, 20);
		this.label6.TabIndex = 17;
		this.label6.Text = "纹理样式";
		this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label7.AutoSize = true;
		this.label7.Location = new System.Drawing.Point(335, 244);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(65, 20);
		this.label7.TabIndex = 19;
		this.label7.Text = "填充底色";
		this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.Btn_BackgroundColor.BackColor = System.Drawing.Color.White;
		this.Btn_BackgroundColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
		this.Btn_BackgroundColor.Location = new System.Drawing.Point(398, 244);
		this.Btn_BackgroundColor.Name = "Btn_BackgroundColor";
		this.Btn_BackgroundColor.Size = new System.Drawing.Size(20, 20);
		this.Btn_BackgroundColor.TabIndex = 20;
		this.Btn_BackgroundColor.UseVisualStyleBackColor = false;
		this.Btn_BackgroundColor.Click += new System.EventHandler(Btn_TableFontColor_Click);
		this.Btn_TableFontColor.BackColor = System.Drawing.Color.Black;
		this.Btn_TableFontColor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
		this.Btn_TableFontColor.Location = new System.Drawing.Point(398, 44);
		this.Btn_TableFontColor.Name = "Btn_TableFontColor";
		this.Btn_TableFontColor.Size = new System.Drawing.Size(20, 20);
		this.Btn_TableFontColor.TabIndex = 22;
		this.Btn_TableFontColor.UseVisualStyleBackColor = false;
		this.Btn_TableFontColor.Click += new System.EventHandler(Btn_TableFontColor_Click);
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(333, 44);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(65, 20);
		this.label8.TabIndex = 21;
		this.label8.Text = "字体颜色";
		this.label9.Location = new System.Drawing.Point(310, 197);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(40, 28);
		this.label9.TabIndex = 26;
		this.label9.Text = "线宽";
		this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.Cmb_LineWidth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_LineWidth.FormattingEnabled = true;
		this.Cmb_LineWidth.Location = new System.Drawing.Point(353, 197);
		this.Cmb_LineWidth.Name = "Cmb_LineWidth";
		this.Cmb_LineWidth.Size = new System.Drawing.Size(65, 28);
		this.Cmb_LineWidth.TabIndex = 27;
		this.Cmb_LineWidth.SelectedIndexChanged += new System.EventHandler(Cmb_LineWidth_SelectedIndexChanged);
		this.flowLayoutPanel1.Controls.Add(this.Lab_NumberdLike);
		this.flowLayoutPanel1.Controls.Add(this.Txt_TableTitle);
		this.flowLayoutPanel1.Controls.Add(this.Chk_TitleNumber);
		this.flowLayoutPanel1.Controls.Add(this.Cmb_TitleName);
		this.flowLayoutPanel1.Controls.Add(this.label3);
		this.flowLayoutPanel1.Controls.Add(this.Cmb_NumberStyle);
		this.flowLayoutPanel1.Controls.Add(this.Chk_IncludeHeadings);
		this.flowLayoutPanel1.Controls.Add(this.Cmb_HeadingsLevel);
		this.flowLayoutPanel1.Controls.Add(this.label4);
		this.flowLayoutPanel1.Controls.Add(this.Cmb_LinkChar);
		this.flowLayoutPanel1.Location = new System.Drawing.Point(12, 80);
		this.flowLayoutPanel1.Name = "flowLayoutPanel1";
		this.flowLayoutPanel1.Size = new System.Drawing.Size(406, 104);
		this.flowLayoutPanel1.TabIndex = 36;
		this.label4.Location = new System.Drawing.Point(227, 66);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(70, 28);
		this.label4.TabIndex = 24;
		this.label4.Text = "连接符号";
		this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label10.Location = new System.Drawing.Point(11, 197);
		this.label10.Name = "label10";
		this.label10.Size = new System.Drawing.Size(40, 28);
		this.label10.TabIndex = 28;
		this.label10.Text = "线型";
		this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_Line.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_Line.FormattingEnabled = true;
		this.Cmb_Line.Items.AddRange(new object[3] { "粗线", "细线", "标题栏底线" });
		this.Cmb_Line.Location = new System.Drawing.Point(51, 197);
		this.Cmb_Line.Name = "Cmb_Line";
		this.Cmb_Line.Size = new System.Drawing.Size(110, 28);
		this.Cmb_Line.TabIndex = 29;
		this.Cmb_Line.SelectedIndexChanged += new System.EventHandler(Cmb_Line_SelectedIndexChanged);
		this.Chk_FixRowHeight.AutoSize = true;
		this.Chk_FixRowHeight.Location = new System.Drawing.Point(271, 10);
		this.Chk_FixRowHeight.Name = "Chk_FixRowHeight";
		this.Chk_FixRowHeight.Size = new System.Drawing.Size(154, 24);
		this.Chk_FixRowHeight.TabIndex = 37;
		this.Chk_FixRowHeight.Text = "按字体大小固定行高";
		this.Chk_FixRowHeight.UseVisualStyleBackColor = true;
		this.Chk_FixRowHeight.CheckedChanged += new System.EventHandler(Chk_FixRowHeight_CheckedChanged);
		this.Cmb_TextureStyle.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		this.Cmb_TextureStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TextureStyle.FormattingEnabled = true;
		this.Cmb_TextureStyle.Items.AddRange(new object[12]
		{
			"", "", "", "", "", "", "", "", "", "",
			"", ""
		});
		this.Cmb_TextureStyle.Location = new System.Drawing.Point(268, 241);
		this.Cmb_TextureStyle.Name = "Cmb_TextureStyle";
		this.Cmb_TextureStyle.Size = new System.Drawing.Size(61, 27);
		this.Cmb_TextureStyle.TabIndex = 38;
		this.Cmb_TextureStyle.SelectedIndexChanged += new System.EventHandler(Cmb_TextureStyle_SelectedIndexChanged);
		this.Lts_LineType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		this.Lts_LineType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Lts_LineType.FormattingEnabled = true;
		this.Lts_LineType.Items.AddRange(new object[4] { "实线", "双实线", "细粗实线", "粗细实线" });
		this.Lts_LineType.Location = new System.Drawing.Point(167, 198);
		this.Lts_LineType.Name = "Lts_LineType";
		this.Lts_LineType.Size = new System.Drawing.Size(141, 27);
		this.Lts_LineType.TabIndex = 24;
		this.Lts_LineType.SelectedIndexChanged += new System.EventHandler(Lts_LineType_SelectedIndexChanged);
		this.Nud_TableColumns.Label = "列";
		this.Nud_TableColumns.Location = new System.Drawing.Point(167, 8);
		this.Nud_TableColumns.Maximum = new decimal(new int[4] { 63, 0, 0, 0 });
		this.Nud_TableColumns.Minimum = new decimal(new int[4] { 3, 0, 0, 0 });
		this.Nud_TableColumns.Name = "Nud_TableColumns";
		this.Nud_TableColumns.Size = new System.Drawing.Size(79, 26);
		this.Nud_TableColumns.TabIndex = 2;
		this.Nud_TableColumns.Value = new decimal(new int[4] { 3, 0, 0, 0 });
		this.Nud_TableColumns.ValueChanged += new System.EventHandler(Nud_TableRows_ValueChanged);
		this.Nud_TableRows.Label = "行";
		this.Nud_TableRows.Location = new System.Drawing.Point(82, 8);
		this.Nud_TableRows.Maximum = new decimal(new int[4] { 32767, 0, 0, 0 });
		this.Nud_TableRows.Minimum = new decimal(new int[4] { 3, 0, 0, 0 });
		this.Nud_TableRows.Name = "Nud_TableRows";
		this.Nud_TableRows.Size = new System.Drawing.Size(79, 26);
		this.Nud_TableRows.TabIndex = 1;
		this.Nud_TableRows.Value = new decimal(new int[4] { 3, 0, 0, 0 });
		this.Nud_TableRows.ValueChanged += new System.EventHandler(Nud_TableRows_ValueChanged);
		this.Txt_TableTitle.Location = new System.Drawing.Point(121, 3);
		this.Txt_TableTitle.Name = "Txt_TableTitle";
		this.Txt_TableTitle.Size = new System.Drawing.Size(282, 26);
		this.Txt_TableTitle.TabIndex = 39;
		this.Txt_TableTitle.WaterMark = "请输入表格标题";
		this.Txt_TableTitle.WaterTextColor = System.Drawing.Color.SteelBlue;
		this.Txt_TableTitle.TextChanged += new System.EventHandler(Txt_TableTitle_TextChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Cmb_TextureStyle);
		base.Controls.Add(this.Chk_FixRowHeight);
		base.Controls.Add(this.Cmb_LineWidth);
		base.Controls.Add(this.label9);
		base.Controls.Add(this.Lts_LineType);
		base.Controls.Add(this.Cmb_Line);
		base.Controls.Add(this.label10);
		base.Controls.Add(this.Cmb_FilledType);
		base.Controls.Add(this.label5);
		base.Controls.Add(this.Btn_BackgroundColor);
		base.Controls.Add(this.label7);
		base.Controls.Add(this.label6);
		base.Controls.Add(this.flowLayoutPanel1);
		base.Controls.Add(this.Btn_TableFontColor);
		base.Controls.Add(this.label8);
		base.Controls.Add(this.Cmb_TableFontSize);
		base.Controls.Add(this.Cmb_TableFontName);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.Nud_TableColumns);
		base.Controls.Add(this.Nud_TableRows);
		base.Controls.Add(this.label1);
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		base.Name = "TableStyleSettings";
		base.Size = new System.Drawing.Size(428, 278);
		this.flowLayoutPanel1.ResumeLayout(false);
		this.flowLayoutPanel1.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_TableColumns).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_TableRows).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}