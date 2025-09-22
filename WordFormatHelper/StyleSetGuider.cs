using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class StyleSetGuider : UserControl
{
	private readonly BindingList<string> FontSizeCha = new BindingList<string>
	{
		"八号", "七号", "小六", "六号", "小五", "五号", "小四", "四号", "小三", "三号",
		"小二", "二号", "小一", "一号", "小初", "初号"
	};

	private readonly List<float> FontSizePoint = new List<float>(16)
	{
		5f, 5.5f, 6.5f, 7.5f, 9f, 10.5f, 12f, 14f, 15f, 16f,
		18f, 22f, 24f, 26f, 36f, 42f
	};

	private readonly List<WdPaperSize> PaperSize = new List<WdPaperSize>(4)
	{
		WdPaperSize.wdPaperA3,
		WdPaperSize.wdPaperA4,
		WdPaperSize.wdPaperA5,
		WdPaperSize.wdPaperB5
	};

	private readonly BindingList<string> StyleNames;

	private static readonly List<CustomStyle> Styles = new List<CustomStyle>(17)
	{
		new CustomStyle("正文", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 1", null, 14f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 2", null, 12f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 3", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 4", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 5", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 6", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 7", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 8", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题 9", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("标题", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("副标题", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("附录标题", null, 14f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("表格标题", null, 0f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("插图标题", null, 0f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("列表段落", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("表内文字", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false)
	};

	private static bool StyleSelectChanged;

	private static int currentStyle;

	private float numIndent;

	private float textIndent;

	private float afterIndent;

	private IContainer components;

	private GroupBox groupBox1;

	private RadioButton Rdo_NewDocument;

	private RadioButton Rdo_UseCurrentDocument;

	private Label label1;

	private ComboBox Cmb_PaperSize;

	private GroupBox Grp_PageSetup;

	private ComboBox Cmb_PaperDirection;

	private Label label2;

	private GroupBox groupBox3;

	private CheckBox Chk_UnderLine;

	private CheckBox Chk_Italic;

	private CheckBox Chk_Bold;

	private ComboBox Cmb_FontSize;

	private Label label4;

	private ListBox Lst_Styles;

	private ComboBox Cmb_PageMargin;

	private Label label5;

	private Label label6;

	private Label label8;

	private NumericUpDownWithUnit Nud_FirstLineIndentByChar;

	private Label label7;

	private NumericUpDownWithUnit Nud_FirstLineIndent;

	private NumericUpDownWithUnit Nud_LeftIndent;

	private Label label11;

	private NumericUpDownWithUnit Nud_AfterSpacing;

	private Label label10;

	private NumericUpDownWithUnit Nud_BefreSpacing;

	private Label label9;

	private NumericUpDownWithUnit Nud_LineSpacing;

	private ComboBox Cmb_SetLevel;

	private Label label12;

	private ComboBox Cmb_GutterPosition;

	private CheckBox Chk_SetGutter;

	private NumericUpDownWithUnit Nud_GutterValue;

	private CheckBox Chk_CreateListLevels;

	private Button Btn_AddStyle;

	private FlowLayoutPanel Pal_Font;

	private Label label3;

	private ComboBox Cmb_FontName;

	private FlowLayoutPanel flowLayoutPanel2;

	private FlowLayoutPanel Pal_ParaIndent;

	private FlowLayoutPanel Pal_ParaSpacing;

	private Button Btn_DelStyle;

	private Label Lab_StyleInfo;

	private Button Btn_ApplySet;

	private FlowLayoutPanel Pal_NumberList;

	private Label label13;

	private ComboBox Cmb_NumberStyle;

	private Label label14;

	private TextBox Txt_NumberFormat;

	private Label label15;

	private ComboBox Cmb_ParaAligment;

	private CheckBox Chk_BeforeBreak;

	private Button Btn_ReadDocumentStyle;

	private ComboBox Cmb_PreSettings;

	private Label label16;

	private WaterMarkTextBoxEx Txt_AddStyleName;

	private CheckBox Chk_SetPage;

	protected override CreateParams CreateParams
	{
		get
		{
			CreateParams obj = base.CreateParams;
			obj.ExStyle |= 33554432;
			return obj;
		}
	}

	public StyleSetGuider()
	{
		InitializeComponent();
		Cmb_PaperSize.SelectedIndex = 1;
		Cmb_PaperDirection.SelectedIndex = 0;
		Cmb_PageMargin.SelectedIndex = 0;
		Cmb_GutterPosition.SelectedIndex = 0;
		FontFamily[] families = new InstalledFontCollection().Families;
		foreach (FontFamily fontFamily in families)
		{
			Cmb_FontName.Items.Add(fontFamily.Name);
		}
		StyleNames = new BindingList<string>();
		foreach (CustomStyle style in Styles)
		{
			if (!Regex.IsMatch(style.Name, "标题 [4-9]"))
			{
				StyleNames.Add(style.Name);
			}
		}
		Lst_Styles.DataSource = StyleNames;
		Cmb_FontSize.DataSource = FontSizeCha;
		Cmb_FontSize.SelectedIndex = -1;
		Cmb_SetLevel.SelectedIndex = 3;
		Lst_Styles.SelectedIndexChanged += Lst_Styles_SelectedIndexChanged;
		Lst_Styles.SelectedIndex = -1;
		Cmb_NumberStyle.SelectedIndex = -1;
		Cmb_ParaAligment.SelectedIndex = -1;
		Cmb_FontName.SelectedIndexChanged += StyleFontChanged;
		Cmb_FontSize.SelectedIndexChanged += StyleFontChanged;
		Cmb_ParaAligment.SelectedIndexChanged += StyleFontChanged;
		Nud_LeftIndent.ValueChanged += IndentSpacingChanged;
		Nud_FirstLineIndent.ValueChanged += IndentSpacingChanged;
		Nud_FirstLineIndentByChar.ValueChanged += IndentSpacingChanged;
		Nud_LineSpacing.ValueChanged += IndentSpacingChanged;
		Nud_BefreSpacing.ValueChanged += IndentSpacingChanged;
		Nud_AfterSpacing.ValueChanged += IndentSpacingChanged;
		Chk_Bold.CheckedChanged += FontStyleChanged;
		Chk_Italic.CheckedChanged += FontStyleChanged;
		Chk_UnderLine.CheckedChanged += FontStyleChanged;
		Chk_BeforeBreak.CheckedChanged += FontStyleChanged;
	}

	private void Chk_SetGutter_CheckedChanged(object sender, EventArgs e)
	{
		Cmb_GutterPosition.Enabled = Chk_SetGutter.Checked;
		Nud_GutterValue.Enabled = Chk_SetGutter.Checked;
		if (Chk_SetGutter.Checked && Cmb_GutterPosition.SelectedIndex == -1)
		{
			Cmb_GutterPosition.SelectedIndex = 0;
		}
	}

	private void Cmb_SetLevel_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (Cmb_SetLevel.SelectedIndex == 0)
		{
			Chk_CreateListLevels.Enabled = false;
			Chk_CreateListLevels.Checked = false;
		}
		else
		{
			Chk_CreateListLevels.Enabled = true;
		}
		for (int i = 1; i < 10; i++)
		{
			if (i <= Cmb_SetLevel.SelectedIndex)
			{
				if (StyleNames.IndexOf("标题 " + i) == -1)
				{
					StyleNames.Insert(i, "标题 " + i);
				}
			}
			else if (StyleNames.IndexOf("标题 " + i) != -1)
			{
				StyleNames.Remove("标题 " + i);
			}
		}
		Lst_Styles.SelectedIndex = -1;
	}

	private void Btn_AddStyle_Click(object sender, EventArgs e)
	{
		StyleNames.Add(Txt_AddStyleName.Text);
		Lst_Styles.SelectedIndex = Lst_Styles.Items.Count - 1;
		CustomStyle item = new CustomStyle(Txt_AddStyleName.Text, null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: true);
		Styles.Add(item);
		Txt_AddStyleName.Text = "";
		Btn_AddStyle.Enabled = false;
		Btn_DelStyle.Enabled = true;
	}

	private void Lst_Styles_SelectedIndexChanged(object sender, EventArgs e)
	{
		StyleSelectChanged = true;
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		if (Lst_Styles.SelectedIndex == -1)
		{
			Pal_Font.Enabled = false;
			Pal_ParaIndent.Enabled = false;
			Pal_ParaSpacing.Enabled = false;
			Pal_NumberList.Enabled = false;
			return;
		}
		Pal_Font.Enabled = true;
		Pal_ParaIndent.Enabled = true;
		Pal_ParaSpacing.Enabled = true;
		foreach (object selectedItem in Lst_Styles.SelectedItems)
		{
			if (Regex.IsMatch(selectedItem.ToString(), "标题 [1-9]|列表段落"))
			{
				Pal_NumberList.Enabled = true;
				break;
			}
			Pal_NumberList.Enabled = false;
		}
		foreach (CustomStyle style in Styles)
		{
			if (style.Name == (string)Lst_Styles.SelectedItem)
			{
				currentStyle = Styles.IndexOf(style);
				break;
			}
		}
		Cmb_FontName.SelectedIndex = Cmb_FontName.Items.IndexOf(Styles[currentStyle].FontName);
		if (FontSizePoint.IndexOf(Styles[currentStyle].FontSize) == -1)
		{
			Cmb_FontSize.Text = Styles[currentStyle].FontSize.ToString();
		}
		else
		{
			Cmb_FontSize.Text = null;
			Cmb_FontSize.SelectedIndex = FontSizePoint.IndexOf(Styles[currentStyle].FontSize);
		}
		Chk_Bold.Checked = Styles[currentStyle].IsBold;
		Chk_Italic.Checked = Styles[currentStyle].IsItalic;
		Chk_UnderLine.Checked = Styles[currentStyle].IsUnderline;
		Cmb_ParaAligment.SelectedIndex = Styles[currentStyle].ParagraphAlignment;
		Nud_LeftIndent.Value = (decimal)application.PointsToCentimeters(Styles[currentStyle].LeftIndent);
		Nud_FirstLineIndent.Value = (decimal)application.PointsToCentimeters(Styles[currentStyle].FirstLineIndent);
		Nud_FirstLineIndentByChar.Value = Styles[currentStyle].FirstLineIndentByChar;
		Nud_LineSpacing.Value = (decimal)Styles[currentStyle].LineSpacing;
		Chk_BeforeBreak.Checked = Styles[currentStyle].BeforeBreak;
		Nud_BefreSpacing.Value = (decimal)Styles[currentStyle].BeforeSpacing;
		Nud_AfterSpacing.Value = (decimal)Styles[currentStyle].AfterSpacing;
		if (Styles[currentStyle].UserDefined)
		{
			Btn_DelStyle.Enabled = true;
		}
		else
		{
			Btn_DelStyle.Enabled = false;
		}
		int num = 0;
		if (Styles[currentStyle].IsBold)
		{
			num = 1;
		}
		if (Styles[currentStyle].IsItalic)
		{
			num |= 2;
		}
		if (Styles[currentStyle].IsUnderline)
		{
			num |= 4;
		}
		try
		{
			Lab_StyleInfo.Font = new System.Drawing.Font(new FontFamily(Styles[currentStyle].FontName), Styles[currentStyle].FontSize, (FontStyle)num);
		}
		catch
		{
			Lab_StyleInfo.Font = new System.Drawing.Font("宋体", Styles[currentStyle].FontSize, (FontStyle)num);
		}
		Cmb_NumberStyle.SelectedIndex = Styles[currentStyle].NumberStyle;
		Txt_NumberFormat.Text = Styles[currentStyle].NumberFormat;
		Lab_StyleInfo.Text = Styles[currentStyle].StyleInfo();
		StyleSelectChanged = false;
	}

	private void Btn_DelStyle_Click(object sender, EventArgs e)
	{
		Styles.RemoveAt(currentStyle);
		StyleNames.Remove((string)Lst_Styles.SelectedItem);
		Lst_Styles.SelectedIndex = -1;
	}

	private void Txt_AddStyleName_Validating(object sender, CancelEventArgs e)
	{
		if (Regex.IsMatch(Txt_AddStyleName.Text, "[ ]{1,}"))
		{
			Txt_AddStyleName.Text = "";
		}
		if (StyleNames.Contains(Txt_AddStyleName.Text))
		{
			MessageBox.Show("该名称样式已存在，请更换名称", "提醒");
			Txt_AddStyleName.Text = "";
		}
		if (Txt_AddStyleName.Text != "")
		{
			Btn_AddStyle.Enabled = true;
		}
		else
		{
			Btn_AddStyle.Enabled = false;
		}
	}

	private void StyleFontChanged(object sender, EventArgs e)
	{
		if (StyleSelectChanged)
		{
			return;
		}
		ComboBox comboBox = sender as ComboBox;
		switch (comboBox.Name)
		{
		case "Cmb_FontName":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style.Name))
					{
						style.FontName = comboBox.Text;
					}
				}
			}
			else
			{
				Styles[currentStyle].FontName = comboBox.Text;
			}
			break;
		case "Cmb_FontSize":
		{
			float fontSize = ((comboBox.SelectedIndex == -1) ? Convert.ToSingle(comboBox.Text) : FontSizePoint[comboBox.SelectedIndex]);
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style2 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style2.Name))
					{
						style2.FontSize = fontSize;
					}
				}
			}
			else
			{
				Styles[currentStyle].FontSize = fontSize;
			}
			break;
		}
		case "Cmb_ParaAligment":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style3 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style3.Name))
					{
						style3.ParagraphAlignment = comboBox.SelectedIndex;
					}
				}
			}
			else
			{
				Styles[currentStyle].ParagraphAlignment = comboBox.SelectedIndex;
			}
			break;
		}
		Lab_StyleInfo.Text = Styles[currentStyle].StyleInfo();
		FontStyle fontStyle = FontStyle.Regular;
		if (Styles[currentStyle].IsBold)
		{
			fontStyle = FontStyle.Bold;
		}
		if (Styles[currentStyle].IsItalic)
		{
			fontStyle |= FontStyle.Italic;
		}
		if (Styles[currentStyle].IsUnderline)
		{
			fontStyle |= FontStyle.Underline;
		}
		try
		{
			Lab_StyleInfo.Font = new System.Drawing.Font(new FontFamily(Styles[currentStyle].FontName), Styles[currentStyle].FontSize, fontStyle);
		}
		catch
		{
			Lab_StyleInfo.Font = new System.Drawing.Font(new FontFamily("宋体"), Styles[currentStyle].FontSize, fontStyle);
		}
	}

	private void IndentSpacingChanged(object sender, EventArgs e)
	{
		if (StyleSelectChanged)
		{
			return;
		}
		NumericUpDownWithUnit numericUpDownWithUnit = sender as NumericUpDownWithUnit;
		switch (numericUpDownWithUnit.Name)
		{
		case "Nud_LeftIndent":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style.Name))
					{
						style.LeftIndent = Globals.ThisAddIn.Application.CentimetersToPoints((float)numericUpDownWithUnit.Value);
					}
				}
			}
			else
			{
				Styles[currentStyle].LeftIndent = Globals.ThisAddIn.Application.CentimetersToPoints((float)numericUpDownWithUnit.Value);
			}
			break;
		case "Nud_FirstLineIndent":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style2 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style2.Name))
					{
						style2.FirstLineIndent = Globals.ThisAddIn.Application.CentimetersToPoints((float)numericUpDownWithUnit.Value);
					}
				}
			}
			else
			{
				Styles[currentStyle].FirstLineIndent = Globals.ThisAddIn.Application.CentimetersToPoints((float)numericUpDownWithUnit.Value);
			}
			break;
		case "Nud_FirstLineIndentByChar":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style3 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style3.Name))
					{
						style3.FirstLineIndentByChar = (int)numericUpDownWithUnit.Value;
					}
				}
			}
			else
			{
				Styles[currentStyle].FirstLineIndentByChar = (int)numericUpDownWithUnit.Value;
			}
			break;
		case "Nud_LineSpacing":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style4 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style4.Name))
					{
						style4.LineSpacing = (float)numericUpDownWithUnit.Value;
					}
				}
			}
			else
			{
				Styles[currentStyle].LineSpacing = (float)numericUpDownWithUnit.Value;
			}
			break;
		case "Nud_BefreSpacing":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style5 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style5.Name))
					{
						style5.BeforeSpacing = (float)numericUpDownWithUnit.Value;
					}
				}
			}
			else
			{
				Styles[currentStyle].BeforeSpacing = (float)numericUpDownWithUnit.Value;
			}
			break;
		case "Nud_AfterSpacing":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style6 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style6.Name))
					{
						style6.AfterSpacing = (float)numericUpDownWithUnit.Value;
					}
				}
			}
			else
			{
				Styles[currentStyle].AfterSpacing = (float)numericUpDownWithUnit.Value;
			}
			break;
		}
		Lab_StyleInfo.Text = Styles[currentStyle].StyleInfo();
	}

	private void FontStyleChanged(object sender, EventArgs e)
	{
		if (StyleSelectChanged)
		{
			return;
		}
		CheckBox checkBox = sender as CheckBox;
		switch (checkBox.Name)
		{
		case "Chk_Bold":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style.Name))
					{
						style.IsBold = checkBox.Checked;
					}
				}
			}
			else
			{
				Styles[currentStyle].IsBold = checkBox.Checked;
			}
			break;
		case "Chk_Italic":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style2 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style2.Name))
					{
						style2.IsItalic = checkBox.Checked;
					}
				}
			}
			else
			{
				Styles[currentStyle].IsItalic = checkBox.Checked;
			}
			break;
		case "Chk_UnderLine":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style3 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style3.Name))
					{
						style3.IsUnderline = checkBox.Checked;
					}
				}
			}
			else
			{
				Styles[currentStyle].IsUnderline = checkBox.Checked;
			}
			break;
		case "Chk_BeforeBreak":
			if (Lst_Styles.SelectedItems.Count > 1)
			{
				foreach (CustomStyle style4 in Styles)
				{
					if (Lst_Styles.SelectedItems.Contains(style4.Name))
					{
						style4.BeforeBreak = checkBox.Checked;
					}
				}
			}
			else
			{
				Styles[currentStyle].BeforeBreak = checkBox.Checked;
			}
			break;
		}
		Lab_StyleInfo.Text = Styles[currentStyle].StyleInfo();
		FontStyle fontStyle = FontStyle.Regular;
		if (Styles[currentStyle].IsBold)
		{
			fontStyle = FontStyle.Bold;
		}
		if (Styles[currentStyle].IsItalic)
		{
			fontStyle |= FontStyle.Italic;
		}
		if (Styles[currentStyle].IsUnderline)
		{
			fontStyle |= FontStyle.Underline;
		}
		try
		{
			Lab_StyleInfo.Font = new System.Drawing.Font(new FontFamily(Styles[currentStyle].FontName), Styles[currentStyle].FontSize, fontStyle);
		}
		catch
		{
			Lab_StyleInfo.Font = new System.Drawing.Font(new FontFamily("宋体"), Styles[currentStyle].FontSize, fontStyle);
		}
	}

	private void Btn_ApplySet_Click(object sender, EventArgs e)
	{
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		Document document = ((!Rdo_NewDocument.Checked) ? Globals.ThisAddIn.Application.ActiveDocument : ((Document)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00020906-0000-0000-C000-000000000046")))));
		if (Chk_SetPage.Checked)
		{
			try
			{
				foreach (Section section in document.Sections)
				{
					section.PageSetup.PaperSize = PaperSize[Cmb_PaperSize.SelectedIndex];
					section.PageSetup.Orientation = ((Cmb_PaperDirection.SelectedIndex != 0) ? WdOrientation.wdOrientLandscape : WdOrientation.wdOrientPortrait);
				}
			}
			catch
			{
			}
			float[] pageMargin = ((Cmb_PageMargin.SelectedIndex == 1) ? new float[4] { 2f, 2f, 2f, 2f } : ((Cmb_PageMargin.SelectedIndex == 2) ? new float[4] { 2.5f, 2.5f, 2.5f, 2.5f } : ((Cmb_PageMargin.SelectedIndex == 3) ? new float[4] { 3f, 3f, 3f, 3f } : ((Cmb_PageMargin.SelectedIndex == 4) ? new float[4] { 3.5f, 3.5f, 3.5f, 3.5f } : ((Cmb_PageMargin.SelectedIndex != 5) ? new float[4] { defaultValue.PageTopMargin, defaultValue.PageBottomMargin, defaultValue.PageLeftMargin, defaultValue.PageRightMargin } : new float[4] { 3.7f, 3.5f, 2.8f, 2.6f })))));
			Globals.ThisAddIn.ApplyPageMargin(document, ApplyToSection: false, setPageMargin: true, pageMargin, Chk_SetGutter.Checked, Cmb_GutterPosition.SelectedIndex, (float)Nud_GutterValue.Value);
		}
		foreach (CustomStyle style in Styles)
		{
			if (StyleNames.Contains(style.Name))
			{
				if (style.Name == "正文")
				{
					Styles styles = document.Styles;
					object Index = "正文";
					styles[ref Index].Font.Size = style.FontSize;
					Globals.ThisAddIn.SetGrid(document, style.FontSize);
				}
				style.SetStyle(document);
			}
		}
		if (Chk_CreateListLevels.Checked)
		{
			Globals.ThisAddIn.AutoCreateLevelList(Cmb_SetLevel.SelectedIndex, numIndent, textIndent, afterIndent);
		}
		(base.Parent as Form).Close();
	}

	private void Cmb_NumberStyle_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (StyleSelectChanged)
		{
			return;
		}
		Txt_NumberFormat.Enabled = Cmb_NumberStyle.SelectedIndex != 0;
		if (Cmb_NumberStyle.SelectedIndex != 0 && Txt_NumberFormat.Text == "")
		{
			Txt_NumberFormat.Text = "%1";
		}
		if (Lst_Styles.SelectedItems.Count > 1)
		{
			foreach (CustomStyle style in Styles)
			{
				if (Lst_Styles.SelectedItems.Contains(style.Name) && Regex.IsMatch(style.Name, "标题 [1-9]|列表段落"))
				{
					style.NumberStyle = Cmb_NumberStyle.SelectedIndex;
				}
			}
			return;
		}
		Styles[currentStyle].NumberStyle = Cmb_NumberStyle.SelectedIndex;
	}

	private void Txt_NumberFormat_TextChanged(object sender, EventArgs e)
	{
		if (StyleSelectChanged)
		{
			return;
		}
		if (!Regex.IsMatch(Txt_NumberFormat.Text, ".*%1.*"))
		{
			MessageBox.Show("格式必须包含%1标题编号!", "提醒");
			return;
		}
		if (Lst_Styles.SelectedItems.Count > 1)
		{
			foreach (CustomStyle style in Styles)
			{
				if (Lst_Styles.SelectedItems.Contains(style.Name) && Regex.IsMatch(style.Name, "标题 [1-9]|列表段落"))
				{
					style.NumberFormat = Txt_NumberFormat.Text;
				}
			}
			return;
		}
		Styles[currentStyle].NumberFormat = Txt_NumberFormat.Text;
	}

	private void Btn_ReadDocumentStyle_Click(object sender, EventArgs e)
	{
		foreach (CustomStyle style2 in Styles)
		{
			try
			{
				Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				object Index = style2.Name;
				Style style = styles[ref Index];
				style2.FontName = style.Font.Name;
				style2.FontSize = style.Font.Size;
				style2.IsBold = style.Font.Bold == -1;
				style2.IsItalic = style.Font.Italic == -1;
				style2.IsUnderline = style.Font.Underline != WdUnderline.wdUnderlineNone;
				style2.FirstLineIndent = ((style.ParagraphFormat.FirstLineIndent > 0f) ? style.ParagraphFormat.FirstLineIndent : 0f);
				style2.LeftIndent = style.ParagraphFormat.LeftIndent;
				style2.FirstLineIndentByChar = (int)style.ParagraphFormat.CharacterUnitFirstLineIndent;
				style2.ParagraphAlignment = (int)style.ParagraphFormat.Alignment;
				style2.BeforeSpacing = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.SpaceBefore);
				style2.AfterSpacing = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.SpaceAfter);
				style2.LineSpacing = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.LineSpacing);
				style2.BeforeBreak = style.ParagraphFormat.PageBreakBefore == -1;
				if (Regex.IsMatch(style2.Name, "标题 [1-9]") && style.ListTemplate != null)
				{
					int num = Globals.ThisAddIn.LevelNumStyle.IndexOf(style.ListTemplate.ListLevels[1].NumberStyle);
					style2.NumberStyle = num + 1;
					style2.NumberFormat = style.ListTemplate.ListLevels[1].NumberFormat ?? "";
				}
			}
			catch
			{
			}
		}
	}

	private void Chk_SetPage_CheckedChanged(object sender, EventArgs e)
	{
		if (Chk_SetPage.Checked)
		{
			Grp_PageSetup.Enabled = true;
			if (Rdo_UseCurrentDocument.Checked)
			{
				Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
				if ((float)activeDocument.PageSetup.PaperSize == 9999999f || (float)activeDocument.PageSetup.Orientation == 9999999f || (float)activeDocument.PageSetup.GutterPos == 9999999f)
				{
					MessageBox.Show("当前文档中存在多个节，且节页面设置不同，开启页面设置会将所有节设置统一。请评估统一设置后对文档的影响！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				}
			}
		}
		else
		{
			Grp_PageSetup.Enabled = false;
		}
	}

	private void Cmb_FontSize_Leave(object sender, EventArgs e)
	{
		if (!StyleSelectChanged && Cmb_FontSize.SelectedIndex == -1 && Cmb_FontSize.Text != "" && Regex.IsMatch(Cmb_FontSize.Text, "^[1-9]{1,4}(\\.5|\\.0){0,1}$"))
		{
			float num = Convert.ToSingle(Cmb_FontSize.Text);
			if (num >= 1f || num <= 1638f)
			{
				Styles[currentStyle].FontSize = Convert.ToSingle(Cmb_FontSize.Text);
			}
		}
	}

	private void Cmb_PreSettings_SelectedIndexChanged(object sender, EventArgs e)
	{
		switch (Cmb_PreSettings.SelectedIndex)
		{
		case 0:
			Cmb_PaperSize.SelectedIndex = 1;
			Cmb_PaperDirection.SelectedIndex = 0;
			Cmb_SetLevel.SelectedIndex = 4;
			Cmb_PageMargin.SelectedIndex = 5;
			Chk_SetGutter.Checked = false;
			Cmb_GutterPosition.SelectedIndex = 0;
			Nud_GutterValue.Value = 0m;
			Chk_CreateListLevels.Checked = false;
			Styles.Where((CustomStyle c) => c.Name == "正文").First().SetValue("仿宋", 16f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.25f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 1").First().SetValue("黑体", 16f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.25f, beforebreak: false, 0f, 0f, 7, "%1、", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 2").First().SetValue("楷体", 16f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.25f, beforebreak: false, 0f, 0f, 7, "(%1)", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 3").First().SetValue("仿宋", 16f, bold: true, italic: false, underline: false, 0, 0f, 0f, 2, 1.25f, beforebreak: false, 0f, 0f, 1, "%1.", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 4").First().SetValue("仿宋", 16f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.25f, beforebreak: false, 0f, 0f, 1, "(%1)", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题").First().SetValue("华文中宋", 22f, bold: false, italic: false, underline: false, 1, 0f, 0f, 0, 1.25f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "副标题").First().SetValue("华文中宋", 16f, bold: false, italic: false, underline: false, 1, 0f, 0f, 0, 1.25f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "附录标题").First().SetValue("仿宋", 16f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.25f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表格标题").First().SetValue("仿宋", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.25f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "插图标题").First().SetValue("仿宋", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.25f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表内文字").First().SetValue("仿宋", 10.5f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			break;
		case 1:
			Cmb_PaperSize.SelectedIndex = 1;
			Cmb_PaperDirection.SelectedIndex = 0;
			Cmb_SetLevel.SelectedIndex = 4;
			Cmb_PageMargin.SelectedIndex = 2;
			Chk_SetGutter.Checked = true;
			Cmb_GutterPosition.SelectedIndex = 2;
			Nud_GutterValue.Value = 0.5m;
			Chk_CreateListLevels.Checked = false;
			Styles.Where((CustomStyle c) => c.Name == "正文").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 1").First().SetValue("黑体", 14f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: false, 0.5f, 0f, 7, "%1、", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 2").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: false, 0.5f, 0f, 7, "(%1)", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 3").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.3f, beforebreak: false, 0f, 0f, 1, "%1.", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 4").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.3f, beforebreak: false, 0f, 0f, 1, "(%1)", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题").First().SetValue("黑体", 22f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "副标题").First().SetValue("黑体", 16f, bold: false, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "附录标题").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表格标题").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "插图标题").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表内文字").First().SetValue("宋体", 10.5f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			break;
		case 2:
			Cmb_PaperSize.SelectedIndex = 1;
			Cmb_PaperDirection.SelectedIndex = 0;
			Cmb_SetLevel.SelectedIndex = 4;
			Cmb_PageMargin.SelectedIndex = 2;
			Chk_SetGutter.Checked = true;
			Cmb_GutterPosition.SelectedIndex = 2;
			Nud_GutterValue.Value = 0.5m;
			Chk_CreateListLevels.Checked = true;
			Styles.Where((CustomStyle c) => c.Name == "正文").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 1").First().SetValue("黑体", 16f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: true, 0.5f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 2").First().SetValue("黑体", 16f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: false, 0.5f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 3").First().SetValue("黑体", 15f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: false, 0.5f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 4").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 2, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题").First().SetValue("宋体", 24f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "副标题").First().SetValue("宋体", 16f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "附录标题").First().SetValue("黑体", 15f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表格标题").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "插图标题").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.3f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表内文字").First().SetValue("宋体", 10.5f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			break;
		case 3:
		{
			Cmb_PaperSize.SelectedIndex = 1;
			Cmb_PaperDirection.SelectedIndex = 0;
			Cmb_SetLevel.SelectedIndex = 4;
			Cmb_PageMargin.SelectedIndex = 2;
			Chk_SetGutter.Checked = true;
			Cmb_GutterPosition.SelectedIndex = 2;
			Nud_GutterValue.Value = 0.5m;
			Chk_CreateListLevels.Checked = true;
			float leftindent = Globals.ThisAddIn.Application.CentimetersToPoints(2.2f);
			Styles.Where((CustomStyle c) => c.Name == "正文").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, leftindent, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 1").First().SetValue("黑体", 14f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.35f, beforebreak: true, 0.5f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 2").First().SetValue("黑体", 14f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1.35f, beforebreak: false, 0.5f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 3").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 0, 0f, 0f, 0, 1.35f, beforebreak: false, 0.5f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题 4").First().SetValue("宋体", 12f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "标题").First().SetValue("宋体", 22f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "副标题").First().SetValue("宋体", 16f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "附录标题").First().SetValue("黑体", 14f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表格标题").First().SetValue("黑体", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "插图标题").First().SetValue("黑体", 12f, bold: true, italic: false, underline: false, 1, 0f, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "列表段落").First().SetValue("宋体", 12f, bold: true, italic: false, underline: false, 0, leftindent, 0f, 0, 1.35f, beforebreak: false, 0f, 0f, 1, "(%1)", userdefined: false);
			Styles.Where((CustomStyle c) => c.Name == "表内文字").First().SetValue("宋体", 10.5f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 1f, beforebreak: false, 0f, 0f, 0, "", userdefined: false);
			numIndent = 0f;
			textIndent = 2.2f;
			afterIndent = 2.2f;
			break;
		}
		}
		Lst_Styles.SelectedItems.Clear();
		Lst_Styles.SelectedIndex = 0;
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
		this.Rdo_NewDocument = new System.Windows.Forms.RadioButton();
		this.Rdo_UseCurrentDocument = new System.Windows.Forms.RadioButton();
		this.label1 = new System.Windows.Forms.Label();
		this.Cmb_PaperSize = new System.Windows.Forms.ComboBox();
		this.Grp_PageSetup = new System.Windows.Forms.GroupBox();
		this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
		this.label2 = new System.Windows.Forms.Label();
		this.Cmb_PaperDirection = new System.Windows.Forms.ComboBox();
		this.label5 = new System.Windows.Forms.Label();
		this.Cmb_PageMargin = new System.Windows.Forms.ComboBox();
		this.Chk_SetGutter = new System.Windows.Forms.CheckBox();
		this.Cmb_GutterPosition = new System.Windows.Forms.ComboBox();
		this.groupBox3 = new System.Windows.Forms.GroupBox();
		this.Btn_ReadDocumentStyle = new System.Windows.Forms.Button();
		this.Pal_NumberList = new System.Windows.Forms.FlowLayoutPanel();
		this.label13 = new System.Windows.Forms.Label();
		this.Cmb_NumberStyle = new System.Windows.Forms.ComboBox();
		this.label14 = new System.Windows.Forms.Label();
		this.Txt_NumberFormat = new System.Windows.Forms.TextBox();
		this.Lab_StyleInfo = new System.Windows.Forms.Label();
		this.Btn_DelStyle = new System.Windows.Forms.Button();
		this.Pal_ParaSpacing = new System.Windows.Forms.FlowLayoutPanel();
		this.label9 = new System.Windows.Forms.Label();
		this.label10 = new System.Windows.Forms.Label();
		this.Chk_BeforeBreak = new System.Windows.Forms.CheckBox();
		this.label11 = new System.Windows.Forms.Label();
		this.Pal_Font = new System.Windows.Forms.FlowLayoutPanel();
		this.label3 = new System.Windows.Forms.Label();
		this.Cmb_FontName = new System.Windows.Forms.ComboBox();
		this.label4 = new System.Windows.Forms.Label();
		this.Cmb_FontSize = new System.Windows.Forms.ComboBox();
		this.Chk_Bold = new System.Windows.Forms.CheckBox();
		this.Chk_Italic = new System.Windows.Forms.CheckBox();
		this.Chk_UnderLine = new System.Windows.Forms.CheckBox();
		this.Pal_ParaIndent = new System.Windows.Forms.FlowLayoutPanel();
		this.label15 = new System.Windows.Forms.Label();
		this.Cmb_ParaAligment = new System.Windows.Forms.ComboBox();
		this.label6 = new System.Windows.Forms.Label();
		this.label7 = new System.Windows.Forms.Label();
		this.label8 = new System.Windows.Forms.Label();
		this.Btn_AddStyle = new System.Windows.Forms.Button();
		this.Chk_CreateListLevels = new System.Windows.Forms.CheckBox();
		this.Cmb_SetLevel = new System.Windows.Forms.ComboBox();
		this.label12 = new System.Windows.Forms.Label();
		this.Lst_Styles = new System.Windows.Forms.ListBox();
		this.Btn_ApplySet = new System.Windows.Forms.Button();
		this.Cmb_PreSettings = new System.Windows.Forms.ComboBox();
		this.label16 = new System.Windows.Forms.Label();
		this.Txt_AddStyleName = new WordFormatHelper.WaterMarkTextBoxEx();
		this.Nud_LineSpacing = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_BefreSpacing = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_AfterSpacing = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_LeftIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_FirstLineIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_FirstLineIndentByChar = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_GutterValue = new WordFormatHelper.NumericUpDownWithUnit();
		this.Chk_SetPage = new System.Windows.Forms.CheckBox();
		this.groupBox1.SuspendLayout();
		this.Grp_PageSetup.SuspendLayout();
		this.flowLayoutPanel2.SuspendLayout();
		this.groupBox3.SuspendLayout();
		this.Pal_NumberList.SuspendLayout();
		this.Pal_ParaSpacing.SuspendLayout();
		this.Pal_Font.SuspendLayout();
		this.Pal_ParaIndent.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_LineSpacing).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_BefreSpacing).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_AfterSpacing).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_LeftIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_FirstLineIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_FirstLineIndentByChar).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_GutterValue).BeginInit();
		base.SuspendLayout();
		this.groupBox1.Controls.Add(this.Chk_SetPage);
		this.groupBox1.Controls.Add(this.Rdo_NewDocument);
		this.groupBox1.Controls.Add(this.Rdo_UseCurrentDocument);
		this.groupBox1.Location = new System.Drawing.Point(3, 3);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(127, 138);
		this.groupBox1.TabIndex = 0;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "文件";
		this.Rdo_NewDocument.AutoSize = true;
		this.Rdo_NewDocument.Location = new System.Drawing.Point(12, 56);
		this.Rdo_NewDocument.Name = "Rdo_NewDocument";
		this.Rdo_NewDocument.Size = new System.Drawing.Size(97, 24);
		this.Rdo_NewDocument.TabIndex = 1;
		this.Rdo_NewDocument.Text = "创建新文档";
		this.Rdo_NewDocument.UseVisualStyleBackColor = true;
		this.Rdo_UseCurrentDocument.AutoSize = true;
		this.Rdo_UseCurrentDocument.Checked = true;
		this.Rdo_UseCurrentDocument.Location = new System.Drawing.Point(12, 24);
		this.Rdo_UseCurrentDocument.Name = "Rdo_UseCurrentDocument";
		this.Rdo_UseCurrentDocument.Size = new System.Drawing.Size(111, 24);
		this.Rdo_UseCurrentDocument.TabIndex = 0;
		this.Rdo_UseCurrentDocument.TabStop = true;
		this.Rdo_UseCurrentDocument.Text = "设置当前文档";
		this.Rdo_UseCurrentDocument.UseVisualStyleBackColor = true;
		this.label1.Location = new System.Drawing.Point(3, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(65, 30);
		this.label1.TabIndex = 3;
		this.label1.Text = "页面大小";
		this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_PaperSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PaperSize.FormattingEnabled = true;
		this.Cmb_PaperSize.Items.AddRange(new object[4] { "A3", "A4", "A5", "B5" });
		this.Cmb_PaperSize.Location = new System.Drawing.Point(74, 3);
		this.Cmb_PaperSize.Name = "Cmb_PaperSize";
		this.Cmb_PaperSize.Size = new System.Drawing.Size(65, 28);
		this.Cmb_PaperSize.TabIndex = 4;
		this.Grp_PageSetup.Controls.Add(this.flowLayoutPanel2);
		this.Grp_PageSetup.Enabled = false;
		this.Grp_PageSetup.Location = new System.Drawing.Point(138, 3);
		this.Grp_PageSetup.Name = "Grp_PageSetup";
		this.Grp_PageSetup.Size = new System.Drawing.Size(275, 138);
		this.Grp_PageSetup.TabIndex = 5;
		this.Grp_PageSetup.TabStop = false;
		this.Grp_PageSetup.Text = "页面";
		this.flowLayoutPanel2.Controls.Add(this.label1);
		this.flowLayoutPanel2.Controls.Add(this.Cmb_PaperSize);
		this.flowLayoutPanel2.Controls.Add(this.label2);
		this.flowLayoutPanel2.Controls.Add(this.Cmb_PaperDirection);
		this.flowLayoutPanel2.Controls.Add(this.label5);
		this.flowLayoutPanel2.Controls.Add(this.Cmb_PageMargin);
		this.flowLayoutPanel2.Controls.Add(this.Chk_SetGutter);
		this.flowLayoutPanel2.Controls.Add(this.Cmb_GutterPosition);
		this.flowLayoutPanel2.Controls.Add(this.Nud_GutterValue);
		this.flowLayoutPanel2.Location = new System.Drawing.Point(6, 25);
		this.flowLayoutPanel2.Name = "flowLayoutPanel2";
		this.flowLayoutPanel2.Size = new System.Drawing.Size(256, 106);
		this.flowLayoutPanel2.TabIndex = 7;
		this.label2.Location = new System.Drawing.Point(145, 0);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(37, 30);
		this.label2.TabIndex = 5;
		this.label2.Text = "方向";
		this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_PaperDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PaperDirection.FormattingEnabled = true;
		this.Cmb_PaperDirection.Items.AddRange(new object[2] { "竖向", "横向" });
		this.Cmb_PaperDirection.Location = new System.Drawing.Point(188, 3);
		this.Cmb_PaperDirection.Name = "Cmb_PaperDirection";
		this.Cmb_PaperDirection.Size = new System.Drawing.Size(60, 28);
		this.Cmb_PaperDirection.TabIndex = 6;
		this.label5.Location = new System.Drawing.Point(3, 34);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(65, 30);
		this.label5.TabIndex = 7;
		this.label5.Text = "页边距";
		this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_PageMargin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PageMargin.FormattingEnabled = true;
		this.Cmb_PageMargin.Items.AddRange(new object[6] { "使用默认值", "全部设置2.0厘米", "全部设置2.5厘米", "全部设置3.0厘米", "全部设置3.5厘米", "上3.7厘米;左2.8厘米" });
		this.Cmb_PageMargin.Location = new System.Drawing.Point(74, 37);
		this.Cmb_PageMargin.Name = "Cmb_PageMargin";
		this.Cmb_PageMargin.Size = new System.Drawing.Size(175, 28);
		this.Cmb_PageMargin.TabIndex = 8;
		this.Chk_SetGutter.Location = new System.Drawing.Point(3, 71);
		this.Chk_SetGutter.Name = "Chk_SetGutter";
		this.Chk_SetGutter.Size = new System.Drawing.Size(70, 30);
		this.Chk_SetGutter.TabIndex = 9;
		this.Chk_SetGutter.Text = "装订线";
		this.Chk_SetGutter.UseVisualStyleBackColor = true;
		this.Chk_SetGutter.CheckedChanged += new System.EventHandler(Chk_SetGutter_CheckedChanged);
		this.Cmb_GutterPosition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_GutterPosition.Enabled = false;
		this.Cmb_GutterPosition.FormattingEnabled = true;
		this.Cmb_GutterPosition.Items.AddRange(new object[3] { "左", "上", "对称" });
		this.Cmb_GutterPosition.Location = new System.Drawing.Point(79, 71);
		this.Cmb_GutterPosition.Name = "Cmb_GutterPosition";
		this.Cmb_GutterPosition.Size = new System.Drawing.Size(74, 28);
		this.Cmb_GutterPosition.TabIndex = 10;
		this.groupBox3.Controls.Add(this.Txt_AddStyleName);
		this.groupBox3.Controls.Add(this.Btn_ReadDocumentStyle);
		this.groupBox3.Controls.Add(this.Pal_NumberList);
		this.groupBox3.Controls.Add(this.Lab_StyleInfo);
		this.groupBox3.Controls.Add(this.Btn_DelStyle);
		this.groupBox3.Controls.Add(this.Pal_ParaSpacing);
		this.groupBox3.Controls.Add(this.Pal_Font);
		this.groupBox3.Controls.Add(this.Pal_ParaIndent);
		this.groupBox3.Controls.Add(this.Btn_AddStyle);
		this.groupBox3.Controls.Add(this.Chk_CreateListLevels);
		this.groupBox3.Controls.Add(this.Cmb_SetLevel);
		this.groupBox3.Controls.Add(this.label12);
		this.groupBox3.Controls.Add(this.Lst_Styles);
		this.groupBox3.Location = new System.Drawing.Point(3, 147);
		this.groupBox3.Name = "groupBox3";
		this.groupBox3.Size = new System.Drawing.Size(410, 559);
		this.groupBox3.TabIndex = 6;
		this.groupBox3.TabStop = false;
		this.groupBox3.Text = "样式设置";
		this.Btn_ReadDocumentStyle.Location = new System.Drawing.Point(12, 25);
		this.Btn_ReadDocumentStyle.Name = "Btn_ReadDocumentStyle";
		this.Btn_ReadDocumentStyle.Size = new System.Drawing.Size(106, 29);
		this.Btn_ReadDocumentStyle.TabIndex = 32;
		this.Btn_ReadDocumentStyle.Text = "读取文中样式";
		this.Btn_ReadDocumentStyle.UseVisualStyleBackColor = true;
		this.Btn_ReadDocumentStyle.Click += new System.EventHandler(Btn_ReadDocumentStyle_Click);
		this.Pal_NumberList.Controls.Add(this.label13);
		this.Pal_NumberList.Controls.Add(this.Cmb_NumberStyle);
		this.Pal_NumberList.Controls.Add(this.label14);
		this.Pal_NumberList.Controls.Add(this.Txt_NumberFormat);
		this.Pal_NumberList.Enabled = false;
		this.Pal_NumberList.Location = new System.Drawing.Point(182, 418);
		this.Pal_NumberList.Name = "Pal_NumberList";
		this.Pal_NumberList.Size = new System.Drawing.Size(215, 74);
		this.Pal_NumberList.TabIndex = 31;
		this.label13.Location = new System.Drawing.Point(3, 0);
		this.label13.Name = "label13";
		this.label13.Size = new System.Drawing.Size(80, 28);
		this.label13.TabIndex = 1;
		this.label13.Text = "编号样式";
		this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_NumberStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_NumberStyle.FormattingEnabled = true;
		this.Cmb_NumberStyle.Items.AddRange(new object[11]
		{
			"无编号", "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,貳,叁...", "甲,乙,丙...",
			"正规编号"
		});
		this.Cmb_NumberStyle.Location = new System.Drawing.Point(89, 3);
		this.Cmb_NumberStyle.Name = "Cmb_NumberStyle";
		this.Cmb_NumberStyle.Size = new System.Drawing.Size(120, 28);
		this.Cmb_NumberStyle.TabIndex = 2;
		this.Cmb_NumberStyle.SelectedIndexChanged += new System.EventHandler(Cmb_NumberStyle_SelectedIndexChanged);
		this.label14.Location = new System.Drawing.Point(3, 34);
		this.label14.Name = "label14";
		this.label14.Size = new System.Drawing.Size(80, 28);
		this.label14.TabIndex = 3;
		this.label14.Text = "编号格式";
		this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Txt_NumberFormat.Enabled = false;
		this.Txt_NumberFormat.Location = new System.Drawing.Point(89, 37);
		this.Txt_NumberFormat.Name = "Txt_NumberFormat";
		this.Txt_NumberFormat.Size = new System.Drawing.Size(120, 26);
		this.Txt_NumberFormat.TabIndex = 4;
		this.Txt_NumberFormat.TextChanged += new System.EventHandler(Txt_NumberFormat_TextChanged);
		this.Lab_StyleInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
		this.Lab_StyleInfo.Location = new System.Drawing.Point(12, 496);
		this.Lab_StyleInfo.Name = "Lab_StyleInfo";
		this.Lab_StyleInfo.Size = new System.Drawing.Size(384, 53);
		this.Lab_StyleInfo.TabIndex = 30;
		this.Lab_StyleInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Btn_DelStyle.Enabled = false;
		this.Btn_DelStyle.Location = new System.Drawing.Point(96, 463);
		this.Btn_DelStyle.Name = "Btn_DelStyle";
		this.Btn_DelStyle.Size = new System.Drawing.Size(76, 29);
		this.Btn_DelStyle.TabIndex = 29;
		this.Btn_DelStyle.Text = "删除样式";
		this.Btn_DelStyle.UseVisualStyleBackColor = true;
		this.Btn_DelStyle.Click += new System.EventHandler(Btn_DelStyle_Click);
		this.Pal_ParaSpacing.Controls.Add(this.label9);
		this.Pal_ParaSpacing.Controls.Add(this.Nud_LineSpacing);
		this.Pal_ParaSpacing.Controls.Add(this.label10);
		this.Pal_ParaSpacing.Controls.Add(this.Nud_BefreSpacing);
		this.Pal_ParaSpacing.Controls.Add(this.Chk_BeforeBreak);
		this.Pal_ParaSpacing.Controls.Add(this.label11);
		this.Pal_ParaSpacing.Controls.Add(this.Nud_AfterSpacing);
		this.Pal_ParaSpacing.Enabled = false;
		this.Pal_ParaSpacing.Location = new System.Drawing.Point(181, 315);
		this.Pal_ParaSpacing.Name = "Pal_ParaSpacing";
		this.Pal_ParaSpacing.Size = new System.Drawing.Size(215, 97);
		this.Pal_ParaSpacing.TabIndex = 8;
		this.label9.Location = new System.Drawing.Point(3, 0);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(65, 30);
		this.label9.TabIndex = 15;
		this.label9.Text = "段落行距";
		this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label10.Location = new System.Drawing.Point(3, 32);
		this.label10.Name = "label10";
		this.label10.Size = new System.Drawing.Size(65, 30);
		this.label10.TabIndex = 17;
		this.label10.Text = "段前间距";
		this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Chk_BeforeBreak.AutoSize = true;
		this.Chk_BeforeBreak.Location = new System.Drawing.Point(150, 35);
		this.Chk_BeforeBreak.Name = "Chk_BeforeBreak";
		this.Chk_BeforeBreak.Size = new System.Drawing.Size(56, 24);
		this.Chk_BeforeBreak.TabIndex = 20;
		this.Chk_BeforeBreak.Text = "分页";
		this.Chk_BeforeBreak.UseVisualStyleBackColor = true;
		this.label11.Location = new System.Drawing.Point(3, 64);
		this.label11.Name = "label11";
		this.label11.Size = new System.Drawing.Size(65, 30);
		this.label11.TabIndex = 19;
		this.label11.Text = "段后间距";
		this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Pal_Font.Controls.Add(this.label3);
		this.Pal_Font.Controls.Add(this.Cmb_FontName);
		this.Pal_Font.Controls.Add(this.label4);
		this.Pal_Font.Controls.Add(this.Cmb_FontSize);
		this.Pal_Font.Controls.Add(this.Chk_Bold);
		this.Pal_Font.Controls.Add(this.Chk_Italic);
		this.Pal_Font.Controls.Add(this.Chk_UnderLine);
		this.Pal_Font.Enabled = false;
		this.Pal_Font.Location = new System.Drawing.Point(182, 63);
		this.Pal_Font.Name = "Pal_Font";
		this.Pal_Font.Size = new System.Drawing.Size(215, 100);
		this.Pal_Font.TabIndex = 7;
		this.label3.Location = new System.Drawing.Point(3, 0);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(40, 28);
		this.label3.TabIndex = 1;
		this.label3.Text = "字体";
		this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_FontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_FontName.FormattingEnabled = true;
		this.Cmb_FontName.Location = new System.Drawing.Point(49, 3);
		this.Cmb_FontName.Name = "Cmb_FontName";
		this.Cmb_FontName.Size = new System.Drawing.Size(160, 28);
		this.Cmb_FontName.TabIndex = 2;
		this.label4.Location = new System.Drawing.Point(3, 34);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(80, 28);
		this.label4.TabIndex = 3;
		this.label4.Text = "大小";
		this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_FontSize.FormattingEnabled = true;
		this.Cmb_FontSize.Location = new System.Drawing.Point(89, 37);
		this.Cmb_FontSize.Name = "Cmb_FontSize";
		this.Cmb_FontSize.Size = new System.Drawing.Size(120, 28);
		this.Cmb_FontSize.TabIndex = 4;
		this.Cmb_FontSize.Leave += new System.EventHandler(Cmb_FontSize_Leave);
		this.Chk_Bold.Location = new System.Drawing.Point(3, 71);
		this.Chk_Bold.Name = "Chk_Bold";
		this.Chk_Bold.Size = new System.Drawing.Size(60, 30);
		this.Chk_Bold.TabIndex = 5;
		this.Chk_Bold.Text = "粗体";
		this.Chk_Bold.UseVisualStyleBackColor = true;
		this.Chk_Italic.Location = new System.Drawing.Point(69, 71);
		this.Chk_Italic.Name = "Chk_Italic";
		this.Chk_Italic.Size = new System.Drawing.Size(60, 30);
		this.Chk_Italic.TabIndex = 6;
		this.Chk_Italic.Text = "斜体";
		this.Chk_Italic.UseVisualStyleBackColor = true;
		this.Chk_UnderLine.Location = new System.Drawing.Point(135, 71);
		this.Chk_UnderLine.Name = "Chk_UnderLine";
		this.Chk_UnderLine.Size = new System.Drawing.Size(70, 30);
		this.Chk_UnderLine.TabIndex = 7;
		this.Chk_UnderLine.Text = "下划线";
		this.Chk_UnderLine.UseVisualStyleBackColor = true;
		this.Pal_ParaIndent.Controls.Add(this.label15);
		this.Pal_ParaIndent.Controls.Add(this.Cmb_ParaAligment);
		this.Pal_ParaIndent.Controls.Add(this.label6);
		this.Pal_ParaIndent.Controls.Add(this.Nud_LeftIndent);
		this.Pal_ParaIndent.Controls.Add(this.label7);
		this.Pal_ParaIndent.Controls.Add(this.Nud_FirstLineIndent);
		this.Pal_ParaIndent.Controls.Add(this.label8);
		this.Pal_ParaIndent.Controls.Add(this.Nud_FirstLineIndentByChar);
		this.Pal_ParaIndent.Enabled = false;
		this.Pal_ParaIndent.Location = new System.Drawing.Point(182, 169);
		this.Pal_ParaIndent.Name = "Pal_ParaIndent";
		this.Pal_ParaIndent.Size = new System.Drawing.Size(215, 140);
		this.Pal_ParaIndent.TabIndex = 7;
		this.label15.Location = new System.Drawing.Point(3, 0);
		this.label15.Name = "label15";
		this.label15.Size = new System.Drawing.Size(100, 30);
		this.label15.TabIndex = 14;
		this.label15.Text = "段落对齐方式";
		this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_ParaAligment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_ParaAligment.FormattingEnabled = true;
		this.Cmb_ParaAligment.Items.AddRange(new object[5] { "左对齐", "居中对齐", "右对齐", "两端对齐", "分散对齐" });
		this.Cmb_ParaAligment.Location = new System.Drawing.Point(109, 3);
		this.Cmb_ParaAligment.Name = "Cmb_ParaAligment";
		this.Cmb_ParaAligment.Size = new System.Drawing.Size(100, 28);
		this.Cmb_ParaAligment.TabIndex = 15;
		this.label6.Location = new System.Drawing.Point(3, 34);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(100, 30);
		this.label6.TabIndex = 8;
		this.label6.Text = "段落左缩进";
		this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label7.Location = new System.Drawing.Point(3, 66);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(100, 30);
		this.label7.TabIndex = 11;
		this.label7.Text = "段落首行缩进";
		this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label8.Location = new System.Drawing.Point(3, 98);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(100, 30);
		this.label8.TabIndex = 13;
		this.label8.Text = "首行字符缩进";
		this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Btn_AddStyle.Enabled = false;
		this.Btn_AddStyle.Location = new System.Drawing.Point(14, 463);
		this.Btn_AddStyle.Name = "Btn_AddStyle";
		this.Btn_AddStyle.Size = new System.Drawing.Size(76, 29);
		this.Btn_AddStyle.TabIndex = 27;
		this.Btn_AddStyle.Text = "添加样式";
		this.Btn_AddStyle.UseVisualStyleBackColor = true;
		this.Btn_AddStyle.Click += new System.EventHandler(Btn_AddStyle_Click);
		this.Chk_CreateListLevels.AutoSize = true;
		this.Chk_CreateListLevels.Location = new System.Drawing.Point(285, 27);
		this.Chk_CreateListLevels.Name = "Chk_CreateListLevels";
		this.Chk_CreateListLevels.Size = new System.Drawing.Size(112, 24);
		this.Chk_CreateListLevels.TabIndex = 22;
		this.Chk_CreateListLevels.Text = "创建多级列表";
		this.Chk_CreateListLevels.UseVisualStyleBackColor = true;
		this.Cmb_SetLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_SetLevel.FormattingEnabled = true;
		this.Cmb_SetLevel.Items.AddRange(new object[10] { "无", "1", "2", "3", "4", "5", "6", "7", "8", "9" });
		this.Cmb_SetLevel.Location = new System.Drawing.Point(220, 25);
		this.Cmb_SetLevel.Name = "Cmb_SetLevel";
		this.Cmb_SetLevel.Size = new System.Drawing.Size(59, 28);
		this.Cmb_SetLevel.TabIndex = 21;
		this.Cmb_SetLevel.SelectedIndexChanged += new System.EventHandler(Cmb_SetLevel_SelectedIndexChanged);
		this.label12.AutoSize = true;
		this.label12.Location = new System.Drawing.Point(139, 29);
		this.label12.Name = "label12";
		this.label12.Size = new System.Drawing.Size(79, 20);
		this.label12.TabIndex = 20;
		this.label12.Text = "显示标题数";
		this.Lst_Styles.FormattingEnabled = true;
		this.Lst_Styles.ItemHeight = 20;
		this.Lst_Styles.Location = new System.Drawing.Point(14, 63);
		this.Lst_Styles.Name = "Lst_Styles";
		this.Lst_Styles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
		this.Lst_Styles.Size = new System.Drawing.Size(158, 364);
		this.Lst_Styles.TabIndex = 0;
		this.Btn_ApplySet.Location = new System.Drawing.Point(258, 712);
		this.Btn_ApplySet.Name = "Btn_ApplySet";
		this.Btn_ApplySet.Size = new System.Drawing.Size(155, 35);
		this.Btn_ApplySet.TabIndex = 7;
		this.Btn_ApplySet.Text = "设置/创建文档";
		this.Btn_ApplySet.UseVisualStyleBackColor = true;
		this.Btn_ApplySet.Click += new System.EventHandler(Btn_ApplySet_Click);
		this.Cmb_PreSettings.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PreSettings.FormattingEnabled = true;
		this.Cmb_PreSettings.Items.AddRange(new object[4] { "公文风格", "论文风格", "报告风格", "条文风格" });
		this.Cmb_PreSettings.Location = new System.Drawing.Point(86, 716);
		this.Cmb_PreSettings.Name = "Cmb_PreSettings";
		this.Cmb_PreSettings.Size = new System.Drawing.Size(120, 28);
		this.Cmb_PreSettings.TabIndex = 8;
		this.Cmb_PreSettings.SelectedIndexChanged += new System.EventHandler(Cmb_PreSettings_SelectedIndexChanged);
		this.label16.AutoSize = true;
		this.label16.Location = new System.Drawing.Point(14, 719);
		this.label16.Name = "label16";
		this.label16.Size = new System.Drawing.Size(65, 20);
		this.label16.TabIndex = 21;
		this.label16.Text = "预设样式";
		this.Txt_AddStyleName.Location = new System.Drawing.Point(14, 433);
		this.Txt_AddStyleName.Name = "Txt_AddStyleName";
		this.Txt_AddStyleName.Size = new System.Drawing.Size(158, 26);
		this.Txt_AddStyleName.TabIndex = 33;
		this.Txt_AddStyleName.WaterMark = "输入添加样式的名称";
		this.Txt_AddStyleName.WaterTextColor = System.Drawing.Color.SteelBlue;
		this.Txt_AddStyleName.Validating += new System.ComponentModel.CancelEventHandler(Txt_AddStyleName_Validating);
		this.Nud_LineSpacing.DecimalPlaces = 2;
		this.Nud_LineSpacing.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_LineSpacing.Label = "行";
		this.Nud_LineSpacing.Location = new System.Drawing.Point(74, 3);
		this.Nud_LineSpacing.Minimum = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Nud_LineSpacing.Name = "Nud_LineSpacing";
		this.Nud_LineSpacing.Size = new System.Drawing.Size(100, 26);
		this.Nud_LineSpacing.TabIndex = 14;
		this.Nud_LineSpacing.Value = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Nud_BefreSpacing.DecimalPlaces = 2;
		this.Nud_BefreSpacing.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_BefreSpacing.Label = "行";
		this.Nud_BefreSpacing.Location = new System.Drawing.Point(74, 35);
		this.Nud_BefreSpacing.Name = "Nud_BefreSpacing";
		this.Nud_BefreSpacing.Size = new System.Drawing.Size(70, 26);
		this.Nud_BefreSpacing.TabIndex = 16;
		this.Nud_AfterSpacing.DecimalPlaces = 2;
		this.Nud_AfterSpacing.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_AfterSpacing.Label = "行";
		this.Nud_AfterSpacing.Location = new System.Drawing.Point(74, 67);
		this.Nud_AfterSpacing.Name = "Nud_AfterSpacing";
		this.Nud_AfterSpacing.Size = new System.Drawing.Size(100, 26);
		this.Nud_AfterSpacing.TabIndex = 18;
		this.Nud_LeftIndent.DecimalPlaces = 2;
		this.Nud_LeftIndent.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_LeftIndent.Label = "厘米";
		this.Nud_LeftIndent.Location = new System.Drawing.Point(109, 37);
		this.Nud_LeftIndent.Name = "Nud_LeftIndent";
		this.Nud_LeftIndent.Size = new System.Drawing.Size(100, 26);
		this.Nud_LeftIndent.TabIndex = 9;
		this.Nud_FirstLineIndent.DecimalPlaces = 2;
		this.Nud_FirstLineIndent.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_FirstLineIndent.Label = "厘米";
		this.Nud_FirstLineIndent.Location = new System.Drawing.Point(109, 69);
		this.Nud_FirstLineIndent.Name = "Nud_FirstLineIndent";
		this.Nud_FirstLineIndent.Size = new System.Drawing.Size(100, 26);
		this.Nud_FirstLineIndent.TabIndex = 10;
		this.Nud_FirstLineIndentByChar.Label = "字符";
		this.Nud_FirstLineIndentByChar.Location = new System.Drawing.Point(109, 101);
		this.Nud_FirstLineIndentByChar.Name = "Nud_FirstLineIndentByChar";
		this.Nud_FirstLineIndentByChar.Size = new System.Drawing.Size(100, 26);
		this.Nud_FirstLineIndentByChar.TabIndex = 12;
		this.Nud_GutterValue.DecimalPlaces = 2;
		this.Nud_GutterValue.Enabled = false;
		this.Nud_GutterValue.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_GutterValue.Label = "厘米";
		this.Nud_GutterValue.Location = new System.Drawing.Point(159, 71);
		this.Nud_GutterValue.Name = "Nud_GutterValue";
		this.Nud_GutterValue.Size = new System.Drawing.Size(89, 26);
		this.Nud_GutterValue.TabIndex = 11;
		this.Chk_SetPage.AutoSize = true;
		this.Chk_SetPage.Location = new System.Drawing.Point(12, 96);
		this.Chk_SetPage.Name = "Chk_SetPage";
		this.Chk_SetPage.Size = new System.Drawing.Size(84, 24);
		this.Chk_SetPage.TabIndex = 2;
		this.Chk_SetPage.Text = "设置页面";
		this.Chk_SetPage.UseVisualStyleBackColor = true;
		this.Chk_SetPage.CheckedChanged += new System.EventHandler(Chk_SetPage_CheckedChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.label16);
		base.Controls.Add(this.Cmb_PreSettings);
		base.Controls.Add(this.Btn_ApplySet);
		base.Controls.Add(this.groupBox3);
		base.Controls.Add(this.Grp_PageSetup);
		base.Controls.Add(this.groupBox1);
		this.DoubleBuffered = true;
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Name = "StyleSetGuider";
		base.Size = new System.Drawing.Size(420, 750);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.Grp_PageSetup.ResumeLayout(false);
		this.flowLayoutPanel2.ResumeLayout(false);
		this.groupBox3.ResumeLayout(false);
		this.groupBox3.PerformLayout();
		this.Pal_NumberList.ResumeLayout(false);
		this.Pal_NumberList.PerformLayout();
		this.Pal_ParaSpacing.ResumeLayout(false);
		this.Pal_ParaSpacing.PerformLayout();
		this.Pal_Font.ResumeLayout(false);
		this.Pal_ParaIndent.ResumeLayout(false);
		((System.ComponentModel.ISupportInitialize)this.Nud_LineSpacing).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_BefreSpacing).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_AfterSpacing).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_LeftIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_FirstLineIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_FirstLineIndentByChar).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_GutterValue).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}