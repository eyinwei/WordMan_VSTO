// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.StyleSetGuider
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
using WordFormatHelper;

public class StyleSetGuider : UserControl
{
	private readonly List<WdPaperSize> PaperSize = new List<WdPaperSize>(21)
	{
		WdPaperSize.wdPaperA3,
		WdPaperSize.wdPaperA4,
		WdPaperSize.wdPaperA4Small,
		WdPaperSize.wdPaperA5,
		WdPaperSize.wdPaperB4,
		WdPaperSize.wdPaperB5,
		WdPaperSize.wdPaperEnvelope10,
		WdPaperSize.wdPaperEnvelopeB4,
		WdPaperSize.wdPaperEnvelopeB5,
		WdPaperSize.wdPaperEnvelopeB6,
		WdPaperSize.wdPaperEnvelopeC3,
		WdPaperSize.wdPaperEnvelopeC4,
		WdPaperSize.wdPaperEnvelopeC5,
		WdPaperSize.wdPaperEnvelopeC6,
		WdPaperSize.wdPaperFolio,
		WdPaperSize.wdPaperLedger,
		WdPaperSize.wdPaperLegal,
		WdPaperSize.wdPaperLetter,
		WdPaperSize.wdPaperLetterSmall,
		WdPaperSize.wdPaperNote,
		WdPaperSize.wdPaperQuarto
	};

	private static List<WordStyleInfo> Styles = new List<WordStyleInfo>();

	private static readonly List<string> AllBuildInStyleNames = new List<string>();

	private static readonly List<WordStyleInfo> BuildInStyles = new List<WordStyleInfo>();

	private static readonly List<WordStyleInfo> CustomStyles = new List<WordStyleInfo>();

	private bool userChanged;

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

	private ComboBox Cmb_PageMargin;

	private Label label5;

	private Label label6;

	private Label label7;

	private Label label11;

	private Label label10;

	private Label label9;

	private ComboBox Cmb_SetLevel;

	private Label label12;

	private ComboBox Cmb_GutterPosition;

	private CheckBox Chk_SetGutter;

	private NumericUpDownWithUnit Nud_GutterValue;

	private CheckBox Chk_CreateListLevels;

	private Button Btn_AddStyle;

	private Label label3;

	private ComboBox Cmb_ChnFontName;

	private Button Btn_DelStyle;

	private Label Lab_StyleInfo;

	private Button Btn_ApplySet;

	private Label label13;

	private ComboBox Cmb_NumberStyle;

	private Label label14;

	private TextBox Txt_NumberFormat;

	private Label label15;

	private ComboBox Cmb_ParaAligment;

	private CheckBox Chk_BeforeBreak;

	private ComboBox Cmb_PreSettings;

	private Label label16;

	private WaterMarkTextBoxEx Txt_NewStyleName;

	private CheckBox Chk_SetPage;

	private GroupBox groupBox2;

	private ComboBox Cmb_EngFontName;

	private Label label17;

	private Button Btn_FontColor;

	private Label label18;

	private ComboBox Cmb_SpaceAfter;

	private ComboBox Cmb_SpaceBefore;

	private ComboBox Cmb_LineSpace;

	private TextBox Txt_FirstLineIndent;

	private TextBox Txt_RightIndent;

	private TextBox Txt_LeftIndent;

	private Button Btn_BuildInStyleSelect;

	private ListBox Lst_Styles;

	private Button Btn_SaveStyles;

	private Button Btn_LoadStyles;

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
		userChanged = false;
		Cmb_PaperSize.SelectedIndex = 1;
		Cmb_PaperDirection.SelectedIndex = 0;
		Cmb_PageMargin.SelectedIndex = 0;
		Cmb_GutterPosition.SelectedIndex = 0;
		FontFamily[] families = new InstalledFontCollection().Families;
		foreach (FontFamily fontFamily in families)
		{
			Cmb_ChnFontName.Items.Add(fontFamily.Name);
			Cmb_EngFontName.Items.Add(fontFamily.Name);
		}
		ComboBox.ObjectCollection items = Cmb_FontSize.Items;
		List<string> fontSizeList = WordStyleInfo.FontSizeList;
		int i = 0;
		object[] array = new object[fontSizeList.Count];
		foreach (string item in fontSizeList)
		{
			array[i] = item;
			i++;
		}
		items.AddRange(array);
		Cmb_ParaAligment.Items.AddRange(((IEnumerable<object>)WordStyleInfo.HAlignments).ToArray());
		Cmb_LineSpace.Items.AddRange(((IEnumerable<object>)WordStyleInfo.LineSpacingValues).ToArray());
		Cmb_SpaceBefore.Items.AddRange(((IEnumerable<object>)WordStyleInfo.ParagraphSpaceValues).ToArray());
		Cmb_SpaceAfter.Items.AddRange(((IEnumerable<object>)WordStyleInfo.ParagraphSpaceValues).ToArray());
		items = Cmb_NumberStyle.Items;
		List<string> listNumberStyleName = WordStyleInfo.ListNumberStyleName;
		i = 0;
		array = new object[listNumberStyleName.Count];
		foreach (string item2 in listNumberStyleName)
		{
			array[i] = item2;
			i++;
		}
		items.AddRange(array);
		Cmb_SetLevel.SelectedIndex = 2;
		if (AllBuildInStyleNames.Count == 0)
		{
			foreach (WdBuiltinStyle buildInStyleName in WordStyleInfo.BuildInStyleNames)
			{
				List<string> allBuildInStyleNames = AllBuildInStyleNames;
				Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				object Index = buildInStyleName;
				allBuildInStyleNames.Add(styles[ref Index].NameLocal);
			}
		}
		UpdateStyleList();
		Lst_Styles.SelectedIndex = -1;
		userChanged = true;
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

	private void Btn_AddStyle_Click(object sender, EventArgs e)
	{
		if (CustomStyles.FindIndex((WordStyleInfo item) => item.StyleName == Txt_NewStyleName.Text) != -1 || AllBuildInStyleNames.Contains(Txt_NewStyleName.Text) || string.IsNullOrEmpty(Txt_NewStyleName.Text.Trim()))
		{
			MessageBox.Show("指定的自定义样式名称为空或样式名已经存在，请重新命名！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			Txt_NewStyleName.Text = "";
		}
		else
		{
			CustomStyles.Add(new WordStyleInfo(Txt_NewStyleName.Text, new WordStyleInfo.StyleParaValues()));
			UpdateStyleList();
			Txt_NewStyleName.Text = "";
		}
	}

	private void Lst_Styles_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (Lst_Styles.SelectedIndex != -1)
		{
			userChanged = false;
			SetUIValues(Styles[Lst_Styles.SelectedIndex]);
			Btn_DelStyle.Enabled = Lst_Styles.SelectedIndices.Count > 0;
		}
	}

	private void SetUIValues(WordStyleInfo style)
	{
		int selectedIndex = Cmb_ChnFontName.Items.IndexOf(style.ChnFontName);
		Cmb_ChnFontName.SelectedIndex = selectedIndex;
		selectedIndex = Cmb_EngFontName.Items.IndexOf(style.EngFontName);
		Cmb_EngFontName.SelectedIndex = selectedIndex;
		selectedIndex = WordStyleInfo.FontSizeList.IndexOf(style.FontSize);
		if (selectedIndex != -1)
		{
			Cmb_FontSize.SelectedIndex = -1;
			Cmb_FontSize.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_FontSize.Text = style.FontSize;
		}
		Chk_Bold.Checked = style.Bold;
		Chk_Italic.Checked = style.Italic;
		Chk_UnderLine.Checked = style.Underline;
		Btn_FontColor.BackColor = style.FontColor;
		selectedIndex = WordStyleInfo.HAlignments.ToList().IndexOf(style.HAlignment);
		Cmb_ParaAligment.SelectedIndex = selectedIndex;
		Txt_LeftIndent.Text = style.LeftIndent;
		Txt_RightIndent.Text = style.RightIndent;
		Txt_FirstLineIndent.Text = style.FirstLineIndent;
		selectedIndex = WordStyleInfo.LineSpacingValues.ToList().IndexOf(style.LineSpace);
		if (selectedIndex != -1)
		{
			Cmb_LineSpace.SelectedIndex = -1;
			Cmb_LineSpace.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_LineSpace.Text = style.LineSpace;
		}
		selectedIndex = WordStyleInfo.ParagraphSpaceValues.ToList().IndexOf(style.SpaceBefore);
		if (selectedIndex != -1)
		{
			Cmb_SpaceBefore.SelectedIndex = -1;
			Cmb_SpaceBefore.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_SpaceBefore.Text = style.SpaceBefore;
		}
		selectedIndex = WordStyleInfo.ParagraphSpaceValues.ToList().IndexOf(style.SpaceAfter);
		if (selectedIndex != -1)
		{
			Cmb_SpaceAfter.SelectedIndex = -1;
			Cmb_SpaceAfter.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_SpaceAfter.Text = style.SpaceAfter;
		}
		Chk_BeforeBreak.Checked = style.BreakBefore;
		Cmb_NumberStyle.SelectedIndex = style.NumberStyle;
		Txt_NumberFormat.Text = style.NumberFormat;
		Lab_StyleInfo.Text = style.GetDescription(out var font);
		Lab_StyleInfo.Font = font;
		userChanged = true;
	}

	private void UpdateStyleList()
	{
		List<WordStyleInfo> buildInStyles = BuildInStyles;
		List<WordStyleInfo> customStyles = CustomStyles;
		List<WordStyleInfo> list = new List<WordStyleInfo>(buildInStyles.Count + customStyles.Count);
		list.AddRange(buildInStyles);
		list.AddRange(customStyles);
		Styles = list;
		Lst_Styles.SelectedIndex = -1;
		Lst_Styles.DisplayMember = "StyleName";
		Lst_Styles.DataSource = Styles;
	}

	private void Btn_DelStyle_Click(object sender, EventArgs e)
	{
		if (Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			if (Styles[selectedIndex].BuildInStyle)
			{
				BuildInStyles.Remove(Styles[selectedIndex]);
			}
			else
			{
				CustomStyles.Remove(Styles[selectedIndex]);
			}
		}
		Lst_Styles.SelectedIndex = -1;
		UpdateStyleList();
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
		string text = "";
		foreach (WordStyleInfo style in Styles)
		{
			if (!style.SetStyle(document))
			{
				text = text + style.StyleName + ";";
			}
		}
		if (!string.IsNullOrEmpty(text))
		{
			MessageBox.Show("样式：" + text.TrimEnd(';') + " 设置失败！请检查样式参数是否正确！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
		if (Chk_CreateListLevels.Checked)
		{
			Globals.ThisAddIn.AutoCreateLevelList(Cmb_SetLevel.SelectedIndex + 1, numIndent, textIndent, afterIndent);
		}
		(base.Parent as Form).Close();
	}

	private void Cmb_NumberStyle_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		Txt_NumberFormat.Enabled = Cmb_NumberStyle.SelectedIndex != 0;
		if (Cmb_NumberStyle.SelectedIndex != 0 && Txt_NumberFormat.Text == "")
		{
			Txt_NumberFormat.Text = "%1";
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			Styles[selectedIndex].NumberStyle = Cmb_NumberStyle.SelectedIndex;
			Styles[selectedIndex].NumberFormat = Txt_NumberFormat.Text;
		}
	}

	private void Txt_NumberFormat_TextChanged(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		if (!Regex.IsMatch(Txt_NumberFormat.Text, ".*%1.*"))
		{
			MessageBox.Show("格式必须包含%1标题编号!", "提醒");
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			Styles[selectedIndex].NumberFormat = Txt_NumberFormat.Text;
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

	private void Chk_CreateListLevels_CheckedChanged(object sender, EventArgs e)
	{
		Cmb_SetLevel.Enabled = Chk_CreateListLevels.Checked;
	}

	private void FontNameAndAlignment_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		ComboBox comboBox = sender as ComboBox;
		if (Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			switch (comboBox.Name)
			{
			case "Cmb_ChnFontName":
				Styles[selectedIndex].ChnFontName = comboBox.Text;
				break;
			case "Cmb_EngFontName":
				Styles[selectedIndex].EngFontName = comboBox.Text;
				break;
			case "Cmb_ParaAligment":
				Styles[selectedIndex].HAlignment = comboBox.Text;
				break;
			}
		}
	}

	private void FontStyleAndBreakPage_CheckedChanged(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		CheckBox checkBox = sender as CheckBox;
		if (Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			switch (checkBox.Name)
			{
			case "Chk_Bold":
				Styles[selectedIndex].Bold = checkBox.Checked;
				break;
			case "Chk_Italic":
				Styles[selectedIndex].Italic = checkBox.Checked;
				break;
			case "Chk_UnderLine":
				Styles[selectedIndex].Underline = checkBox.Checked;
				break;
			case "Chk_BeforeBreak":
				Styles[selectedIndex].BreakBefore = checkBox.Checked;
				break;
			}
		}
	}

	private void Btn_FontColor_BackColorChanged(object sender, EventArgs e)
	{
		if (!userChanged || Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			Styles[selectedIndex].FontColor = Btn_FontColor.BackColor;
		}
	}

	private void Cmb_FontSize_Validated(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		int num = WordStyleInfo.FontSizeList.IndexOf(Cmb_FontSize.Text);
		if (num == -1)
		{
			if (Regex.IsMatch(Cmb_FontSize.Text, "^\\d+(?:\\.(?:0|5))?(?:\\s+)?磅?$"))
			{
				Cmb_FontSize.Text = Cmb_FontSize.Text.TrimEnd(' ', '磅') + " 磅";
			}
			else
			{
				Cmb_FontSize.SelectedIndex = -1;
				Cmb_FontSize.SelectedIndex = 10;
			}
		}
		else
		{
			Cmb_FontSize.SelectedIndex = -1;
			Cmb_FontSize.SelectedIndex = num;
		}
	}

	private void Cmb_FontSize_TextChanged(object sender, EventArgs e)
	{
		if (!userChanged || Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			Styles[selectedIndex].FontSize = Cmb_FontSize.Text;
		}
	}

	private void IndentSetting_TextChanged(object sender, EventArgs e)
	{
		if (!userChanged || Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			string name = (sender as TextBox).Name;
			if (!(name == "Txt_LeftIndent"))
			{
				if (name == "Txt_RightIndent")
				{
					Styles[selectedIndex].RightIndent = (sender as TextBox).Text;
				}
			}
			else
			{
				Styles[selectedIndex].LeftIndent = (sender as TextBox).Text;
			}
		}
	}

	private void IndentChanged_Validated(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		TextBox textBox = sender as TextBox;
		string s = textBox.Text.TrimEnd(' ', '磅', '厘', '米');
		try
		{
			float num = float.Parse(s);
			if (textBox.Text.EndsWith("厘米"))
			{
				textBox.Text = num.ToString("0.00 厘米");
			}
			else
			{
				textBox.Text = num.ToString("0.00 磅");
			}
		}
		catch
		{
			textBox.Text = "0.00 厘米";
		}
	}

	private void Txt_FirstLineIndent_Validated(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		TextBox textBox = sender as TextBox;
		string s = textBox.Text.TrimEnd(' ', '磅', '字', '符');
		try
		{
			float num = float.Parse(s);
			if (textBox.Text.EndsWith("字符"))
			{
				textBox.Text = num.ToString("0 字符");
			}
			else
			{
				textBox.Text = num.ToString("0.00 磅");
			}
		}
		catch
		{
			textBox.Text = "0 字符";
		}
	}

	private void Txt_FirstLineIndent_TextChanged(object sender, EventArgs e)
	{
		if (!userChanged || Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			Styles[selectedIndex].FirstLineIndent = (sender as TextBox).Text;
		}
	}

	private void Cmb_LineSpace_TextChanged(object sender, EventArgs e)
	{
		if (!userChanged || Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			Styles[selectedIndex].LineSpace = (sender as ComboBox).Text;
		}
	}

	private void Cmb_LineSpace_Validated(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		int num = WordStyleInfo.LineSpacingValues.ToList().IndexOf(Cmb_LineSpace.Text.Trim());
		if (num == -1)
		{
			string s = Cmb_LineSpace.Text.TrimEnd(' ', '行', '磅');
			try
			{
				float num2 = float.Parse(s);
				if (Cmb_LineSpace.Text.EndsWith("行"))
				{
					Cmb_LineSpace.Text = num2.ToString("0.00 行");
				}
				else
				{
					Cmb_LineSpace.Text = num2.ToString("0.00 磅");
				}
				return;
			}
			catch
			{
				Cmb_LineSpace.SelectedIndex = -1;
				Cmb_LineSpace.SelectedIndex = 1;
				return;
			}
		}
		Cmb_LineSpace.SelectedIndex = -1;
		Cmb_LineSpace.SelectedIndex = num;
	}

	private void ParagraphSpace_TextChanged(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		ComboBox comboBox = sender as ComboBox;
		if (Lst_Styles.SelectedIndices.Count <= 0)
		{
			return;
		}
		foreach (int selectedIndex in Lst_Styles.SelectedIndices)
		{
			string name = comboBox.Name;
			if (!(name == "Cmb_SpaceBefore"))
			{
				if (name == "Cmb_SpaceAfter")
				{
					Styles[selectedIndex].SpaceAfter = comboBox.Text;
				}
			}
			else
			{
				Styles[selectedIndex].SpaceBefore = comboBox.Text;
			}
		}
	}

	private void ParagraphSpace_Validated(object sender, EventArgs e)
	{
		if (!userChanged)
		{
			return;
		}
		ComboBox comboBox = sender as ComboBox;
		int num = WordStyleInfo.ParagraphSpaceValues.ToList().IndexOf(comboBox.Text.Trim());
		if (num == -1)
		{
			string s = comboBox.Text.TrimEnd(' ', '行', '磅');
			try
			{
				float num2 = float.Parse(s);
				if (comboBox.Text.EndsWith("行"))
				{
					comboBox.Text = num2.ToString("0.00 行");
				}
				else
				{
					comboBox.Text = num2.ToString("0.00 磅");
				}
				return;
			}
			catch
			{
				comboBox.SelectedIndex = -1;
				comboBox.SelectedIndex = 1;
				return;
			}
		}
		comboBox.SelectedIndex = -1;
		comboBox.SelectedIndex = num;
	}

	private void Btn_FontColor_Click(object sender, EventArgs e)
	{
		ColorDialog colorDialog = new ColorDialog
		{
			Color = Btn_FontColor.BackColor,
			AnyColor = true,
			AllowFullOpen = true,
			FullOpen = true,
			SolidColorOnly = true
		};
		if (colorDialog.ShowDialog(base.Parent as Form) == DialogResult.OK)
		{
			Btn_FontColor.BackColor = colorDialog.Color;
		}
	}

	private void Btn_BuildInStyleSelect_Click(object sender, EventArgs e)
	{
		List<int> list = new List<int>();
		for (int i = 0; i < BuildInStyles.Count; i++)
		{
			int num = AllBuildInStyleNames.IndexOf(BuildInStyles[i].StyleName);
			if (num != -1)
			{
				list.Add(num);
			}
		}
		BuildInStyleSeletor buildInStyleSeletor = new BuildInStyleSeletor(AllBuildInStyleNames.ToArray(), list.ToArray());
		if (buildInStyleSeletor.ShowDialog(base.Parent as Form) == DialogResult.OK)
		{
			List<WordStyleInfo> list2 = new List<WordStyleInfo>();
			int[] selectedIndices = buildInStyleSeletor.SelectedIndices;
			foreach (int index in selectedIndices)
			{
				WdBuiltinStyle name = WordStyleInfo.BuildInStyleNames[index];
				WordStyleInfo wordStyleInfo = BuildInStyles.Find((WordStyleInfo item) => item.BuildInStyleName == name);
				if (wordStyleInfo == null)
				{
					Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
					object Index = name;
					list2.Add(new WordStyleInfo(styles[ref Index], name));
				}
				else
				{
					list2.Add(wordStyleInfo);
				}
			}
			BuildInStyles.Clear();
			BuildInStyles.AddRange(list2);
		}
		UpdateStyleList();
	}

	private void Cmb_PreSettings_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (Styles.Count <= 0 || MessageBox.Show("当前存在样式，选择预定义样式会覆盖当前所有设置，是否继续？", "Word格式助手", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk) != DialogResult.Cancel)
		{
			switch (Cmb_PreSettings.SelectedIndex)
			{
			case 0:
			{
				Cmb_PaperSize.SelectedIndex = 1;
				Cmb_PaperDirection.SelectedIndex = 0;
				Cmb_PageMargin.SelectedIndex = 5;
				Chk_SetGutter.Checked = false;
				Cmb_GutterPosition.SelectedIndex = 0;
				Nud_GutterValue.Value = 0m;
				Chk_CreateListLevels.Checked = false;
				BuildInStyles.Clear();
				WordStyleInfo.StyleParaValues styleParaValues = new WordStyleInfo.StyleParaValues();
				styleParaValues.ChnFontName = "仿宋";
				styleParaValues.EngFontName = "仿宋";
				styleParaValues.FontSize = "三号";
				styleParaValues.FirstLineIndent = "2 字符";
				styleParaValues.LineSpace = "1.25 行";
				styleParaValues.HAlignment = "两端对齐";
				WordStyleInfo.StyleParaValues para = styleParaValues;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleNormal, para));
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.NumberStyle = 7;
				para.NumberFormat = "%1、";
				para.HAlignment = "左对齐";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading1, para));
				para.ChnFontName = "楷体";
				para.EngFontName = "楷体";
				para.NumberFormat = "(%1)";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading2, para));
				para.ChnFontName = "仿宋";
				para.EngFontName = "仿宋";
				para.Bold = true;
				para.NumberStyle = 1;
				para.NumberFormat = "%1.";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading3, para));
				para.Bold = false;
				para.NumberFormat = "(%1)";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading4, para));
				para.ChnFontName = "华文中宋";
				para.EngFontName = "华文中宋";
				para.FontSize = "二号";
				para.HAlignment = "中对齐";
				para.FirstLineIndent = "0.00 磅";
				para.NumberStyle = -1;
				para.NumberFormat = "";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleTitle, para));
				para.FontSize = "三号";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleSubtitle, para));
				CustomStyles.Clear();
				para.ChnFontName = "仿宋";
				para.EngFontName = "仿宋";
				para.Bold = true;
				para.HAlignment = "左对齐";
				CustomStyles.Add(new WordStyleInfo("附录标题", para));
				para.FontSize = "小四";
				para.HAlignment = "中对齐";
				CustomStyles.Add(new WordStyleInfo("表格标题", para));
				CustomStyles.Add(new WordStyleInfo("插图标题", para));
				para.FontSize = "五号";
				para.Bold = false;
				para.LineSpace = "单倍行距";
				para.HAlignment = "左对齐";
				CustomStyles.Add(new WordStyleInfo("表内文字", para));
				break;
			}
			case 1:
			{
				Cmb_PaperSize.SelectedIndex = 1;
				Cmb_PaperDirection.SelectedIndex = 0;
				Cmb_PageMargin.SelectedIndex = 2;
				Chk_SetGutter.Checked = true;
				Cmb_GutterPosition.SelectedIndex = 2;
				Nud_GutterValue.Value = 0.5m;
				Chk_CreateListLevels.Checked = false;
				BuildInStyles.Clear();
				WordStyleInfo.StyleParaValues styleParaValues = new WordStyleInfo.StyleParaValues();
				styleParaValues.ChnFontName = "宋体";
				styleParaValues.EngFontName = "宋体";
				styleParaValues.FontSize = "小四";
				styleParaValues.FirstLineIndent = "2 字符";
				styleParaValues.LineSpace = "1.3 行";
				styleParaValues.HAlignment = "两端对齐";
				WordStyleInfo.StyleParaValues para = styleParaValues;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleNormal, para));
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.FontSize = "四号";
				para.Bold = true;
				para.SpaceBefore = "0.50 行";
				para.HAlignment = "左对齐";
				para.FirstLineIndent = "0.00 磅";
				para.NumberStyle = 7;
				para.NumberFormat = "%1、";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading1, para));
				para.ChnFontName = "宋体";
				para.EngFontName = "宋体";
				para.FontSize = "小四";
				para.Bold = false;
				para.NumberStyle = 7;
				para.NumberFormat = "(%1)";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading2, para));
				para.FirstLineIndent = "2 字符";
				para.NumberStyle = 1;
				para.NumberFormat = "%1.";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading3, para));
				para.NumberFormat = "(%1)";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading4, para));
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.FontSize = "二号";
				para.Bold = true;
				para.SpaceBefore = "0.00 行";
				para.HAlignment = "中对齐";
				para.FirstLineIndent = "0.00 磅";
				para.NumberStyle = -1;
				para.NumberFormat = "";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleTitle, para));
				para.FontSize = "三号";
				para.Bold = false;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleSubtitle, para));
				CustomStyles.Clear();
				para.ChnFontName = "宋体";
				para.EngFontName = "宋体";
				para.FontSize = "小四";
				para.Bold = true;
				para.SpaceBefore = "0.00 行";
				para.HAlignment = "左对齐";
				para.NumberStyle = -1;
				para.NumberFormat = "";
				CustomStyles.Add(new WordStyleInfo("附录标题", para));
				para.HAlignment = "中对齐";
				CustomStyles.Add(new WordStyleInfo("表格标题", para));
				CustomStyles.Add(new WordStyleInfo("插图标题", para));
				para.FontSize = "五号";
				para.Bold = false;
				para.LineSpace = "单倍行距";
				para.HAlignment = "左对齐";
				CustomStyles.Add(new WordStyleInfo("表内文字", para));
				break;
			}
			case 2:
			{
				Cmb_PaperSize.SelectedIndex = 1;
				Cmb_PaperDirection.SelectedIndex = 0;
				Cmb_PageMargin.SelectedIndex = 2;
				Chk_SetGutter.Checked = true;
				Cmb_GutterPosition.SelectedIndex = 2;
				Nud_GutterValue.Value = 0.5m;
				Chk_CreateListLevels.Checked = true;
				BuildInStyles.Clear();
				WordStyleInfo.StyleParaValues styleParaValues = new WordStyleInfo.StyleParaValues();
				styleParaValues.ChnFontName = "宋体";
				styleParaValues.EngFontName = "宋体";
				styleParaValues.FontSize = "小四";
				styleParaValues.FirstLineIndent = "2 字符";
				styleParaValues.LineSpace = "1.3 行";
				styleParaValues.HAlignment = "两端对齐";
				WordStyleInfo.StyleParaValues para = styleParaValues;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleNormal, para));
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.FontSize = "三号";
				para.Bold = true;
				para.HAlignment = "左对齐";
				para.FirstLineIndent = "0.00 磅";
				para.BreakBefore = true;
				para.SpaceBefore = "0.50 行";
				para.NumberStyle = 0;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading1, para));
				para.Bold = false;
				para.BreakBefore = false;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading2, para));
				para.FontSize = "小三";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading3, para));
				para.ChnFontName = "宋体";
				para.EngFontName = "宋体";
				para.FontSize = "小四";
				para.FirstLineIndent = "2 字符";
				para.SpaceBefore = "0.00 行";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading4, para));
				para.FontSize = "小一";
				para.Bold = true;
				para.HAlignment = "中对齐";
				para.FirstLineIndent = "0.00 磅";
				para.NumberStyle = -1;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleTitle, para));
				para.FontSize = "三号";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleSubtitle, para));
				CustomStyles.Clear();
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.FontSize = "小三";
				para.Bold = true;
				para.HAlignment = "左对齐";
				CustomStyles.Add(new WordStyleInfo("附录标题", para));
				para.ChnFontName = "宋体";
				para.EngFontName = "宋体";
				para.FontSize = "小四";
				para.HAlignment = "中对齐";
				CustomStyles.Add(new WordStyleInfo("表格标题", para));
				CustomStyles.Add(new WordStyleInfo("插图标题", para));
				para.FontSize = "五号";
				para.Bold = false;
				para.HAlignment = "左对齐";
				para.LineSpace = "单倍行距";
				CustomStyles.Add(new WordStyleInfo("表内文字", para));
				break;
			}
			case 3:
			{
				Cmb_PaperSize.SelectedIndex = 1;
				Cmb_PaperDirection.SelectedIndex = 0;
				Cmb_SetLevel.SelectedIndex = 3;
				Cmb_PageMargin.SelectedIndex = 2;
				Chk_SetGutter.Checked = true;
				Cmb_GutterPosition.SelectedIndex = 2;
				Nud_GutterValue.Value = 0.5m;
				Chk_CreateListLevels.Checked = true;
				BuildInStyles.Clear();
				WordStyleInfo.StyleParaValues styleParaValues = new WordStyleInfo.StyleParaValues();
				styleParaValues.ChnFontName = "宋体";
				styleParaValues.EngFontName = "宋体";
				styleParaValues.FontSize = "小四";
				styleParaValues.LineSpace = "1.35 行";
				styleParaValues.HAlignment = "两端对齐";
				styleParaValues.LeftIndent = "2.20 厘米";
				WordStyleInfo.StyleParaValues para = styleParaValues;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleNormal, para));
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.FontSize = "四号";
				para.Bold = true;
				para.LeftIndent = "0.00 厘米";
				para.SpaceBefore = "0.50 行";
				para.NumberStyle = 0;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading1, para));
				para.Bold = false;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading2, para));
				para.ChnFontName = "宋体";
				para.EngFontName = "宋体";
				para.FontSize = "小四";
				para.Bold = true;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading3, para));
				para.Bold = false;
				para.SpaceBefore = "0.00 行";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleHeading4, para));
				para.FontSize = "二号";
				para.Bold = true;
				para.HAlignment = "中对齐";
				para.NumberStyle = -1;
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleTitle, para));
				para.FontSize = "三号";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleSubtitle, para));
				para.FontSize = "小四";
				para.Bold = false;
				para.LeftIndent = "2.20 厘米";
				para.NumberStyle = 1;
				para.NumberFormat = "(%1)";
				BuildInStyles.Add(new WordStyleInfo(WdBuiltinStyle.wdStyleList, para));
				CustomStyles.Clear();
				para.ChnFontName = "黑体";
				para.EngFontName = "黑体";
				para.FontSize = "四号";
				para.Bold = false;
				para.HAlignment = "左对齐";
				para.LeftIndent = "0.00 厘米";
				para.NumberStyle = -1;
				para.NumberFormat = "";
				CustomStyles.Add(new WordStyleInfo("附录标题", para));
				para.FontSize = "小四";
				para.Bold = true;
				para.HAlignment = "中对齐";
				para.SpaceBefore = "0.50 行";
				CustomStyles.Add(new WordStyleInfo("表格标题", para));
				para.SpaceBefore = "0.00 行";
				para.SpaceAfter = "0.50 行";
				CustomStyles.Add(new WordStyleInfo("插图标题", para));
				para.ChnFontName = "宋体";
				para.EngFontName = "宋体";
				para.FontSize = "五号";
				para.Bold = false;
				para.HAlignment = "左对齐";
				para.LineSpace = "单倍行距";
				CustomStyles.Add(new WordStyleInfo("表内文字", para));
				numIndent = 0f;
				textIndent = 2.2f;
				afterIndent = 2.2f;
				break;
			}
			}
			UpdateStyleList();
			Lst_Styles.SelectedItems.Clear();
			Lst_Styles.SelectedIndex = 0;
		}
	}

	private void Btn_SaveStyles_Click(object sender, EventArgs e)
	{
		if (Styles.Count > 0)
		{
			SaveFileDialog saveFileDialog = new SaveFileDialog
			{
				Filter = "样式文件|*.xml",
				Title = "Word格式助手",
				DefaultExt = "xml",
				AddExtension = true,
				SupportMultiDottedExtensions = true
			};
			if (saveFileDialog.ShowDialog(base.Parent as Form) == DialogResult.OK)
			{
				StyleSerializationHelper.SerializeListToXml(Styles, saveFileDialog.FileName);
			}
		}
	}

	private void Btn_LoadStyles_Click(object sender, EventArgs e)
	{
		if (Styles.Count > 0 && MessageBox.Show("当前存在样式，导入样式会覆盖当前样式，是否继续？", "Word格式助手", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
		{
			return;
		}
		OpenFileDialog openFileDialog = new OpenFileDialog
		{
			Multiselect = false,
			DefaultExt = "xml",
			Filter = "样式文件|*.xml",
			CheckFileExists = true,
			Title = "Word格式助手",
			SupportMultiDottedExtensions = true
		};
		if (openFileDialog.ShowDialog(base.Parent as Form) != DialogResult.OK)
		{
			return;
		}
		List<WordStyleInfo> list = StyleSerializationHelper.DeserializeListFromXml<WordStyleInfo>(openFileDialog.FileName);
		BuildInStyles.Clear();
		CustomStyles.Clear();
		foreach (WordStyleInfo item in list)
		{
			if (item.BuildInStyle)
			{
				BuildInStyles.Add(item);
			}
			else
			{
				CustomStyles.Add(item);
			}
		}
		UpdateStyleList();
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
		this.Chk_SetPage = new System.Windows.Forms.CheckBox();
		this.Rdo_NewDocument = new System.Windows.Forms.RadioButton();
		this.Rdo_UseCurrentDocument = new System.Windows.Forms.RadioButton();
		this.label1 = new System.Windows.Forms.Label();
		this.Cmb_PaperSize = new System.Windows.Forms.ComboBox();
		this.Grp_PageSetup = new System.Windows.Forms.GroupBox();
		this.Nud_GutterValue = new WordFormatHelper.NumericUpDownWithUnit();
		this.Cmb_GutterPosition = new System.Windows.Forms.ComboBox();
		this.Chk_SetGutter = new System.Windows.Forms.CheckBox();
		this.Cmb_PageMargin = new System.Windows.Forms.ComboBox();
		this.label5 = new System.Windows.Forms.Label();
		this.Cmb_PaperDirection = new System.Windows.Forms.ComboBox();
		this.label2 = new System.Windows.Forms.Label();
		this.groupBox3 = new System.Windows.Forms.GroupBox();
		this.Lst_Styles = new System.Windows.Forms.ListBox();
		this.Btn_BuildInStyleSelect = new System.Windows.Forms.Button();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.Cmb_SpaceAfter = new System.Windows.Forms.ComboBox();
		this.Cmb_SpaceBefore = new System.Windows.Forms.ComboBox();
		this.Cmb_LineSpace = new System.Windows.Forms.ComboBox();
		this.Txt_FirstLineIndent = new System.Windows.Forms.TextBox();
		this.Txt_RightIndent = new System.Windows.Forms.TextBox();
		this.Txt_LeftIndent = new System.Windows.Forms.TextBox();
		this.label18 = new System.Windows.Forms.Label();
		this.Txt_NumberFormat = new System.Windows.Forms.TextBox();
		this.label14 = new System.Windows.Forms.Label();
		this.Cmb_NumberStyle = new System.Windows.Forms.ComboBox();
		this.label13 = new System.Windows.Forms.Label();
		this.label11 = new System.Windows.Forms.Label();
		this.Chk_BeforeBreak = new System.Windows.Forms.CheckBox();
		this.label10 = new System.Windows.Forms.Label();
		this.label9 = new System.Windows.Forms.Label();
		this.label7 = new System.Windows.Forms.Label();
		this.label6 = new System.Windows.Forms.Label();
		this.Cmb_ParaAligment = new System.Windows.Forms.ComboBox();
		this.label15 = new System.Windows.Forms.Label();
		this.Btn_FontColor = new System.Windows.Forms.Button();
		this.Chk_UnderLine = new System.Windows.Forms.CheckBox();
		this.Chk_Italic = new System.Windows.Forms.CheckBox();
		this.Chk_Bold = new System.Windows.Forms.CheckBox();
		this.Cmb_FontSize = new System.Windows.Forms.ComboBox();
		this.label4 = new System.Windows.Forms.Label();
		this.Cmb_EngFontName = new System.Windows.Forms.ComboBox();
		this.label17 = new System.Windows.Forms.Label();
		this.Cmb_ChnFontName = new System.Windows.Forms.ComboBox();
		this.label3 = new System.Windows.Forms.Label();
		this.Txt_NewStyleName = new WordFormatHelper.WaterMarkTextBoxEx();
		this.Lab_StyleInfo = new System.Windows.Forms.Label();
		this.Btn_DelStyle = new System.Windows.Forms.Button();
		this.Btn_AddStyle = new System.Windows.Forms.Button();
		this.Chk_CreateListLevels = new System.Windows.Forms.CheckBox();
		this.Cmb_SetLevel = new System.Windows.Forms.ComboBox();
		this.label12 = new System.Windows.Forms.Label();
		this.Btn_ApplySet = new System.Windows.Forms.Button();
		this.Cmb_PreSettings = new System.Windows.Forms.ComboBox();
		this.label16 = new System.Windows.Forms.Label();
		this.Btn_SaveStyles = new System.Windows.Forms.Button();
		this.Btn_LoadStyles = new System.Windows.Forms.Button();
		this.groupBox1.SuspendLayout();
		this.Grp_PageSetup.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_GutterValue).BeginInit();
		this.groupBox3.SuspendLayout();
		this.groupBox2.SuspendLayout();
		base.SuspendLayout();
		this.groupBox1.Controls.Add(this.Chk_SetPage);
		this.groupBox1.Controls.Add(this.Rdo_NewDocument);
		this.groupBox1.Controls.Add(this.Rdo_UseCurrentDocument);
		this.groupBox1.Location = new System.Drawing.Point(3, 3);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(265, 129);
		this.groupBox1.TabIndex = 0;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "文件";
		this.Chk_SetPage.AutoSize = true;
		this.Chk_SetPage.Location = new System.Drawing.Point(12, 96);
		this.Chk_SetPage.Name = "Chk_SetPage";
		this.Chk_SetPage.Size = new System.Drawing.Size(84, 24);
		this.Chk_SetPage.TabIndex = 2;
		this.Chk_SetPage.Text = "设置页面";
		this.Chk_SetPage.UseVisualStyleBackColor = true;
		this.Chk_SetPage.CheckedChanged += new System.EventHandler(Chk_SetPage_CheckedChanged);
		this.Rdo_NewDocument.AutoSize = true;
		this.Rdo_NewDocument.Location = new System.Drawing.Point(12, 59);
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
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(18, 26);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(65, 20);
		this.label1.TabIndex = 3;
		this.label1.Text = "页面大小";
		this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_PaperSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PaperSize.FormattingEnabled = true;
		this.Cmb_PaperSize.Items.AddRange(new object[21]
		{
			"A3", "A4", "Small A4", "A5", "B4", "B5", "10#信封", "B4信封", "B5信封", "B6信封",
			"C3信封", "C4信封", "C5信封", "C6信封", "对开本", "账簿", "法律纸张", "信纸", "小信纸", "便签",
			"4开本"
		});
		this.Cmb_PaperSize.Location = new System.Drawing.Point(96, 22);
		this.Cmb_PaperSize.Name = "Cmb_PaperSize";
		this.Cmb_PaperSize.Size = new System.Drawing.Size(80, 28);
		this.Cmb_PaperSize.TabIndex = 4;
		this.Grp_PageSetup.Controls.Add(this.Nud_GutterValue);
		this.Grp_PageSetup.Controls.Add(this.Cmb_GutterPosition);
		this.Grp_PageSetup.Controls.Add(this.Chk_SetGutter);
		this.Grp_PageSetup.Controls.Add(this.Cmb_PageMargin);
		this.Grp_PageSetup.Controls.Add(this.label5);
		this.Grp_PageSetup.Controls.Add(this.Cmb_PaperDirection);
		this.Grp_PageSetup.Controls.Add(this.label2);
		this.Grp_PageSetup.Controls.Add(this.Cmb_PaperSize);
		this.Grp_PageSetup.Controls.Add(this.label1);
		this.Grp_PageSetup.Enabled = false;
		this.Grp_PageSetup.Location = new System.Drawing.Point(274, 3);
		this.Grp_PageSetup.Name = "Grp_PageSetup";
		this.Grp_PageSetup.Size = new System.Drawing.Size(323, 129);
		this.Grp_PageSetup.TabIndex = 5;
		this.Grp_PageSetup.TabStop = false;
		this.Grp_PageSetup.Text = "页面";
		this.Nud_GutterValue.DecimalPlaces = 2;
		this.Nud_GutterValue.Enabled = false;
		this.Nud_GutterValue.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_GutterValue.Label = "厘米";
		this.Nud_GutterValue.Location = new System.Drawing.Point(202, 95);
		this.Nud_GutterValue.Name = "Nud_GutterValue";
		this.Nud_GutterValue.Size = new System.Drawing.Size(112, 26);
		this.Nud_GutterValue.TabIndex = 11;
		this.Cmb_GutterPosition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_GutterPosition.Enabled = false;
		this.Cmb_GutterPosition.FormattingEnabled = true;
		this.Cmb_GutterPosition.Items.AddRange(new object[3] { "左", "上", "对称" });
		this.Cmb_GutterPosition.Location = new System.Drawing.Point(96, 94);
		this.Cmb_GutterPosition.Name = "Cmb_GutterPosition";
		this.Cmb_GutterPosition.Size = new System.Drawing.Size(100, 28);
		this.Cmb_GutterPosition.TabIndex = 10;
		this.Chk_SetGutter.AutoSize = true;
		this.Chk_SetGutter.Location = new System.Drawing.Point(22, 96);
		this.Chk_SetGutter.Name = "Chk_SetGutter";
		this.Chk_SetGutter.Size = new System.Drawing.Size(70, 24);
		this.Chk_SetGutter.TabIndex = 9;
		this.Chk_SetGutter.Text = "装订线";
		this.Chk_SetGutter.UseVisualStyleBackColor = true;
		this.Chk_SetGutter.CheckedChanged += new System.EventHandler(Chk_SetGutter_CheckedChanged);
		this.Cmb_PageMargin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PageMargin.FormattingEnabled = true;
		this.Cmb_PageMargin.Items.AddRange(new object[6] { "使用默认值", "全部设置2.0厘米", "全部设置2.5厘米", "全部设置3.0厘米", "全部设置3.5厘米", "上3.7厘米;左2.8厘米" });
		this.Cmb_PageMargin.Location = new System.Drawing.Point(96, 57);
		this.Cmb_PageMargin.Name = "Cmb_PageMargin";
		this.Cmb_PageMargin.Size = new System.Drawing.Size(218, 28);
		this.Cmb_PageMargin.TabIndex = 8;
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(18, 61);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(51, 20);
		this.label5.TabIndex = 7;
		this.label5.Text = "页边距";
		this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_PaperDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PaperDirection.FormattingEnabled = true;
		this.Cmb_PaperDirection.Items.AddRange(new object[2] { "竖向", "横向" });
		this.Cmb_PaperDirection.Location = new System.Drawing.Point(234, 22);
		this.Cmb_PaperDirection.Name = "Cmb_PaperDirection";
		this.Cmb_PaperDirection.Size = new System.Drawing.Size(80, 28);
		this.Cmb_PaperDirection.TabIndex = 6;
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(191, 26);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(37, 20);
		this.label2.TabIndex = 5;
		this.label2.Text = "方向";
		this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.groupBox3.Controls.Add(this.Lst_Styles);
		this.groupBox3.Controls.Add(this.Btn_BuildInStyleSelect);
		this.groupBox3.Controls.Add(this.groupBox2);
		this.groupBox3.Controls.Add(this.Txt_NewStyleName);
		this.groupBox3.Controls.Add(this.Lab_StyleInfo);
		this.groupBox3.Controls.Add(this.Btn_DelStyle);
		this.groupBox3.Controls.Add(this.Btn_AddStyle);
		this.groupBox3.Location = new System.Drawing.Point(3, 138);
		this.groupBox3.Name = "groupBox3";
		this.groupBox3.Size = new System.Drawing.Size(594, 363);
		this.groupBox3.TabIndex = 6;
		this.groupBox3.TabStop = false;
		this.groupBox3.Text = "样式设置";
		this.Lst_Styles.FormattingEnabled = true;
		this.Lst_Styles.ItemHeight = 20;
		this.Lst_Styles.Location = new System.Drawing.Point(12, 25);
		this.Lst_Styles.Name = "Lst_Styles";
		this.Lst_Styles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
		this.Lst_Styles.Size = new System.Drawing.Size(168, 244);
		this.Lst_Styles.TabIndex = 35;
		this.Lst_Styles.SelectedIndexChanged += new System.EventHandler(Lst_Styles_SelectedIndexChanged);
		this.Btn_BuildInStyleSelect.FlatAppearance.BorderSize = 0;
		this.Btn_BuildInStyleSelect.FlatAppearance.MouseOverBackColor = System.Drawing.Color.RoyalBlue;
		this.Btn_BuildInStyleSelect.Location = new System.Drawing.Point(10, 269);
		this.Btn_BuildInStyleSelect.Name = "Btn_BuildInStyleSelect";
		this.Btn_BuildInStyleSelect.Size = new System.Drawing.Size(170, 28);
		this.Btn_BuildInStyleSelect.TabIndex = 29;
		this.Btn_BuildInStyleSelect.Text = "选择内置样式";
		this.Btn_BuildInStyleSelect.UseVisualStyleBackColor = false;
		this.Btn_BuildInStyleSelect.Click += new System.EventHandler(Btn_BuildInStyleSelect_Click);
		this.groupBox2.Controls.Add(this.Cmb_SpaceAfter);
		this.groupBox2.Controls.Add(this.Cmb_SpaceBefore);
		this.groupBox2.Controls.Add(this.Cmb_LineSpace);
		this.groupBox2.Controls.Add(this.Txt_FirstLineIndent);
		this.groupBox2.Controls.Add(this.Txt_RightIndent);
		this.groupBox2.Controls.Add(this.Txt_LeftIndent);
		this.groupBox2.Controls.Add(this.label18);
		this.groupBox2.Controls.Add(this.Txt_NumberFormat);
		this.groupBox2.Controls.Add(this.label14);
		this.groupBox2.Controls.Add(this.Cmb_NumberStyle);
		this.groupBox2.Controls.Add(this.label13);
		this.groupBox2.Controls.Add(this.label11);
		this.groupBox2.Controls.Add(this.Chk_BeforeBreak);
		this.groupBox2.Controls.Add(this.label10);
		this.groupBox2.Controls.Add(this.label9);
		this.groupBox2.Controls.Add(this.label7);
		this.groupBox2.Controls.Add(this.label6);
		this.groupBox2.Controls.Add(this.Cmb_ParaAligment);
		this.groupBox2.Controls.Add(this.label15);
		this.groupBox2.Controls.Add(this.Btn_FontColor);
		this.groupBox2.Controls.Add(this.Chk_UnderLine);
		this.groupBox2.Controls.Add(this.Chk_Italic);
		this.groupBox2.Controls.Add(this.Chk_Bold);
		this.groupBox2.Controls.Add(this.Cmb_FontSize);
		this.groupBox2.Controls.Add(this.label4);
		this.groupBox2.Controls.Add(this.Cmb_EngFontName);
		this.groupBox2.Controls.Add(this.label17);
		this.groupBox2.Controls.Add(this.Cmb_ChnFontName);
		this.groupBox2.Controls.Add(this.label3);
		this.groupBox2.Location = new System.Drawing.Point(189, 22);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.Size = new System.Drawing.Size(399, 275);
		this.groupBox2.TabIndex = 34;
		this.groupBox2.TabStop = false;
		this.groupBox2.Text = "样式设置";
		this.Cmb_SpaceAfter.FormattingEnabled = true;
		this.Cmb_SpaceAfter.Location = new System.Drawing.Point(269, 202);
		this.Cmb_SpaceAfter.Name = "Cmb_SpaceAfter";
		this.Cmb_SpaceAfter.Size = new System.Drawing.Size(111, 28);
		this.Cmb_SpaceAfter.TabIndex = 28;
		this.Cmb_SpaceAfter.TextChanged += new System.EventHandler(ParagraphSpace_TextChanged);
		this.Cmb_SpaceAfter.Validated += new System.EventHandler(ParagraphSpace_Validated);
		this.Cmb_SpaceBefore.FormattingEnabled = true;
		this.Cmb_SpaceBefore.Location = new System.Drawing.Point(79, 202);
		this.Cmb_SpaceBefore.Name = "Cmb_SpaceBefore";
		this.Cmb_SpaceBefore.Size = new System.Drawing.Size(111, 28);
		this.Cmb_SpaceBefore.TabIndex = 27;
		this.Cmb_SpaceBefore.TextChanged += new System.EventHandler(ParagraphSpace_TextChanged);
		this.Cmb_SpaceBefore.Validated += new System.EventHandler(ParagraphSpace_Validated);
		this.Cmb_LineSpace.FormattingEnabled = true;
		this.Cmb_LineSpace.Location = new System.Drawing.Point(269, 166);
		this.Cmb_LineSpace.Name = "Cmb_LineSpace";
		this.Cmb_LineSpace.Size = new System.Drawing.Size(111, 28);
		this.Cmb_LineSpace.TabIndex = 26;
		this.Cmb_LineSpace.TextChanged += new System.EventHandler(Cmb_LineSpace_TextChanged);
		this.Cmb_LineSpace.Validated += new System.EventHandler(Cmb_LineSpace_Validated);
		this.Txt_FirstLineIndent.Location = new System.Drawing.Point(79, 167);
		this.Txt_FirstLineIndent.Name = "Txt_FirstLineIndent";
		this.Txt_FirstLineIndent.Size = new System.Drawing.Size(111, 26);
		this.Txt_FirstLineIndent.TabIndex = 25;
		this.Txt_FirstLineIndent.TextChanged += new System.EventHandler(Txt_FirstLineIndent_TextChanged);
		this.Txt_FirstLineIndent.Validated += new System.EventHandler(Txt_FirstLineIndent_Validated);
		this.Txt_RightIndent.Location = new System.Drawing.Point(269, 131);
		this.Txt_RightIndent.Name = "Txt_RightIndent";
		this.Txt_RightIndent.Size = new System.Drawing.Size(111, 26);
		this.Txt_RightIndent.TabIndex = 24;
		this.Txt_RightIndent.TextChanged += new System.EventHandler(IndentSetting_TextChanged);
		this.Txt_RightIndent.Validated += new System.EventHandler(IndentChanged_Validated);
		this.Txt_LeftIndent.Location = new System.Drawing.Point(79, 131);
		this.Txt_LeftIndent.Name = "Txt_LeftIndent";
		this.Txt_LeftIndent.Size = new System.Drawing.Size(111, 26);
		this.Txt_LeftIndent.TabIndex = 23;
		this.Txt_LeftIndent.TextChanged += new System.EventHandler(IndentSetting_TextChanged);
		this.Txt_LeftIndent.Validated += new System.EventHandler(IndentChanged_Validated);
		this.label18.AutoSize = true;
		this.label18.Location = new System.Drawing.Point(199, 134);
		this.label18.Name = "label18";
		this.label18.Size = new System.Drawing.Size(51, 20);
		this.label18.TabIndex = 21;
		this.label18.Text = "右缩进";
		this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Txt_NumberFormat.Location = new System.Drawing.Point(269, 239);
		this.Txt_NumberFormat.Name = "Txt_NumberFormat";
		this.Txt_NumberFormat.Size = new System.Drawing.Size(111, 26);
		this.Txt_NumberFormat.TabIndex = 4;
		this.Txt_NumberFormat.TextChanged += new System.EventHandler(Txt_NumberFormat_TextChanged);
		this.label14.AutoSize = true;
		this.label14.Location = new System.Drawing.Point(199, 242);
		this.label14.Name = "label14";
		this.label14.Size = new System.Drawing.Size(65, 20);
		this.label14.TabIndex = 3;
		this.label14.Text = "编号格式";
		this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_NumberStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_NumberStyle.FormattingEnabled = true;
		this.Cmb_NumberStyle.Location = new System.Drawing.Point(79, 238);
		this.Cmb_NumberStyle.Name = "Cmb_NumberStyle";
		this.Cmb_NumberStyle.Size = new System.Drawing.Size(111, 28);
		this.Cmb_NumberStyle.TabIndex = 2;
		this.Cmb_NumberStyle.SelectedIndexChanged += new System.EventHandler(Cmb_NumberStyle_SelectedIndexChanged);
		this.label13.AutoSize = true;
		this.label13.Location = new System.Drawing.Point(8, 242);
		this.label13.Name = "label13";
		this.label13.Size = new System.Drawing.Size(65, 20);
		this.label13.TabIndex = 1;
		this.label13.Text = "编号样式";
		this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label11.AutoSize = true;
		this.label11.Location = new System.Drawing.Point(199, 206);
		this.label11.Name = "label11";
		this.label11.Size = new System.Drawing.Size(65, 20);
		this.label11.TabIndex = 19;
		this.label11.Text = "段后间距";
		this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Chk_BeforeBreak.AutoSize = true;
		this.Chk_BeforeBreak.Location = new System.Drawing.Point(176, 96);
		this.Chk_BeforeBreak.Name = "Chk_BeforeBreak";
		this.Chk_BeforeBreak.Size = new System.Drawing.Size(84, 24);
		this.Chk_BeforeBreak.TabIndex = 20;
		this.Chk_BeforeBreak.Text = "段前分页";
		this.Chk_BeforeBreak.UseVisualStyleBackColor = true;
		this.Chk_BeforeBreak.CheckedChanged += new System.EventHandler(FontStyleAndBreakPage_CheckedChanged);
		this.label10.AutoSize = true;
		this.label10.Location = new System.Drawing.Point(8, 206);
		this.label10.Name = "label10";
		this.label10.Size = new System.Drawing.Size(65, 20);
		this.label10.TabIndex = 17;
		this.label10.Text = "段前间距";
		this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label9.AutoSize = true;
		this.label9.Location = new System.Drawing.Point(199, 170);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(65, 20);
		this.label9.TabIndex = 15;
		this.label9.Text = "段落行距";
		this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label7.AutoSize = true;
		this.label7.Location = new System.Drawing.Point(8, 170);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(65, 20);
		this.label7.TabIndex = 11;
		this.label7.Text = "首行缩进";
		this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(8, 134);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(51, 20);
		this.label6.TabIndex = 8;
		this.label6.Text = "左缩进";
		this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_ParaAligment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_ParaAligment.FormattingEnabled = true;
		this.Cmb_ParaAligment.Location = new System.Drawing.Point(79, 94);
		this.Cmb_ParaAligment.Name = "Cmb_ParaAligment";
		this.Cmb_ParaAligment.Size = new System.Drawing.Size(86, 28);
		this.Cmb_ParaAligment.TabIndex = 15;
		this.Cmb_ParaAligment.SelectedIndexChanged += new System.EventHandler(FontNameAndAlignment_SelectedIndexChanged);
		this.label15.AutoSize = true;
		this.label15.Location = new System.Drawing.Point(8, 98);
		this.label15.Name = "label15";
		this.label15.Size = new System.Drawing.Size(65, 20);
		this.label15.TabIndex = 14;
		this.label15.Text = "段落对齐";
		this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Btn_FontColor.Location = new System.Drawing.Point(362, 59);
		this.Btn_FontColor.Name = "Btn_FontColor";
		this.Btn_FontColor.Size = new System.Drawing.Size(26, 26);
		this.Btn_FontColor.TabIndex = 8;
		this.Btn_FontColor.UseVisualStyleBackColor = true;
		this.Btn_FontColor.BackColorChanged += new System.EventHandler(Btn_FontColor_BackColorChanged);
		this.Btn_FontColor.Click += new System.EventHandler(Btn_FontColor_Click);
		this.Chk_UnderLine.AutoSize = true;
		this.Chk_UnderLine.Location = new System.Drawing.Point(290, 60);
		this.Chk_UnderLine.Name = "Chk_UnderLine";
		this.Chk_UnderLine.Size = new System.Drawing.Size(70, 24);
		this.Chk_UnderLine.TabIndex = 7;
		this.Chk_UnderLine.Text = "下划线";
		this.Chk_UnderLine.UseVisualStyleBackColor = true;
		this.Chk_UnderLine.CheckedChanged += new System.EventHandler(FontStyleAndBreakPage_CheckedChanged);
		this.Chk_Italic.AutoSize = true;
		this.Chk_Italic.Location = new System.Drawing.Point(234, 60);
		this.Chk_Italic.Name = "Chk_Italic";
		this.Chk_Italic.Size = new System.Drawing.Size(56, 24);
		this.Chk_Italic.TabIndex = 6;
		this.Chk_Italic.Text = "斜体";
		this.Chk_Italic.UseVisualStyleBackColor = true;
		this.Chk_Italic.CheckedChanged += new System.EventHandler(FontStyleAndBreakPage_CheckedChanged);
		this.Chk_Bold.AutoSize = true;
		this.Chk_Bold.Location = new System.Drawing.Point(176, 60);
		this.Chk_Bold.Name = "Chk_Bold";
		this.Chk_Bold.Size = new System.Drawing.Size(56, 24);
		this.Chk_Bold.TabIndex = 5;
		this.Chk_Bold.Text = "粗体";
		this.Chk_Bold.UseVisualStyleBackColor = true;
		this.Chk_Bold.CheckedChanged += new System.EventHandler(FontStyleAndBreakPage_CheckedChanged);
		this.Cmb_FontSize.FormattingEnabled = true;
		this.Cmb_FontSize.Location = new System.Drawing.Point(79, 58);
		this.Cmb_FontSize.Name = "Cmb_FontSize";
		this.Cmb_FontSize.Size = new System.Drawing.Size(86, 28);
		this.Cmb_FontSize.TabIndex = 4;
		this.Cmb_FontSize.TextChanged += new System.EventHandler(Cmb_FontSize_TextChanged);
		this.Cmb_FontSize.Validated += new System.EventHandler(Cmb_FontSize_Validated);
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(8, 62);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(65, 20);
		this.label4.TabIndex = 3;
		this.label4.Text = "字体大小";
		this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_EngFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_EngFontName.FormattingEnabled = true;
		this.Cmb_EngFontName.Location = new System.Drawing.Point(269, 22);
		this.Cmb_EngFontName.Name = "Cmb_EngFontName";
		this.Cmb_EngFontName.Size = new System.Drawing.Size(111, 28);
		this.Cmb_EngFontName.TabIndex = 4;
		this.Cmb_EngFontName.SelectedIndexChanged += new System.EventHandler(FontNameAndAlignment_SelectedIndexChanged);
		this.label17.AutoSize = true;
		this.label17.Location = new System.Drawing.Point(199, 26);
		this.label17.Name = "label17";
		this.label17.Size = new System.Drawing.Size(65, 20);
		this.label17.TabIndex = 3;
		this.label17.Text = "西文字体";
		this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Cmb_ChnFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_ChnFontName.FormattingEnabled = true;
		this.Cmb_ChnFontName.Location = new System.Drawing.Point(79, 22);
		this.Cmb_ChnFontName.Name = "Cmb_ChnFontName";
		this.Cmb_ChnFontName.Size = new System.Drawing.Size(111, 28);
		this.Cmb_ChnFontName.TabIndex = 2;
		this.Cmb_ChnFontName.SelectedIndexChanged += new System.EventHandler(FontNameAndAlignment_SelectedIndexChanged);
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(8, 26);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(65, 20);
		this.label3.TabIndex = 1;
		this.label3.Text = "中文字体";
		this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Txt_NewStyleName.Location = new System.Drawing.Point(10, 299);
		this.Txt_NewStyleName.Name = "Txt_NewStyleName";
		this.Txt_NewStyleName.Size = new System.Drawing.Size(170, 26);
		this.Txt_NewStyleName.TabIndex = 33;
		this.Txt_NewStyleName.WaterMark = "输入添加样式的名称";
		this.Txt_NewStyleName.WaterTextColor = System.Drawing.Color.SteelBlue;
		this.Lab_StyleInfo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
		this.Lab_StyleInfo.Location = new System.Drawing.Point(189, 304);
		this.Lab_StyleInfo.Name = "Lab_StyleInfo";
		this.Lab_StyleInfo.Size = new System.Drawing.Size(399, 53);
		this.Lab_StyleInfo.TabIndex = 30;
		this.Lab_StyleInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Btn_DelStyle.Enabled = false;
		this.Btn_DelStyle.Location = new System.Drawing.Point(104, 327);
		this.Btn_DelStyle.Name = "Btn_DelStyle";
		this.Btn_DelStyle.Size = new System.Drawing.Size(76, 29);
		this.Btn_DelStyle.TabIndex = 29;
		this.Btn_DelStyle.Text = "删除样式";
		this.Btn_DelStyle.UseVisualStyleBackColor = true;
		this.Btn_DelStyle.Click += new System.EventHandler(Btn_DelStyle_Click);
		this.Btn_AddStyle.Location = new System.Drawing.Point(10, 327);
		this.Btn_AddStyle.Name = "Btn_AddStyle";
		this.Btn_AddStyle.Size = new System.Drawing.Size(76, 29);
		this.Btn_AddStyle.TabIndex = 27;
		this.Btn_AddStyle.Text = "添加样式";
		this.Btn_AddStyle.UseVisualStyleBackColor = true;
		this.Btn_AddStyle.Click += new System.EventHandler(Btn_AddStyle_Click);
		this.Chk_CreateListLevels.AutoSize = true;
		this.Chk_CreateListLevels.Location = new System.Drawing.Point(235, 511);
		this.Chk_CreateListLevels.Name = "Chk_CreateListLevels";
		this.Chk_CreateListLevels.Size = new System.Drawing.Size(56, 24);
		this.Chk_CreateListLevels.TabIndex = 22;
		this.Chk_CreateListLevels.Text = "创建";
		this.Chk_CreateListLevels.UseVisualStyleBackColor = true;
		this.Chk_CreateListLevels.CheckedChanged += new System.EventHandler(Chk_CreateListLevels_CheckedChanged);
		this.Cmb_SetLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_SetLevel.Enabled = false;
		this.Cmb_SetLevel.FormattingEnabled = true;
		this.Cmb_SetLevel.Items.AddRange(new object[9] { "1", "2", "3", "4", "5", "6", "7", "8", "9" });
		this.Cmb_SetLevel.Location = new System.Drawing.Point(295, 509);
		this.Cmb_SetLevel.Name = "Cmb_SetLevel";
		this.Cmb_SetLevel.Size = new System.Drawing.Size(59, 28);
		this.Cmb_SetLevel.TabIndex = 21;
		this.label12.AutoSize = true;
		this.label12.Location = new System.Drawing.Point(365, 512);
		this.label12.Name = "label12";
		this.label12.Size = new System.Drawing.Size(79, 20);
		this.label12.TabIndex = 20;
		this.label12.Text = "级多级标题";
		this.Btn_ApplySet.Location = new System.Drawing.Point(491, 541);
		this.Btn_ApplySet.Name = "Btn_ApplySet";
		this.Btn_ApplySet.Size = new System.Drawing.Size(100, 30);
		this.Btn_ApplySet.TabIndex = 7;
		this.Btn_ApplySet.Text = "应用样式";
		this.Btn_ApplySet.UseVisualStyleBackColor = true;
		this.Btn_ApplySet.Click += new System.EventHandler(Btn_ApplySet_Click);
		this.Cmb_PreSettings.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PreSettings.FormattingEnabled = true;
		this.Cmb_PreSettings.Items.AddRange(new object[4] { "公文风格", "论文风格", "报告风格", "条文风格" });
		this.Cmb_PreSettings.Location = new System.Drawing.Point(82, 509);
		this.Cmb_PreSettings.Name = "Cmb_PreSettings";
		this.Cmb_PreSettings.Size = new System.Drawing.Size(120, 28);
		this.Cmb_PreSettings.TabIndex = 8;
		this.Cmb_PreSettings.SelectedIndexChanged += new System.EventHandler(Cmb_PreSettings_SelectedIndexChanged);
		this.label16.AutoSize = true;
		this.label16.Location = new System.Drawing.Point(13, 512);
		this.label16.Name = "label16";
		this.label16.Size = new System.Drawing.Size(65, 20);
		this.label16.TabIndex = 21;
		this.label16.Text = "预设样式";
		this.Btn_SaveStyles.Location = new System.Drawing.Point(121, 541);
		this.Btn_SaveStyles.Name = "Btn_SaveStyles";
		this.Btn_SaveStyles.Size = new System.Drawing.Size(100, 30);
		this.Btn_SaveStyles.TabIndex = 23;
		this.Btn_SaveStyles.Text = "保存样式";
		this.Btn_SaveStyles.UseVisualStyleBackColor = true;
		this.Btn_SaveStyles.Click += new System.EventHandler(Btn_SaveStyles_Click);
		this.Btn_LoadStyles.Location = new System.Drawing.Point(15, 541);
		this.Btn_LoadStyles.Name = "Btn_LoadStyles";
		this.Btn_LoadStyles.Size = new System.Drawing.Size(100, 30);
		this.Btn_LoadStyles.TabIndex = 24;
		this.Btn_LoadStyles.Text = "读取样式";
		this.Btn_LoadStyles.UseVisualStyleBackColor = true;
		this.Btn_LoadStyles.Click += new System.EventHandler(Btn_LoadStyles_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Btn_LoadStyles);
		base.Controls.Add(this.Btn_SaveStyles);
		base.Controls.Add(this.Cmb_SetLevel);
		base.Controls.Add(this.label16);
		base.Controls.Add(this.Cmb_PreSettings);
		base.Controls.Add(this.Btn_ApplySet);
		base.Controls.Add(this.groupBox3);
		base.Controls.Add(this.Grp_PageSetup);
		base.Controls.Add(this.Chk_CreateListLevels);
		base.Controls.Add(this.groupBox1);
		base.Controls.Add(this.label12);
		this.DoubleBuffered = true;
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Name = "StyleSetGuider";
		base.Size = new System.Drawing.Size(600, 575);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.Grp_PageSetup.ResumeLayout(false);
		this.Grp_PageSetup.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_GutterValue).EndInit();
		this.groupBox3.ResumeLayout(false);
		this.groupBox3.PerformLayout();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
