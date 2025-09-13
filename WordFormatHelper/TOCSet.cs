using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class TOCSet : UserControl
{
	private readonly List<string> FontSizeCha = new List<string>(16)
	{
		"八号", "七号", "小六", "六号", "小五", "五号", "小四", "四号", "小三", "三号",
		"小二", "二号", "小一", "一号", "小初", "初号"
	};

	private readonly List<float> FontSizePoint = new List<float>(16)
	{
		5f, 5.5f, 6.5f, 7.5f, 9f, 10.5f, 12f, 14f, 15f, 16f,
		18f, 22f, 24f, 26f, 36f, 42f
	};

	private static readonly List<CustomStyle> TOCStyles = new List<CustomStyle>(9)
	{
		new CustomStyle("wdStyleTOC1", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC2", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC3", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC4", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC5", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC6", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC7", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC8", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false),
		new CustomStyle("wdStyleTOC9", null, 0f, bold: false, italic: false, underline: false, 0, 0f, 0f, 0, 0f, beforebreak: false, 0f, 0f, 0, null, userdefined: false)
	};

	private bool NotUserChanged;

	private static bool Initialized;

	private readonly float[] GapFromNumberToText = new float[3]
	{
		Globals.ThisAddIn.Application.CentimetersToPoints(1f),
		Globals.ThisAddIn.Application.CentimetersToPoints(0.5f),
		Globals.ThisAddIn.Application.CentimetersToPoints(1.5f)
	};

	private static TocSettings currentSetting = new TocSettings();

	private IContainer components;

	private Label label1;

	private Label label2;

	private CheckBox Chk_TOCUsePageNumber;

	private ComboBox Cmb_TOCLevel;

	private ComboBox Cmb_TOCIndent;

	private ComboBox Cmb_PageNumberLeader;

	private ComboBox Cmb_FontName;

	private ComboBox Cmb_FontSize;

	private ToggleButton Tog_FontBold;

	private ToggleButton Tog_FontItalic;

	private ComboBox Cmb_TOCLevelStyle;

	private GroupBox groupBox1;

	private Label label6;

	private Label label5;

	private Label label4;

	private Label label9;

	private Label label8;

	private Label label7;

	private NumericUpDownWithUnit Nud_AfterSpace;

	private NumericUpDownWithUnit Nud_BeforeSpace;

	private NumericUpDownWithUnit Nud_LineSpace;

	private Button Btn_InsertTOC;

	private Panel Pan_Style;

	private CheckBox Chk_SetAllTOCStyle;

	private CheckBox Chk_ReplaceCurrentContents;

	private ComboBox Cmb_GapFromNumberToText;

	private CheckBox Chk_TryAlignNumber;

	public TOCSet()
	{
		InitializeComponent();
		NotUserChanged = true;
		Cmb_TOCLevel.SelectedIndex = currentSetting.Levels;
		Cmb_TOCIndent.SelectedIndex = currentSetting.IndentStyle;
		Chk_TOCUsePageNumber.Checked = currentSetting.UsePageNumber;
		Cmb_PageNumberLeader.Enabled = currentSetting.UsePageNumber;
		Cmb_PageNumberLeader.SelectedIndex = currentSetting.Leader;
		Cmb_GapFromNumberToText.SelectedIndex = currentSetting.IndentGap;
		Chk_ReplaceCurrentContents.Checked = currentSetting.ReplaceCurrentTOC;
		Chk_TryAlignNumber.Checked = currentSetting.TryAlignPageNumber;
		FontFamily[] families = new InstalledFontCollection().Families;
		foreach (FontFamily fontFamily in families)
		{
			Cmb_FontName.Items.Add(fontFamily.Name);
		}
		ComboBox.ObjectCollection items = Cmb_FontSize.Items;
		object[] items2 = FontSizeCha.ToArray();
		items.AddRange(items2);
		Cmb_TOCLevelStyle.SelectedIndex = 0;
		NotUserChanged = false;
		if (!Initialized)
		{
			for (int j = 1; j < 10; j++)
			{
				string fName = "wdStyleTOC" + j;
				Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				object Index = WdBuiltinStyle.wdStyleNormal;
				SetBodyTextStyle(fName, styles[ref Index]);
			}
			Initialized = true;
		}
		CustomStyle customStyle = TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC1").FirstOrDefault();
		customStyle.IsBold = true;
		customStyle.BeforeSpacing = 0.5f;
		UpdataTOCLevelStyle("wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1));
	}

	private void Chk_TOCUsePageNumber_CheckedChanged(object sender, EventArgs e)
	{
		bool flag = (sender as CheckBox).Checked;
		switch ((sender as CheckBox).Name)
		{
		case "Chk_TOCUsePageNumber":
			Cmb_PageNumberLeader.Enabled = Chk_TOCUsePageNumber.Checked;
			currentSetting.UsePageNumber = flag;
			break;
		case "Chk_ReplaceCurrentContents":
			currentSetting.ReplaceCurrentTOC = flag;
			break;
		case "Chk_TryAlignNumber":
			currentSetting.TryAlignPageNumber = flag;
			break;
		}
	}

	private void Cmb_TOCLevel_SelectedIndexChanged(object sender, EventArgs e)
	{
		int selectedIndex = Cmb_TOCLevelStyle.SelectedIndex;
		Cmb_TOCLevelStyle.Items.Clear();
		for (int i = 0; i <= Cmb_TOCLevel.SelectedIndex; i++)
		{
			Cmb_TOCLevelStyle.Items.Add("第" + (i + 1) + "级");
		}
		if (selectedIndex < Cmb_TOCLevelStyle.Items.Count)
		{
			Cmb_TOCLevelStyle.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_TOCLevelStyle.SelectedIndex = 0;
		}
		currentSetting.Levels = Cmb_TOCLevel.SelectedIndex;
	}

	private void UpdataTOCLevelStyle(string fName)
	{
		NotUserChanged = true;
		CustomStyle customStyle = TOCStyles.Where((CustomStyle c) => c.Name == fName).FirstOrDefault();
		Cmb_FontName.SelectedIndex = Cmb_FontName.Items.IndexOf(customStyle.FontName);
		if (FontSizePoint.IndexOf(customStyle.FontSize) != -1)
		{
			Cmb_FontSize.Text = null;
			Cmb_FontSize.SelectedIndex = FontSizePoint.IndexOf(customStyle.FontSize);
		}
		else
		{
			Cmb_FontSize.Text = customStyle.FontSize.ToString();
		}
		Tog_FontBold.Pressed = customStyle.IsBold;
		Tog_FontItalic.Pressed = customStyle.IsItalic;
		Nud_LineSpace.Value = (decimal)customStyle.LineSpacing;
		Nud_BeforeSpace.Value = (decimal)customStyle.BeforeSpacing;
		Nud_AfterSpace.Value = (decimal)customStyle.AfterSpacing;
		NotUserChanged = false;
	}

	private void SetBodyTextStyle(string fName, Style BodyText)
	{
		CustomStyle customStyle = TOCStyles.Where((CustomStyle c) => c.Name == fName).FirstOrDefault();
		customStyle.FontName = BodyText.Font.Name;
		customStyle.FontSize = BodyText.Font.Size;
		customStyle.LineSpacing = Globals.ThisAddIn.Application.PointsToLines(BodyText.ParagraphFormat.LineSpacing);
		customStyle.BeforeSpacing = Globals.ThisAddIn.Application.PointsToLines(BodyText.ParagraphFormat.SpaceBefore);
		customStyle.AfterSpacing = Globals.ThisAddIn.Application.PointsToLines(BodyText.ParagraphFormat.SpaceAfter);
	}

	private void Cmb_TOCLevelStyle_SelectedIndexChanged(object sender, EventArgs e)
	{
		UpdataTOCLevelStyle("wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1));
	}

	private void Cmb_FontName_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (NotUserChanged)
		{
			return;
		}
		if (Chk_SetAllTOCStyle.Checked)
		{
			foreach (CustomStyle tOCStyle in TOCStyles)
			{
				if (TOCStyles.IndexOf(tOCStyle) > Cmb_TOCLevel.SelectedIndex)
				{
					break;
				}
				tOCStyle.FontName = Cmb_FontName.Text;
			}
			return;
		}
		TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1)).FirstOrDefault().FontName = Cmb_FontName.Text;
	}

	private void Tog_FontBold_Click(object sender, EventArgs e)
	{
		if (NotUserChanged)
		{
			return;
		}
		string name;
		if (Chk_SetAllTOCStyle.Checked)
		{
			foreach (CustomStyle tOCStyle in TOCStyles)
			{
				if (TOCStyles.IndexOf(tOCStyle) > Cmb_TOCLevel.SelectedIndex)
				{
					break;
				}
				name = (sender as ToggleButton).Name;
				if (!(name == "Tog_FontBold"))
				{
					if (name == "Tog_FontItalic")
					{
						tOCStyle.IsItalic = (sender as ToggleButton).Pressed;
					}
				}
				else
				{
					tOCStyle.IsBold = (sender as ToggleButton).Pressed;
				}
			}
			return;
		}
		CustomStyle customStyle = TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1)).FirstOrDefault();
		name = (sender as ToggleButton).Name;
		if (!(name == "Tog_FontBold"))
		{
			if (name == "Tog_FontItalic")
			{
				customStyle.IsItalic = (sender as ToggleButton).Pressed;
			}
		}
		else
		{
			customStyle.IsBold = (sender as ToggleButton).Pressed;
		}
	}

	private void Nud_LineSpace_ValueChanged(object sender, EventArgs e)
	{
		if (NotUserChanged)
		{
			return;
		}
		if (Chk_SetAllTOCStyle.Checked)
		{
			foreach (CustomStyle tOCStyle in TOCStyles)
			{
				if (TOCStyles.IndexOf(tOCStyle) > Cmb_TOCLevel.SelectedIndex)
				{
					break;
				}
				switch ((sender as NumericUpDownWithUnit).Name)
				{
				case "Nud_LineSpace":
					tOCStyle.LineSpacing = (float)(sender as NumericUpDownWithUnit).Value;
					break;
				case "Nud_BeforeSpace":
					tOCStyle.BeforeSpacing = (float)(sender as NumericUpDownWithUnit).Value;
					break;
				case "Nud_AfterSpace":
					tOCStyle.AfterSpacing = (float)(sender as NumericUpDownWithUnit).Value;
					break;
				}
			}
			return;
		}
		CustomStyle customStyle = TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1)).FirstOrDefault();
		switch ((sender as NumericUpDownWithUnit).Name)
		{
		case "Nud_LineSpace":
			customStyle.LineSpacing = (float)(sender as NumericUpDownWithUnit).Value;
			break;
		case "Nud_BeforeSpace":
			customStyle.BeforeSpacing = (float)(sender as NumericUpDownWithUnit).Value;
			break;
		case "Nud_AfterSpace":
			customStyle.AfterSpacing = (float)(sender as NumericUpDownWithUnit).Value;
			break;
		}
	}

	private void Cmb_FontSize_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (NotUserChanged)
		{
			return;
		}
		if (Chk_SetAllTOCStyle.Checked)
		{
			foreach (CustomStyle tOCStyle in TOCStyles)
			{
				if (TOCStyles.IndexOf(tOCStyle) > Cmb_TOCLevel.SelectedIndex)
				{
					break;
				}
				tOCStyle.FontSize = FontSizePoint[Cmb_FontSize.SelectedIndex];
			}
			return;
		}
		TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1)).FirstOrDefault().FontSize = FontSizePoint[Cmb_FontSize.SelectedIndex];
	}

	private void Cmb_FontSize_Leave(object sender, EventArgs e)
	{
		if (NotUserChanged || Cmb_FontSize.SelectedIndex != -1 || !Regex.IsMatch(Cmb_FontSize.Text, "^[1-9]{1,4}(\\.5|\\.0){0,1}$"))
		{
			return;
		}
		float num = Convert.ToSingle(Cmb_FontSize.Text);
		if (!(num >= 1f) && !(num <= 1638f))
		{
			return;
		}
		if (Chk_SetAllTOCStyle.Checked)
		{
			foreach (CustomStyle tOCStyle in TOCStyles)
			{
				tOCStyle.FontSize = num;
			}
			return;
		}
		TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC" + (Cmb_TOCLevelStyle.SelectedIndex + 1)).FirstOrDefault().FontSize = num;
	}

	private void Btn_InsertTOC_Click(object sender, EventArgs e)
	{
		int num = Cmb_TOCLevel.SelectedIndex + 1;
		string[] array = new string[9] { "8", "8.8", "8.8.8", "8.8.8.8", "8.8.8.8.8", "8.8.8.8.8.8", "8.8.8.8.8.8.8", "8.8.8.8.8.8.8.8", "8.8.8.8.8.8.8.8.8" };
		int[] array2 = new int[num];
		float num2 = 0f;
		Section first = Globals.ThisAddIn.Application.Selection.Sections.First;
		float num3 = first.PageSetup.PageWidth - first.PageSetup.LeftMargin - first.PageSetup.RightMargin;
		WdTabLeader[] array3 = new WdTabLeader[5]
		{
			WdTabLeader.wdTabLeaderSpaces,
			WdTabLeader.wdTabLeaderDots,
			WdTabLeader.wdTabLeaderDashes,
			WdTabLeader.wdTabLeaderLines,
			WdTabLeader.wdTabLeaderMiddleDot
		};
		Range range = Globals.ThisAddIn.Application.Selection.Range;
		if (Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents.Count > 0)
		{
			if (Chk_ReplaceCurrentContents.Checked)
			{
				range = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents[1].Range;
			}
			else if (MessageBox.Show("是否替换目录？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
			{
				range = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents[1].Range;
			}
		}
		if (first.PageSetup.Gutter > 0f && first.PageSetup.GutterPos != WdGutterStyle.wdGutterPosTop)
		{
			num3 -= first.PageSetup.Gutter;
		}
		int i;
		object Index;
		for (i = 0; i < num; i++)
		{
			Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
			Index = (WdBuiltinStyle)(-20 - i);
			Style style = styles[ref Index];
			CustomStyle customStyle = TOCStyles.Where((CustomStyle c) => c.Name == "wdStyleTOC" + (i + 1)).FirstOrDefault();
			array2[i] = TextRenderer.MeasureText(array[i], new System.Drawing.Font(new FontFamily(customStyle.FontName), customStyle.FontSize)).Width;
			Index = "";
			style.set_BaseStyle(ref Index);
			style.Font.Name = customStyle.FontName;
			style.Font.Size = customStyle.FontSize;
			style.Font.Bold = (customStyle.IsBold ? (-1) : 0);
			style.Font.Italic = (customStyle.IsItalic ? (-1) : 0);
			style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
			style.ParagraphFormat.LeftIndent = 0f;
			style.ParagraphFormat.FirstLineIndent = 0f;
			style.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			style.ParagraphFormat.CharacterUnitLeftIndent = 0f;
			style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
			style.ParagraphFormat.LineSpacing = Globals.ThisAddIn.Application.LinesToPoints(customStyle.LineSpacing);
			style.ParagraphFormat.SpaceBefore = Globals.ThisAddIn.Application.LinesToPoints(customStyle.BeforeSpacing);
			style.ParagraphFormat.SpaceAfter = Globals.ThisAddIn.Application.LinesToPoints(customStyle.AfterSpacing);
		}
		object fVertical;
		object Index2;
		switch (Cmb_TOCIndent.SelectedIndex)
		{
		case 0:
		{
			Microsoft.Office.Interop.Word.Application application3 = Globals.ThisAddIn.Application;
			float pixels3 = array2[array2.Length - 1];
			Index = Type.Missing;
			num2 = application3.PixelsToPoints(pixels3, ref Index) + GapFromNumberToText[Cmb_GapFromNumberToText.SelectedIndex];
			for (int num6 = 0; num6 < num; num6++)
			{
				Styles styles3 = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				Index = (WdBuiltinStyle)(-20 - num6);
				Style style3 = styles3[ref Index];
				style3.ParagraphFormat.FirstLineIndent = 0f - num2;
				style3.ParagraphFormat.LeftIndent = num2;
				style3.ParagraphFormat.TabStops.ClearAll();
				TabStops tabStops3 = style3.ParagraphFormat.TabStops;
				float position3 = num2;
				Index = WdTabAlignment.wdAlignTabLeft;
				Index2 = Type.Missing;
				tabStops3.Add(position3, ref Index, ref Index2);
				TabStops tabStops4 = style3.ParagraphFormat.TabStops;
				float position4 = num3;
				Index2 = WdTabAlignment.wdAlignTabRight;
				Index = array3[Cmb_PageNumberLeader.SelectedIndex];
				tabStops4.Add(position4, ref Index2, ref Index);
			}
			break;
		}
		case 1:
		{
			for (int num7 = 0; num7 < num; num7++)
			{
				Styles styles4 = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				Index = (WdBuiltinStyle)(-20 - num7);
				Style style4 = styles4[ref Index];
				if (num7 == 0)
				{
					Microsoft.Office.Interop.Word.Application application4 = Globals.ThisAddIn.Application;
					float pixels4 = array2[0];
					Index = Type.Missing;
					num2 = application4.PixelsToPoints(pixels4, ref Index) + GapFromNumberToText[Cmb_GapFromNumberToText.SelectedIndex];
					style4.ParagraphFormat.FirstLineIndent = 0f - num2;
					style4.ParagraphFormat.LeftIndent = num2;
					style4.ParagraphFormat.TabStops.ClearAll();
					TabStops tabStops5 = style4.ParagraphFormat.TabStops;
					float position5 = num2;
					Index = WdTabAlignment.wdAlignTabLeft;
					Index2 = Type.Missing;
					tabStops5.Add(position5, ref Index, ref Index2);
				}
				else
				{
					Microsoft.Office.Interop.Word.Application application5 = Globals.ThisAddIn.Application;
					float pixels5 = array2[array2.Length - 1];
					Index2 = Type.Missing;
					num2 = application5.PixelsToPoints(pixels5, ref Index2) + GapFromNumberToText[Cmb_GapFromNumberToText.SelectedIndex];
					style4.ParagraphFormat.FirstLineIndent = 0f - num2;
					ParagraphFormat paragraphFormat2 = style4.ParagraphFormat;
					Microsoft.Office.Interop.Word.Application application6 = Globals.ThisAddIn.Application;
					float pixels6 = array2[0];
					Index2 = Type.Missing;
					paragraphFormat2.LeftIndent = application6.PixelsToPoints(pixels6, ref Index2) + 5f + num2;
					style4.ParagraphFormat.TabStops.ClearAll();
					TabStops tabStops6 = style4.ParagraphFormat.TabStops;
					Microsoft.Office.Interop.Word.Application application7 = Globals.ThisAddIn.Application;
					float pixels7 = array2[0];
					fVertical = Type.Missing;
					float position6 = application7.PixelsToPoints(pixels7, ref fVertical) + 5f + num2;
					Index2 = WdTabAlignment.wdAlignTabLeft;
					Index = Type.Missing;
					tabStops6.Add(position6, ref Index2, ref Index);
				}
				TabStops tabStops7 = style4.ParagraphFormat.TabStops;
				float position7 = num3;
				Index = WdTabAlignment.wdAlignTabRight;
				Index2 = array3[Cmb_PageNumberLeader.SelectedIndex];
				tabStops7.Add(position7, ref Index, ref Index2);
			}
			break;
		}
		case 2:
		{
			for (int num4 = 0; num4 < num; num4++)
			{
				Styles styles2 = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				Index2 = (WdBuiltinStyle)(-20 - num4);
				Style style2 = styles2[ref Index2];
				float num5 = num2;
				Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
				float pixels = array2[num4];
				Index2 = Type.Missing;
				num2 = num5 + application.PixelsToPoints(pixels, ref Index2) + GapFromNumberToText[Cmb_GapFromNumberToText.SelectedIndex];
				ParagraphFormat paragraphFormat = style2.ParagraphFormat;
				Microsoft.Office.Interop.Word.Application application2 = Globals.ThisAddIn.Application;
				float pixels2 = array2[num4];
				Index2 = Type.Missing;
				paragraphFormat.FirstLineIndent = 0f - (application2.PixelsToPoints(pixels2, ref Index2) + GapFromNumberToText[Cmb_GapFromNumberToText.SelectedIndex]);
				style2.ParagraphFormat.LeftIndent = num2;
				style2.ParagraphFormat.TabStops.ClearAll();
				TabStops tabStops = style2.ParagraphFormat.TabStops;
				float position = num2;
				Index2 = WdTabAlignment.wdAlignTabLeft;
				Index = Type.Missing;
				tabStops.Add(position, ref Index2, ref Index);
				TabStops tabStops2 = style2.ParagraphFormat.TabStops;
				float position2 = num3;
				Index = WdTabAlignment.wdAlignTabRight;
				Index2 = array3[Cmb_PageNumberLeader.SelectedIndex];
				tabStops2.Add(position2, ref Index, ref Index2);
			}
			break;
		}
		}
		TablesOfContents tablesOfContents = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents;
		Range range2 = range;
		Index = false;
		fVertical = 1;
		object LowerHeadingLevel = num;
		Index2 = Chk_TOCUsePageNumber.Checked;
		object UseFields = Type.Missing;
		object TableID = Type.Missing;
		object RightAlignPageNumbers = Type.Missing;
		object IncludePageNumbers = Index2;
		object AddedStyles = Type.Missing;
		object UseHyperlinks = Type.Missing;
		object HidePageNumbersInWeb = Type.Missing;
		object UseOutlineLevels = Type.Missing;
		TableOfContents tableOfContents = tablesOfContents.Add(range2, ref Index, ref fVertical, ref LowerHeadingLevel, ref UseFields, ref TableID, ref RightAlignPageNumbers, ref IncludePageNumbers, ref AddedStyles, ref UseHyperlinks, ref HidePageNumbersInWeb, ref UseOutlineLevels);
		if (!Chk_TryAlignNumber.Checked)
		{
			return;
		}
		tableOfContents.RightAlignPageNumbers = true;
		foreach (Paragraph paragraph in tableOfContents.Range.Paragraphs)
		{
			if (Regex.IsMatch(paragraph.Range.Text, "^[^\\u0009]*\\u0009[^\\u0009]*$"))
			{
				Find find = paragraph.Range.Find;
				UseOutlineLevels = "\t";
				HidePageNumbersInWeb = Type.Missing;
				UseHyperlinks = Type.Missing;
				AddedStyles = Type.Missing;
				IncludePageNumbers = Type.Missing;
				RightAlignPageNumbers = Type.Missing;
				TableID = Type.Missing;
				UseFields = Type.Missing;
				LowerHeadingLevel = Type.Missing;
				fVertical = "\t\t";
				Index = WdReplace.wdReplaceOne;
				Index2 = Type.Missing;
				object MatchDiacritics = Type.Missing;
				object MatchAlefHamza = Type.Missing;
				object MatchControl = Type.Missing;
				find.Execute(ref UseOutlineLevels, ref HidePageNumbersInWeb, ref UseHyperlinks, ref AddedStyles, ref IncludePageNumbers, ref RightAlignPageNumbers, ref TableID, ref UseFields, ref LowerHeadingLevel, ref fVertical, ref Index, ref Index2, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
			}
		}
	}

	private void Chk_SetAllTOCStyle_CheckedChanged(object sender, EventArgs e)
	{
		Cmb_TOCLevelStyle.Enabled = !Chk_SetAllTOCStyle.Checked;
	}

	private void Cmb_PageNumberLeader_SelectedIndexChanged(object sender, EventArgs e)
	{
		int selectedIndex = (sender as ComboBox).SelectedIndex;
		switch ((sender as ComboBox).Name)
		{
		case "Cmb_PageNumberLeader":
			currentSetting.Leader = selectedIndex;
			break;
		case "Cmb_TOCIndent":
			currentSetting.IndentStyle = selectedIndex;
			break;
		case "Cmb_GapFromNumberToText":
			currentSetting.IndentGap = selectedIndex;
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
		this.label1 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.Chk_TOCUsePageNumber = new System.Windows.Forms.CheckBox();
		this.Cmb_TOCLevel = new System.Windows.Forms.ComboBox();
		this.Cmb_TOCIndent = new System.Windows.Forms.ComboBox();
		this.Cmb_PageNumberLeader = new System.Windows.Forms.ComboBox();
		this.Cmb_FontName = new System.Windows.Forms.ComboBox();
		this.Cmb_FontSize = new System.Windows.Forms.ComboBox();
		this.Cmb_TOCLevelStyle = new System.Windows.Forms.ComboBox();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.Chk_SetAllTOCStyle = new System.Windows.Forms.CheckBox();
		this.Pan_Style = new System.Windows.Forms.Panel();
		this.label9 = new System.Windows.Forms.Label();
		this.label5 = new System.Windows.Forms.Label();
		this.label8 = new System.Windows.Forms.Label();
		this.label7 = new System.Windows.Forms.Label();
		this.label6 = new System.Windows.Forms.Label();
		this.label4 = new System.Windows.Forms.Label();
		this.Btn_InsertTOC = new System.Windows.Forms.Button();
		this.Chk_ReplaceCurrentContents = new System.Windows.Forms.CheckBox();
		this.Cmb_GapFromNumberToText = new System.Windows.Forms.ComboBox();
		this.Chk_TryAlignNumber = new System.Windows.Forms.CheckBox();
		this.Tog_FontBold = new WordFormatHelper.ToggleButton();
		this.Nud_AfterSpace = new WordFormatHelper.NumericUpDownWithUnit();
		this.Nud_BeforeSpace = new WordFormatHelper.NumericUpDownWithUnit();
		this.Tog_FontItalic = new WordFormatHelper.ToggleButton();
		this.Nud_LineSpace = new WordFormatHelper.NumericUpDownWithUnit();
		this.groupBox1.SuspendLayout();
		this.Pan_Style.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_AfterSpace).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_BeforeSpace).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_LineSpace).BeginInit();
		base.SuspendLayout();
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(15, 14);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(65, 20);
		this.label1.TabIndex = 0;
		this.label1.Text = "目录级数";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(15, 50);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(65, 20);
		this.label2.TabIndex = 1;
		this.label2.Text = "缩进样式";
		this.Chk_TOCUsePageNumber.AutoSize = true;
		this.Chk_TOCUsePageNumber.Location = new System.Drawing.Point(161, 13);
		this.Chk_TOCUsePageNumber.Name = "Chk_TOCUsePageNumber";
		this.Chk_TOCUsePageNumber.Size = new System.Drawing.Size(140, 24);
		this.Chk_TOCUsePageNumber.TabIndex = 2;
		this.Chk_TOCUsePageNumber.Text = "使用页码，前导符";
		this.Chk_TOCUsePageNumber.UseVisualStyleBackColor = true;
		this.Chk_TOCUsePageNumber.CheckedChanged += new System.EventHandler(Chk_TOCUsePageNumber_CheckedChanged);
		this.Cmb_TOCLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TOCLevel.FormattingEnabled = true;
		this.Cmb_TOCLevel.Items.AddRange(new object[9] { "1级", "2级", "3级", "4级", "5级", "6级", "7级", "8级", "9级" });
		this.Cmb_TOCLevel.Location = new System.Drawing.Point(86, 11);
		this.Cmb_TOCLevel.Name = "Cmb_TOCLevel";
		this.Cmb_TOCLevel.Size = new System.Drawing.Size(61, 28);
		this.Cmb_TOCLevel.TabIndex = 4;
		this.Cmb_TOCLevel.SelectedIndexChanged += new System.EventHandler(Cmb_TOCLevel_SelectedIndexChanged);
		this.Cmb_TOCIndent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TOCIndent.FormattingEnabled = true;
		this.Cmb_TOCIndent.Items.AddRange(new object[3] { "所有标题序号顶格对齐，文本对齐", "一级标题顶格，后续标题统一缩进", "所有标题递进式缩进" });
		this.Cmb_TOCIndent.Location = new System.Drawing.Point(86, 47);
		this.Cmb_TOCIndent.Name = "Cmb_TOCIndent";
		this.Cmb_TOCIndent.Size = new System.Drawing.Size(251, 28);
		this.Cmb_TOCIndent.TabIndex = 5;
		this.Cmb_TOCIndent.SelectedIndexChanged += new System.EventHandler(Cmb_PageNumberLeader_SelectedIndexChanged);
		this.Cmb_PageNumberLeader.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_PageNumberLeader.FormattingEnabled = true;
		this.Cmb_PageNumberLeader.Items.AddRange(new object[5] { "无符号", "...... 底部虚点", "--- 短划线", "___ 底部细线", "······ 居中虚点" });
		this.Cmb_PageNumberLeader.Location = new System.Drawing.Point(307, 11);
		this.Cmb_PageNumberLeader.Name = "Cmb_PageNumberLeader";
		this.Cmb_PageNumberLeader.Size = new System.Drawing.Size(104, 28);
		this.Cmb_PageNumberLeader.TabIndex = 6;
		this.Cmb_PageNumberLeader.SelectedIndexChanged += new System.EventHandler(Cmb_PageNumberLeader_SelectedIndexChanged);
		this.Cmb_FontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_FontName.FormattingEnabled = true;
		this.Cmb_FontName.Location = new System.Drawing.Point(46, 7);
		this.Cmb_FontName.Name = "Cmb_FontName";
		this.Cmb_FontName.Size = new System.Drawing.Size(195, 28);
		this.Cmb_FontName.TabIndex = 8;
		this.Cmb_FontName.SelectedIndexChanged += new System.EventHandler(Cmb_FontName_SelectedIndexChanged);
		this.Cmb_FontSize.FormattingEnabled = true;
		this.Cmb_FontSize.Location = new System.Drawing.Point(247, 7);
		this.Cmb_FontSize.Name = "Cmb_FontSize";
		this.Cmb_FontSize.Size = new System.Drawing.Size(81, 28);
		this.Cmb_FontSize.TabIndex = 9;
		this.Cmb_FontSize.SelectedIndexChanged += new System.EventHandler(Cmb_FontSize_SelectedIndexChanged);
		this.Cmb_FontSize.Leave += new System.EventHandler(Cmb_FontSize_Leave);
		this.Cmb_TOCLevelStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_TOCLevelStyle.FormattingEnabled = true;
		this.Cmb_TOCLevelStyle.Location = new System.Drawing.Point(49, 24);
		this.Cmb_TOCLevelStyle.Name = "Cmb_TOCLevelStyle";
		this.Cmb_TOCLevelStyle.Size = new System.Drawing.Size(125, 28);
		this.Cmb_TOCLevelStyle.TabIndex = 12;
		this.Cmb_TOCLevelStyle.SelectedIndexChanged += new System.EventHandler(Cmb_TOCLevelStyle_SelectedIndexChanged);
		this.groupBox1.Controls.Add(this.Chk_SetAllTOCStyle);
		this.groupBox1.Controls.Add(this.Pan_Style);
		this.groupBox1.Controls.Add(this.label4);
		this.groupBox1.Controls.Add(this.Cmb_TOCLevelStyle);
		this.groupBox1.Location = new System.Drawing.Point(9, 90);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(405, 145);
		this.groupBox1.TabIndex = 13;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "字体段落设置";
		this.Chk_SetAllTOCStyle.AutoSize = true;
		this.Chk_SetAllTOCStyle.Location = new System.Drawing.Point(234, 26);
		this.Chk_SetAllTOCStyle.Name = "Chk_SetAllTOCStyle";
		this.Chk_SetAllTOCStyle.Size = new System.Drawing.Size(168, 24);
		this.Chk_SetAllTOCStyle.TabIndex = 16;
		this.Chk_SetAllTOCStyle.Text = "调整所有目录字体段落";
		this.Chk_SetAllTOCStyle.UseVisualStyleBackColor = true;
		this.Chk_SetAllTOCStyle.CheckedChanged += new System.EventHandler(Chk_SetAllTOCStyle_CheckedChanged);
		this.Pan_Style.Controls.Add(this.label9);
		this.Pan_Style.Controls.Add(this.label5);
		this.Pan_Style.Controls.Add(this.label8);
		this.Pan_Style.Controls.Add(this.Cmb_FontSize);
		this.Pan_Style.Controls.Add(this.label7);
		this.Pan_Style.Controls.Add(this.Tog_FontBold);
		this.Pan_Style.Controls.Add(this.Nud_AfterSpace);
		this.Pan_Style.Controls.Add(this.Cmb_FontName);
		this.Pan_Style.Controls.Add(this.Nud_BeforeSpace);
		this.Pan_Style.Controls.Add(this.Tog_FontItalic);
		this.Pan_Style.Controls.Add(this.Nud_LineSpace);
		this.Pan_Style.Controls.Add(this.label6);
		this.Pan_Style.Location = new System.Drawing.Point(3, 56);
		this.Pan_Style.Name = "Pan_Style";
		this.Pan_Style.Size = new System.Drawing.Size(399, 79);
		this.Pan_Style.TabIndex = 15;
		this.label9.AutoSize = true;
		this.label9.Location = new System.Drawing.Point(277, 47);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(37, 20);
		this.label9.TabIndex = 22;
		this.label9.Text = "段后";
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(3, 11);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(37, 20);
		this.label5.TabIndex = 15;
		this.label5.Text = "字体";
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(161, 47);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(37, 20);
		this.label8.TabIndex = 21;
		this.label8.Text = "段前";
		this.label7.AutoSize = true;
		this.label7.Location = new System.Drawing.Point(45, 47);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(37, 20);
		this.label7.TabIndex = 20;
		this.label7.Text = "行距";
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(3, 47);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(37, 20);
		this.label6.TabIndex = 16;
		this.label6.Text = "段落";
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(6, 28);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(37, 20);
		this.label4.TabIndex = 14;
		this.label4.Text = "目录";
		this.Btn_InsertTOC.Location = new System.Drawing.Point(290, 241);
		this.Btn_InsertTOC.Name = "Btn_InsertTOC";
		this.Btn_InsertTOC.Size = new System.Drawing.Size(124, 36);
		this.Btn_InsertTOC.TabIndex = 14;
		this.Btn_InsertTOC.Text = "插入目录";
		this.Btn_InsertTOC.UseVisualStyleBackColor = true;
		this.Btn_InsertTOC.Click += new System.EventHandler(Btn_InsertTOC_Click);
		this.Chk_ReplaceCurrentContents.AutoSize = true;
		this.Chk_ReplaceCurrentContents.Location = new System.Drawing.Point(18, 248);
		this.Chk_ReplaceCurrentContents.Name = "Chk_ReplaceCurrentContents";
		this.Chk_ReplaceCurrentContents.Size = new System.Drawing.Size(140, 24);
		this.Chk_ReplaceCurrentContents.TabIndex = 15;
		this.Chk_ReplaceCurrentContents.Text = "总是替换现有目录";
		this.Chk_ReplaceCurrentContents.UseVisualStyleBackColor = true;
		this.Chk_ReplaceCurrentContents.CheckedChanged += new System.EventHandler(Chk_TOCUsePageNumber_CheckedChanged);
		this.Cmb_GapFromNumberToText.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_GapFromNumberToText.FormattingEnabled = true;
		this.Cmb_GapFromNumberToText.Items.AddRange(new object[3] { "适中", "紧凑", "宽松" });
		this.Cmb_GapFromNumberToText.Location = new System.Drawing.Point(343, 47);
		this.Cmb_GapFromNumberToText.Name = "Cmb_GapFromNumberToText";
		this.Cmb_GapFromNumberToText.Size = new System.Drawing.Size(68, 28);
		this.Cmb_GapFromNumberToText.TabIndex = 16;
		this.Cmb_GapFromNumberToText.SelectedIndexChanged += new System.EventHandler(Cmb_PageNumberLeader_SelectedIndexChanged);
		this.Chk_TryAlignNumber.AutoSize = true;
		this.Chk_TryAlignNumber.Location = new System.Drawing.Point(165, 248);
		this.Chk_TryAlignNumber.Name = "Chk_TryAlignNumber";
		this.Chk_TryAlignNumber.Size = new System.Drawing.Size(112, 24);
		this.Chk_TryAlignNumber.TabIndex = 17;
		this.Chk_TryAlignNumber.Text = "尝试对齐页码";
		this.Chk_TryAlignNumber.UseVisualStyleBackColor = true;
		this.Chk_TryAlignNumber.CheckedChanged += new System.EventHandler(Chk_TOCUsePageNumber_CheckedChanged);
		this.Tog_FontBold.BackColor = System.Drawing.Color.AliceBlue;
		this.Tog_FontBold.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
		this.Tog_FontBold.Location = new System.Drawing.Point(331, 7);
		this.Tog_FontBold.Name = "Tog_FontBold";
		this.Tog_FontBold.Pressed = false;
		this.Tog_FontBold.Size = new System.Drawing.Size(28, 28);
		this.Tog_FontBold.TabIndex = 10;
		this.Tog_FontBold.Text = "B";
		this.Tog_FontBold.UseVisualStyleBackColor = false;
		this.Tog_FontBold.Click += new System.EventHandler(Tog_FontBold_Click);
		this.Nud_AfterSpace.DecimalPlaces = 2;
		this.Nud_AfterSpace.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_AfterSpace.Label = "行";
		this.Nud_AfterSpace.Location = new System.Drawing.Point(316, 44);
		this.Nud_AfterSpace.Name = "Nud_AfterSpace";
		this.Nud_AfterSpace.Size = new System.Drawing.Size(74, 26);
		this.Nud_AfterSpace.TabIndex = 19;
		this.Nud_AfterSpace.ValueChanged += new System.EventHandler(Nud_LineSpace_ValueChanged);
		this.Nud_BeforeSpace.DecimalPlaces = 2;
		this.Nud_BeforeSpace.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_BeforeSpace.Label = "行";
		this.Nud_BeforeSpace.Location = new System.Drawing.Point(200, 44);
		this.Nud_BeforeSpace.Name = "Nud_BeforeSpace";
		this.Nud_BeforeSpace.Size = new System.Drawing.Size(74, 26);
		this.Nud_BeforeSpace.TabIndex = 18;
		this.Nud_BeforeSpace.ValueChanged += new System.EventHandler(Nud_LineSpace_ValueChanged);
		this.Tog_FontItalic.BackColor = System.Drawing.Color.AliceBlue;
		this.Tog_FontItalic.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, 134);
		this.Tog_FontItalic.Location = new System.Drawing.Point(362, 7);
		this.Tog_FontItalic.Name = "Tog_FontItalic";
		this.Tog_FontItalic.Pressed = false;
		this.Tog_FontItalic.Size = new System.Drawing.Size(28, 28);
		this.Tog_FontItalic.TabIndex = 11;
		this.Tog_FontItalic.Text = "I";
		this.Tog_FontItalic.UseVisualStyleBackColor = false;
		this.Tog_FontItalic.Click += new System.EventHandler(Tog_FontBold_Click);
		this.Nud_LineSpace.DecimalPlaces = 2;
		this.Nud_LineSpace.Increment = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_LineSpace.Label = "行";
		this.Nud_LineSpace.Location = new System.Drawing.Point(84, 44);
		this.Nud_LineSpace.Minimum = new decimal(new int[4] { 1, 0, 0, 65536 });
		this.Nud_LineSpace.Name = "Nud_LineSpace";
		this.Nud_LineSpace.Size = new System.Drawing.Size(74, 26);
		this.Nud_LineSpace.TabIndex = 17;
		this.Nud_LineSpace.Value = new decimal(new int[4] { 1, 0, 0, 0 });
		this.Nud_LineSpace.ValueChanged += new System.EventHandler(Nud_LineSpace_ValueChanged);
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.Chk_TryAlignNumber);
		base.Controls.Add(this.Cmb_GapFromNumberToText);
		base.Controls.Add(this.Chk_ReplaceCurrentContents);
		base.Controls.Add(this.Btn_InsertTOC);
		base.Controls.Add(this.groupBox1);
		base.Controls.Add(this.Cmb_PageNumberLeader);
		base.Controls.Add(this.Cmb_TOCIndent);
		base.Controls.Add(this.Cmb_TOCLevel);
		base.Controls.Add(this.Chk_TOCUsePageNumber);
		base.Controls.Add(this.label2);
		base.Controls.Add(this.label1);
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		base.Name = "TOCSet";
		base.Size = new System.Drawing.Size(420, 285);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.Pan_Style.ResumeLayout(false);
		this.Pan_Style.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Nud_AfterSpace).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_BeforeSpace).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Nud_LineSpace).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}