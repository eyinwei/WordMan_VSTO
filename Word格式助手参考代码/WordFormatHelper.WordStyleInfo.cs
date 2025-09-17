// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.WordStyleInfo
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using WordFormatHelper;

[Serializable]
public class WordStyleInfo
{
	public struct StyleParaValues
	{
		public string ChnFontName;

		public string EngFontName;

		public string FontSize;

		public Color FontColor;

		public bool Bold;

		public bool Italic;

		public bool Underline;

		public string HAlignment;

		public string LeftIndent;

		public string RightIndent;

		public string FirstLineIndent;

		public string LineSpace;

		public string SpaceBefore;

		public string SpaceAfter;

		public bool BreakBefore;

		public int NumberStyle;

		public string NumberFormat;

		public StyleParaValues()
		{
			ChnFontName = "宋体";
			EngFontName = "宋体";
			FontSize = "五号";
			FontColor = Color.Black;
			Bold = false;
			Italic = false;
			Underline = false;
			HAlignment = "左对齐";
			LeftIndent = "0.00 厘米";
			RightIndent = "0.00 厘米";
			FirstLineIndent = "0.00 磅";
			LineSpace = "单倍行距";
			SpaceBefore = "0.00 行";
			SpaceAfter = "0.00 行";
			BreakBefore = false;
			NumberStyle = -1;
			NumberFormat = "";
		}
	}

	private WdBuiltinStyle buildInName;

	public static readonly string[] HAlignments = new string[5] { "左对齐", "中对齐", "右对齐", "两端对齐", "分散对齐" };

	public static readonly string[] LineSpacingValues = new string[3] { "单倍行距", "1.5倍行距", "双倍行距" };

	public static readonly string[] ParagraphSpaceValues = new string[10] { "自动", "0 行", "0.5 行", "1 行", "1.25 行", "1.5 行", "1.8 行", "2 行", "2.5 行", "3 行" };

	public static readonly List<string> FontSizeList = new List<string>(16)
	{
		"初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四",
		"五号", "小五", "六号", "小六", "七号", "八号"
	};

	public static readonly List<float> FontSizeValueList = new List<float>(16)
	{
		42f, 36f, 26f, 24f, 22f, 18f, 16f, 15f, 14f, 12f,
		10.5f, 9f, 7.5f, 6.5f, 5.5f, 5f
	};

	public static readonly List<WdListNumberStyle> ListNumberStyles = new List<WdListNumberStyle>(10)
	{
		WdListNumberStyle.wdListNumberStyleArabic,
		WdListNumberStyle.wdListNumberStyleLegalLZ,
		WdListNumberStyle.wdListNumberStyleUppercaseLetter,
		WdListNumberStyle.wdListNumberStyleLowercaseLetter,
		WdListNumberStyle.wdListNumberStyleUppercaseRoman,
		WdListNumberStyle.wdListNumberStyleLowercaseRoman,
		WdListNumberStyle.wdListNumberStyleSimpChinNum1,
		WdListNumberStyle.wdListNumberStyleSimpChinNum2,
		WdListNumberStyle.wdListNumberStyleZodiac1,
		WdListNumberStyle.wdListNumberStyleLegal
	};

	public static readonly List<string> ListNumberStyleName = new List<string>(11)
	{
		"无编号", "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,貳,叁...", "甲,乙,丙...",
		"正规编号"
	};

	public static readonly List<WdBuiltinStyle> BuildInStyleNames = new List<WdBuiltinStyle>(87)
	{
		WdBuiltinStyle.wdStyleNormal,
		WdBuiltinStyle.wdStyleHeading1,
		WdBuiltinStyle.wdStyleHeading2,
		WdBuiltinStyle.wdStyleHeading3,
		WdBuiltinStyle.wdStyleHeading4,
		WdBuiltinStyle.wdStyleHeading5,
		WdBuiltinStyle.wdStyleHeading6,
		WdBuiltinStyle.wdStyleHeading7,
		WdBuiltinStyle.wdStyleHeading8,
		WdBuiltinStyle.wdStyleHeading9,
		WdBuiltinStyle.wdStyleIndex1,
		WdBuiltinStyle.wdStyleIndex2,
		WdBuiltinStyle.wdStyleIndex3,
		WdBuiltinStyle.wdStyleIndex4,
		WdBuiltinStyle.wdStyleIndex5,
		WdBuiltinStyle.wdStyleIndex6,
		WdBuiltinStyle.wdStyleIndex7,
		WdBuiltinStyle.wdStyleIndex8,
		WdBuiltinStyle.wdStyleIndex9,
		WdBuiltinStyle.wdStyleTOC1,
		WdBuiltinStyle.wdStyleTOC2,
		WdBuiltinStyle.wdStyleTOC3,
		WdBuiltinStyle.wdStyleTOC4,
		WdBuiltinStyle.wdStyleTOC5,
		WdBuiltinStyle.wdStyleTOC6,
		WdBuiltinStyle.wdStyleTOC7,
		WdBuiltinStyle.wdStyleTOC8,
		WdBuiltinStyle.wdStyleTOC9,
		WdBuiltinStyle.wdStyleNormalIndent,
		WdBuiltinStyle.wdStyleFootnoteText,
		WdBuiltinStyle.wdStyleCommentText,
		WdBuiltinStyle.wdStyleHeader,
		WdBuiltinStyle.wdStyleFooter,
		WdBuiltinStyle.wdStyleIndexHeading,
		WdBuiltinStyle.wdStyleCaption,
		WdBuiltinStyle.wdStyleTableOfFigures,
		WdBuiltinStyle.wdStyleEnvelopeAddress,
		WdBuiltinStyle.wdStyleEnvelopeReturn,
		WdBuiltinStyle.wdStyleEndnoteText,
		WdBuiltinStyle.wdStyleTableOfAuthorities,
		WdBuiltinStyle.wdStyleMacroText,
		WdBuiltinStyle.wdStyleTOAHeading,
		WdBuiltinStyle.wdStyleList,
		WdBuiltinStyle.wdStyleListBullet,
		WdBuiltinStyle.wdStyleListNumber,
		WdBuiltinStyle.wdStyleList2,
		WdBuiltinStyle.wdStyleList3,
		WdBuiltinStyle.wdStyleList4,
		WdBuiltinStyle.wdStyleList5,
		WdBuiltinStyle.wdStyleListBullet2,
		WdBuiltinStyle.wdStyleListBullet3,
		WdBuiltinStyle.wdStyleListBullet4,
		WdBuiltinStyle.wdStyleListBullet5,
		WdBuiltinStyle.wdStyleListNumber2,
		WdBuiltinStyle.wdStyleListNumber3,
		WdBuiltinStyle.wdStyleListNumber4,
		WdBuiltinStyle.wdStyleListNumber5,
		WdBuiltinStyle.wdStyleTitle,
		WdBuiltinStyle.wdStyleClosing,
		WdBuiltinStyle.wdStyleSignature,
		WdBuiltinStyle.wdStyleBodyText,
		WdBuiltinStyle.wdStyleBodyTextIndent,
		WdBuiltinStyle.wdStyleListContinue,
		WdBuiltinStyle.wdStyleListContinue2,
		WdBuiltinStyle.wdStyleListContinue3,
		WdBuiltinStyle.wdStyleListContinue4,
		WdBuiltinStyle.wdStyleListContinue5,
		WdBuiltinStyle.wdStyleMessageHeader,
		WdBuiltinStyle.wdStyleSubtitle,
		WdBuiltinStyle.wdStyleSalutation,
		WdBuiltinStyle.wdStyleDate,
		WdBuiltinStyle.wdStyleBodyTextFirstIndent,
		WdBuiltinStyle.wdStyleBodyTextFirstIndent2,
		WdBuiltinStyle.wdStyleNoteHeading,
		WdBuiltinStyle.wdStyleBodyText2,
		WdBuiltinStyle.wdStyleBodyText3,
		WdBuiltinStyle.wdStyleBodyTextIndent2,
		WdBuiltinStyle.wdStyleBodyTextIndent3,
		WdBuiltinStyle.wdStyleBlockQuotation,
		WdBuiltinStyle.wdStyleNavPane,
		WdBuiltinStyle.wdStylePlainText,
		WdBuiltinStyle.wdStyleHtmlNormal,
		WdBuiltinStyle.wdStyleHtmlAddress,
		WdBuiltinStyle.wdStyleHtmlPre,
		WdBuiltinStyle.wdStyleNormalObject,
		WdBuiltinStyle.wdStyleListParagraph,
		WdBuiltinStyle.wdStyleQuote
	};

	[Browsable(false)]
	[XmlIgnore]
	public bool BuildInStyle { get; private set; }

	[Browsable(false)]
	[XmlIgnore]
	public WdBuiltinStyle BuildInStyleName
	{
		get
		{
			if (BuildInStyle)
			{
				return buildInName;
			}
			return (WdBuiltinStyle)0;
		}
		private set
		{
			try
			{
				buildInName = value;
			}
			catch
			{
				buildInName = (WdBuiltinStyle)0;
			}
		}
	}

	[Browsable(false)]
	public int BuildInStyleNameInt
	{
		get
		{
			return (int)buildInName;
		}
		set
		{
			if (value == 0)
			{
				BuildInStyle = false;
				buildInName = (WdBuiltinStyle)0;
			}
			else
			{
				BuildInStyle = true;
				BuildInStyleName = (WdBuiltinStyle)value;
			}
		}
	}

	public string StyleName { get; set; }

	public string ChnFontName { get; set; }

	public string EngFontName { get; set; }

	public string FontSize { get; set; }

	[XmlIgnore]
	public Color FontColor { get; set; }

	[Browsable(false)]
	public string FontColorValue
	{
		get
		{
			return "#" + FontColor.A.ToString("X2") + FontColor.R.ToString("X2") + FontColor.G.ToString("X2") + FontColor.B.ToString("X2");
		}
		set
		{
			try
			{
				if (!Regex.IsMatch(value, "^#([0-9A-Fa-f]{8})$"))
				{
					throw new Exception();
				}
				int alpha = int.Parse(value.Substring(1, 2), NumberStyles.HexNumber);
				int red = int.Parse(value.Substring(3, 2), NumberStyles.HexNumber);
				int green = int.Parse(value.Substring(5, 2), NumberStyles.HexNumber);
				int blue = int.Parse(value.Substring(7, 2), NumberStyles.HexNumber);
				FontColor = Color.FromArgb(alpha, red, green, blue);
			}
			catch
			{
				FontColor = Color.Black;
			}
		}
	}

	public bool Bold { get; set; }

	public bool Italic { get; set; }

	public bool Underline { get; set; }

	public string LeftIndent { get; set; }

	public string RightIndent { get; set; }

	[Browsable(false)]
	public string FirstLineIndent { get; set; }

	public string LineSpace { get; set; }

	public string SpaceBefore { get; set; }

	public string SpaceAfter { get; set; }

	public bool BreakBefore { get; set; }

	public string HAlignment { get; set; }

	[Browsable(false)]
	public int NumberStyle { get; set; }

	[Browsable(false)]
	public string NumberFormat { get; set; }

	public WordStyleInfo(Style style, [Optional] WdBuiltinStyle bN)
	{
		if (bN != 0)
		{
			buildInName = bN;
		}
		StyleName = style.NameLocal;
		BuildInStyle = style.BuiltIn;
		if (style.Font.NameFarEast.StartsWith("+"))
		{
			string V_2 = style.Font.NameFarEast;
			if (!(V_2 == "+中文标题"))
			{
				if (V_2 == "+中文正文")
				{
					ChnFontName = (style.Parent as Document).DocumentTheme.ThemeFontScheme.MinorFont.Item(MsoFontLanguageIndex.msoThemeLatin).Name;
				}
			}
			else
			{
				ChnFontName = (style.Parent as Document).DocumentTheme.ThemeFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin).Name;
			}
		}
		else
		{
			ChnFontName = style.Font.NameFarEast;
		}
		EngFontName = style.Font.Name;
		int num = FontSizeValueList.IndexOf(style.Font.Size);
		if (num == -1)
		{
			FontSize = style.Font.Size.ToString("0.0 磅");
		}
		else
		{
			FontSize = FontSizeList[num];
		}
		if (style.Font.TextColor.ObjectThemeColor != WdThemeColorIndex.wdNotThemeColor)
		{
			Color color = ThisAddIn.ColorFormInt(Globals.ThisAddIn.Application.ActiveDocument.DocumentTheme.ThemeColorScheme.Colors(GetThemeColorIndex(style.Font.TextColor.ObjectThemeColor)).RGB);
			byte r = color.R;
			byte g = color.G;
			byte b = color.B;
			if (style.Font.TextColor.TintAndShade >= 0f)
			{
				r += (byte)((float)(255 - r) * style.Font.TextColor.TintAndShade);
				g += (byte)((float)(255 - g) * style.Font.TextColor.TintAndShade);
				b += (byte)((float)(255 - b) * style.Font.TextColor.TintAndShade);
			}
			else
			{
				r = (byte)((float)(int)r * (1f + style.Font.TextColor.TintAndShade));
				g = (byte)((float)(int)g * (1f + style.Font.TextColor.TintAndShade));
				b = (byte)((float)(int)b * (1f + style.Font.TextColor.TintAndShade));
			}
			FontColor = Color.FromArgb(255, r, g, b);
		}
		else
		{
			FontColor = ThisAddIn.ColorFormInt(style.Font.TextColor.RGB);
		}
		Bold = style.Font.Bold == -1;
		Italic = style.Font.Italic == -1;
		Underline = style.Font.Underline != WdUnderline.wdUnderlineNone;
		LeftIndent = (style.ParagraphFormat.LeftIndent * 2.54f / 72f).ToString("0.00 厘米");
		RightIndent = (style.ParagraphFormat.RightIndent * 2.54f / 72f).ToString("0.00 厘米");
		switch (style.ParagraphFormat.LineSpacingRule)
		{
		case WdLineSpacing.wdLineSpaceSingle:
			LineSpace = "单倍行距";
			break;
		case WdLineSpacing.wdLineSpaceDouble:
			LineSpace = "双倍行距";
			break;
		case WdLineSpacing.wdLineSpace1pt5:
			LineSpace = "1.5倍行距";
			break;
		case WdLineSpacing.wdLineSpaceAtLeast:
		case WdLineSpacing.wdLineSpaceExactly:
			LineSpace = (style.ParagraphFormat.LineSpacing * 2.54f / 72f).ToString("0.00 厘米");
			break;
		case WdLineSpacing.wdLineSpaceMultiple:
			LineSpace = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.LineSpacing).ToString("0.00 行");
			break;
		}
		SpaceBefore = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.SpaceBefore).ToString("0.00 行");
		SpaceAfter = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.SpaceAfter).ToString("0.00 行");
		FirstLineIndent = style.ParagraphFormat.FirstLineIndent.ToString("0.00 磅");
		HAlignment = style.ParagraphFormat.Alignment switch
		{
			WdParagraphAlignment.wdAlignParagraphLeft => "左对齐", 
			WdParagraphAlignment.wdAlignParagraphCenter => "中对齐", 
			WdParagraphAlignment.wdAlignParagraphRight => "右对齐", 
			WdParagraphAlignment.wdAlignParagraphJustify => "两端对齐", 
			WdParagraphAlignment.wdAlignParagraphDistribute => "分散对齐", 
			_ => "左对齐", 
		};
		BreakBefore = style.ParagraphFormat.PageBreakBefore == -1;
		NumberStyle = -1;
		NumberFormat = "";
	}

	public WordStyleInfo(string styleName, StyleParaValues para = default(StyleParaValues))
	{
		StyleName = styleName;
		buildInName = (WdBuiltinStyle)0;
		BuildInStyle = false;
		SetStyleValue(para);
	}

	public WordStyleInfo(WdBuiltinStyle builtinStyle, StyleParaValues para = default(StyleParaValues))
	{
		if (!BuildInStyleNames.Contains(builtinStyle))
		{
			throw new InvalidEnumArgumentException("builtinStyle");
		}
		Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
		object Index = builtinStyle;
		StyleName = styles[ref Index].NameLocal;
		buildInName = builtinStyle;
		BuildInStyle = true;
		SetStyleValue(para);
	}

	public WordStyleInfo()
	{
	}

	private MsoThemeColorSchemeIndex GetThemeColorIndex(WdThemeColorIndex index)
	{
		return index switch
		{
			WdThemeColorIndex.wdThemeColorMainLight1 => MsoThemeColorSchemeIndex.msoThemeLight1, 
			WdThemeColorIndex.wdThemeColorMainDark2 => MsoThemeColorSchemeIndex.msoThemeDark2, 
			WdThemeColorIndex.wdThemeColorMainLight2 => MsoThemeColorSchemeIndex.msoThemeLight2, 
			WdThemeColorIndex.wdThemeColorAccent1 => MsoThemeColorSchemeIndex.msoThemeAccent1, 
			WdThemeColorIndex.wdThemeColorAccent2 => MsoThemeColorSchemeIndex.msoThemeAccent2, 
			WdThemeColorIndex.wdThemeColorAccent3 => MsoThemeColorSchemeIndex.msoThemeAccent3, 
			WdThemeColorIndex.wdThemeColorAccent4 => MsoThemeColorSchemeIndex.msoThemeAccent4, 
			WdThemeColorIndex.wdThemeColorAccent5 => MsoThemeColorSchemeIndex.msoThemeAccent5, 
			WdThemeColorIndex.wdThemeColorAccent6 => MsoThemeColorSchemeIndex.msoThemeAccent6, 
			WdThemeColorIndex.wdThemeColorHyperlink => MsoThemeColorSchemeIndex.msoThemeHyperlink, 
			WdThemeColorIndex.wdThemeColorHyperlinkFollowed => MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink, 
			_ => MsoThemeColorSchemeIndex.msoThemeDark1, 
		};
	}

	public bool SetStyle(Document docForApply)
	{
		try
		{
			Style style;
			if (BuildInStyle)
			{
				Styles styles = docForApply.Styles;
				object Index = buildInName;
				style = styles[ref Index];
			}
			else
			{
				try
				{
					Styles styles2 = docForApply.Styles;
					object Index = StyleName;
					style = styles2[ref Index];
				}
				catch
				{
					Styles styles3 = docForApply.Styles;
					string styleName = StyleName;
					object Index = WdStyleType.wdStyleTypeParagraph;
					style = styles3.Add(styleName, ref Index);
				}
			}
			if (buildInName != WdBuiltinStyle.wdStyleNormal)
			{
				Style style2 = style;
				object Index = "";
				style2.set_BaseStyle(ref Index);
				Style style3 = style;
				Index = WdBuiltinStyle.wdStyleNormal;
				style3.set_NextParagraphStyle(ref Index);
			}
			style.Font.NameFarEast = ChnFontName;
			style.Font.Name = EngFontName;
			int num = FontSizeList.IndexOf(FontSize);
			string s = FontSize.TrimEnd(' ', '磅');
			if (num == -1)
			{
				style.Font.Size = float.Parse(s);
				if (buildInName == WdBuiltinStyle.wdStyleNormal)
				{
					Globals.ThisAddIn.SetGrid(docForApply, float.Parse(s));
				}
			}
			else
			{
				style.Font.Size = FontSizeValueList[num];
				if (buildInName == WdBuiltinStyle.wdStyleNormal)
				{
					Globals.ThisAddIn.SetGrid(docForApply, FontSizeValueList[num]);
				}
			}
			style.Font.TextColor.RGB = ThisAddIn.RGB(FontColor.R, FontColor.G, FontColor.B);
			style.Font.Bold = (Bold ? (-1) : 0);
			style.Font.Italic = (Italic ? (-1) : 0);
			if (Underline)
			{
				style.Font.Underline = WdUnderline.wdUnderlineSingle;
			}
			else
			{
				style.Font.Underline = WdUnderline.wdUnderlineNone;
			}
			if (NumberStyle != -1)
			{
				if (NumberStyle > 0)
				{
					ListTemplates listTemplates = Globals.ThisAddIn.Application.ListGalleries[WdListGalleryType.wdNumberGallery].ListTemplates;
					object Index = 1;
					ListTemplate listTemplate = listTemplates[ref Index];
					listTemplate.ListLevels[1].NumberStyle = ListNumberStyles[NumberStyle - 1];
					listTemplate.ListLevels[1].NumberFormat = NumberFormat;
					listTemplate.ListLevels[1].TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
					Style style4 = style;
					Index = Type.Missing;
					style4.LinkToListTemplate(listTemplate, ref Index);
				}
				else
				{
					Style style5 = style;
					object Index = Type.Missing;
					style5.LinkToListTemplate(null, ref Index);
				}
			}
			ParagraphFormat paragraphFormat = (ParagraphFormat)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209F4-0000-0000-C000-000000000046")));
			ParagraphFormat paragraphFormat2 = paragraphFormat;
			paragraphFormat2.Alignment = HAlignment switch
			{
				"左对齐" => WdParagraphAlignment.wdAlignParagraphLeft, 
				"中对齐" => WdParagraphAlignment.wdAlignParagraphCenter, 
				"右对齐" => WdParagraphAlignment.wdAlignParagraphRight, 
				"两端对齐" => WdParagraphAlignment.wdAlignParagraphJustify, 
				"分散对齐" => WdParagraphAlignment.wdAlignParagraphDistribute, 
				_ => WdParagraphAlignment.wdAlignParagraphLeft, 
			};
			ParagraphFormat paragraphFormat3 = paragraphFormat;
			if (LeftIndent.EndsWith("厘米"))
			{
				s = LeftIndent.TrimEnd(' ', '厘', '米');
				paragraphFormat3.LeftIndent = float.Parse(s) * 72f / 2.54f;
			}
			else
			{
				s = LeftIndent.TrimEnd(' ', '磅');
				paragraphFormat3.LeftIndent = float.Parse(s);
			}
			if (RightIndent.EndsWith("厘米"))
			{
				s = RightIndent.TrimEnd(' ', '厘', '米');
				paragraphFormat3.RightIndent = float.Parse(s) * 72f / 2.54f;
			}
			else
			{
				s = RightIndent.TrimEnd(' ', '磅');
				paragraphFormat3.RightIndent = float.Parse(s);
			}
			if (FirstLineIndent.EndsWith("字符"))
			{
				s = FirstLineIndent.TrimEnd(' ', '字', '符');
				paragraphFormat3.IndentFirstLineCharWidth(short.Parse(s));
			}
			else
			{
				s = FirstLineIndent.TrimEnd(' ', '磅');
				paragraphFormat3.FirstLineIndent = float.Parse(s);
			}
			switch (LineSpace)
			{
			case "单倍行距":
				paragraphFormat3.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				paragraphFormat3.Space1();
				break;
			case "1.5倍行距":
				paragraphFormat3.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
				paragraphFormat3.Space15();
				break;
			case "双倍行距":
				paragraphFormat3.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
				paragraphFormat3.Space2();
				break;
			default:
				if (LineSpace.EndsWith("行"))
				{
					s = LineSpace.TrimEnd(' ', '行');
					paragraphFormat3.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
					paragraphFormat3.LineSpacing = Globals.ThisAddIn.Application.LinesToPoints(float.Parse(s));
				}
				else
				{
					s = LineSpace.TrimEnd(' ', '磅');
					paragraphFormat3.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
					paragraphFormat3.LineSpacing = float.Parse(s);
				}
				break;
			}
			if (SpaceBefore == "自动")
			{
				paragraphFormat3.SpaceBeforeAuto = -1;
			}
			else if (SpaceBefore.EndsWith("行"))
			{
				s = SpaceBefore.TrimEnd(' ', '行');
				paragraphFormat3.SpaceBeforeAuto = 0;
				paragraphFormat3.SpaceBefore = Globals.ThisAddIn.Application.LinesToPoints(float.Parse(s));
			}
			else
			{
				s = SpaceBefore.TrimEnd(' ', '磅');
				paragraphFormat3.SpaceBeforeAuto = 0;
				paragraphFormat3.SpaceBefore = float.Parse(s);
			}
			if (SpaceAfter == "自动")
			{
				paragraphFormat3.SpaceAfterAuto = -1;
			}
			else if (SpaceAfter.EndsWith("行"))
			{
				s = SpaceAfter.TrimEnd(' ', '行');
				paragraphFormat3.SpaceAfterAuto = 0;
				paragraphFormat3.SpaceAfter = Globals.ThisAddIn.Application.LinesToPoints(float.Parse(s));
			}
			else
			{
				s = SpaceAfter.TrimEnd(' ', '磅');
				paragraphFormat3.SpaceAfterAuto = 0;
				paragraphFormat3.SpaceAfter = float.Parse(s);
			}
			paragraphFormat3.PageBreakBefore = (BreakBefore ? (-1) : 0);
			style.ParagraphFormat = paragraphFormat3;
			style.QuickStyle = true;
			return true;
		}
		catch
		{
			return false;
		}
	}

	public void SetStyleValue(StyleParaValues para)
	{
		ChnFontName = para.ChnFontName;
		EngFontName = para.EngFontName;
		FontSize = para.FontSize;
		Bold = para.Bold;
		Italic = para.Italic;
		Underline = para.Underline;
		FontColor = para.FontColor;
		HAlignment = para.HAlignment;
		LeftIndent = para.LeftIndent;
		RightIndent = para.RightIndent;
		FirstLineIndent = para.FirstLineIndent;
		LineSpace = para.LineSpace;
		SpaceBefore = para.SpaceBefore;
		SpaceAfter = para.SpaceAfter;
		BreakBefore = para.BreakBefore;
		NumberStyle = para.NumberStyle;
		NumberFormat = para.NumberFormat;
	}

	public string GetDescription(out System.Drawing.Font font)
	{
		string text = "";
		FontStyle fontStyle = FontStyle.Regular;
		text = text + "中文字体：" + ChnFontName + "； 西文字体：" + EngFontName + "；大小：" + FontSize + "；";
		if (Bold)
		{
			text += "粗体；";
			fontStyle |= FontStyle.Bold;
		}
		if (Italic)
		{
			text += "斜体；";
			fontStyle |= FontStyle.Italic;
		}
		if (Underline)
		{
			text += "下划线；";
			fontStyle |= FontStyle.Underline;
		}
		text = text + "段落" + HAlignment + "；左缩进：" + LeftIndent + "；右缩进：" + RightIndent + "；";
		text = text + "段落行距：" + LineSpace + "；段前：" + SpaceBefore + "；段后：" + SpaceAfter + "；";
		if (BreakBefore)
		{
			text += "段前分行；";
		}
		font = new System.Drawing.Font(new FontFamily(ChnFontName), 10.5f, fontStyle);
		return text;
	}
}
