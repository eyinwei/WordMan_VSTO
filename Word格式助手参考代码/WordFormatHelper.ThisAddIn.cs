// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.ThisAddIn
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using EnocheastyBarCode.BarCode;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using WeChatOcr;
using WordFormatHelper;

[StartupObject(0)]
[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
public sealed class ThisAddIn : AddInBase
{
	internal enum HeaderFooterTextType
	{
		None,
		Left,
		Center,
		Right,
		LeftRight,
		LeftCenter,
		CenterRight,
		All
	}

	internal struct TableSettings
	{
		internal int Rows;

		internal int Columns;

		internal string FontName;

		internal bool FixRowHeight;

		internal float FontSize;

		internal Color FontColor;

		internal string TableTitle;

		internal bool CaptionLab;

		internal string CaptionTitle;

		internal int CaptionNumberStyle;

		internal bool CaptionIncludeHeadings;

		internal int HeadingsLevel;

		internal int LinkChar;

		internal int FillType;

		internal int TextureStyle;

		internal Color BackgrounColor;

		internal int OuterLineType;

		internal int InnerLineType;

		internal int TitleRowLineType;

		internal int OuterLineWidth;

		internal int InnerLineWidth;

		internal int TitleRowLineWidth;

		internal Color OuterLineColor;

		internal Color InnerLineColor;

		internal Color TitleRowLineColor;

		public TableSettings()
		{
			Rows = 5;
			Columns = 5;
			FontName = "宋体";
			FixRowHeight = false;
			FontSize = 10.5f;
			FontColor = Color.Black;
			TableTitle = "表格标题";
			CaptionLab = false;
			CaptionTitle = null;
			CaptionNumberStyle = 0;
			CaptionIncludeHeadings = false;
			HeadingsLevel = 0;
			LinkChar = 0;
			FillType = 0;
			TextureStyle = 0;
			BackgrounColor = Color.LightGray;
			OuterLineType = 0;
			InnerLineType = 0;
			TitleRowLineType = 0;
			OuterLineWidth = 5;
			InnerLineWidth = 2;
			TitleRowLineWidth = 2;
			OuterLineColor = Color.Black;
			InnerLineColor = Color.Black;
			TitleRowLineColor = Color.Black;
		}
	}

	internal struct TextFormatSet
	{
		public int SetType;

		public bool IgnoreNumberDot;

		public int RemoveSpaceType;

		internal bool RemoveBrackets;

		internal bool FullWidthBracket;

		public TextFormatSet()
		{
			SetType = 7;
			IgnoreNumberDot = false;
			RemoveSpaceType = 0;
			RemoveBrackets = false;
			FullWidthBracket = false;
		}
	}

	internal readonly MsoLineStyle[] LineTypes = (MsoLineStyle[])Enum.GetValues(typeof(MsoLineStyle));

	internal readonly List<WdListNumberStyle> LevelNumStyle = new List<WdListNumberStyle>(10)
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

	internal readonly List<WdCaptionNumberStyle> CaptionNumStyle = new List<WdCaptionNumberStyle>(9)
	{
		WdCaptionNumberStyle.wdCaptionNumberStyleArabic,
		WdCaptionNumberStyle.wdCaptionNumberStyleArabicFullWidth,
		WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseLetter,
		WdCaptionNumberStyle.wdCaptionNumberStyleLowercaseLetter,
		WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman,
		WdCaptionNumberStyle.wdCaptionNumberStyleLowercaseRoman,
		WdCaptionNumberStyle.wdCaptionNumberStyleSimpChinNum2,
		WdCaptionNumberStyle.wdCaptionNumberStyleTradChinNum2,
		WdCaptionNumberStyle.wdCaptionNumberStyleZodiac1
	};

	internal readonly List<WdRowAlignment> tableAlignmentType = new List<WdRowAlignment>(3)
	{
		WdRowAlignment.wdAlignRowLeft,
		WdRowAlignment.wdAlignRowRight,
		WdRowAlignment.wdAlignRowCenter
	};

	internal readonly List<WdParagraphAlignment> inlineShapeAlignmentType = new List<WdParagraphAlignment>(3)
	{
		WdParagraphAlignment.wdAlignParagraphLeft,
		WdParagraphAlignment.wdAlignParagraphRight,
		WdParagraphAlignment.wdAlignParagraphCenter
	};

	internal readonly List<WdShapePosition> shapeAlignmentType = new List<WdShapePosition>(3)
	{
		WdShapePosition.wdShapeLeft,
		WdShapePosition.wdShapeRight,
		WdShapePosition.wdShapeCenter
	};

	internal readonly List<WdLineStyle> lineStyle = new List<WdLineStyle>(4)
	{
		WdLineStyle.wdLineStyleSingle,
		WdLineStyle.wdLineStyleDouble,
		WdLineStyle.wdLineStyleThinThickMedGap,
		WdLineStyle.wdLineStyleThickThinMedGap
	};

	internal readonly WdLineWidth[] tLineWidth = (WdLineWidth[])Enum.GetValues(typeof(WdLineWidth));

	internal readonly List<WdTextureIndex> textureIndex = new List<WdTextureIndex>(12)
	{
		WdTextureIndex.wdTextureDiagonalCross,
		WdTextureIndex.wdTextureCross,
		WdTextureIndex.wdTextureDiagonalDown,
		WdTextureIndex.wdTextureDiagonalUp,
		WdTextureIndex.wdTextureVertical,
		WdTextureIndex.wdTextureHorizontal,
		WdTextureIndex.wdTextureDarkDiagonalCross,
		WdTextureIndex.wdTextureDarkCross,
		WdTextureIndex.wdTextureDarkDiagonalDown,
		WdTextureIndex.wdTextureDarkDiagonalUp,
		WdTextureIndex.wdTextureDarkVertical,
		WdTextureIndex.wdTextureDarkHorizontal
	};

	internal readonly WordFormatHelperDefault defaultValue = new WordFormatHelperDefault();

	internal static TableSettings tableSettings = new TableSettings();

	internal static TextFormatSet textFormatSet = new TextFormatSet();

	internal static FixFormatPainterSetting formatPainter;

	public const int GWL_HWNDPARENT = -8;

	public const float Undefined = 9999999f;

	internal static ImageOcr OCR_Engine = null;

	internal CustomTaskPaneCollection CustomTaskPanes;

	internal SmartTagCollection VstoSmartTags;

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	private object missing = Type.Missing;

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	internal Microsoft.Office.Interop.Word.Application Application;

	[DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SetWindowLongPtrW", ExactSpelling = true, SetLastError = true)]
	public static extern nint SetWindowLongPtrImp(IntPtr hWnd, int nIndex, long dwNewLong);

	private void ThisAddIn_Startup(object sender, EventArgs e)
	{
		if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\FormatPainterSettings.xml"))
		{
			formatPainter = FixFormatPainterSetting.FromXmlFile(AppDomain.CurrentDomain.BaseDirectory + "\\FormatPainterSettings.xml");
			return;
		}
		formatPainter = new FixFormatPainterSetting
		{
			CurrentID = 0
		};
		FixFormatPainterSetting.FixFormat item = new FixFormatPainterSetting.FixFormat
		{
			Id = 0,
			StyleName = "默认",
			Discription = "默认格式刷样式",
			ChnFontName = "宋体",
			EngFontName = "Times New Roman",
			FontSize = 10.5f,
			Bold = false,
			Italic = false,
			Underline = false,
			UseColor = true,
			TextColor = Color.FromArgb(255, 9, 96, 167).ToArgb(),
			Shading = true,
			ShadingColor = Color.FromArgb(255, 253, 241, 191).ToArgb()
		};
		formatPainter.StoredFormat.Add(item);
		formatPainter.ToXmlFile(AppDomain.CurrentDomain.BaseDirectory + "\\FormatPainterSettings.xml");
	}

	private void ThisAddIn_Shutdown(object sender, EventArgs e)
	{
		formatPainter.ToXmlFile(AppDomain.CurrentDomain.BaseDirectory + "\\FormatPainterSettings.xml");
		OCR_Engine?.Dispose();
	}

	internal void ApplyPageMargin(Microsoft.Office.Interop.Word.Document wordDoc, bool ApplyToSection, bool setPageMargin, float[] pageMargin, bool setBookbinding, int BookbindingStyle, float gutter)
	{
		if (ApplyToSection)
		{
			Section first = Application.Selection.Sections.First;
			if (setPageMargin)
			{
				for (int i = 0; i < 4; i++)
				{
					pageMargin[i] = Application.CentimetersToPoints(pageMargin[i]);
				}
				first.PageSetup.TopMargin = pageMargin[0];
				first.PageSetup.BottomMargin = pageMargin[1];
				first.PageSetup.LeftMargin = pageMargin[2];
				first.PageSetup.RightMargin = pageMargin[3];
			}
			if (setBookbinding)
			{
				switch (BookbindingStyle)
				{
				case 0:
					first.PageSetup.MirrorMargins = 0;
					first.PageSetup.GutterPos = WdGutterStyle.wdGutterPosLeft;
					break;
				case 1:
					first.PageSetup.MirrorMargins = 0;
					first.PageSetup.GutterPos = WdGutterStyle.wdGutterPosTop;
					break;
				case 2:
					first.PageSetup.GutterPos = WdGutterStyle.wdGutterPosLeft;
					first.PageSetup.MirrorMargins = -1;
					break;
				default:
					first.PageSetup.MirrorMargins = 0;
					break;
				}
				first.PageSetup.Gutter = Application.CentimetersToPoints(gutter);
			}
			return;
		}
		if (setPageMargin)
		{
			for (int j = 0; j < 4; j++)
			{
				pageMargin[j] = Application.CentimetersToPoints(pageMargin[j]);
			}
			foreach (Section section3 in wordDoc.Sections)
			{
				try
				{
					section3.PageSetup.TopMargin = pageMargin[0];
					section3.PageSetup.BottomMargin = pageMargin[1];
					section3.PageSetup.LeftMargin = pageMargin[2];
					section3.PageSetup.RightMargin = pageMargin[3];
				}
				catch
				{
					MessageBox.Show("第" + section3.Index + "节，页边距设置发生错误!", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
			}
		}
		if (!setBookbinding)
		{
			return;
		}
		foreach (Section section4 in wordDoc.Sections)
		{
			try
			{
				switch (BookbindingStyle)
				{
				case 0:
					section4.PageSetup.MirrorMargins = 0;
					section4.PageSetup.GutterPos = WdGutterStyle.wdGutterPosLeft;
					break;
				case 1:
					section4.PageSetup.MirrorMargins = 0;
					section4.PageSetup.GutterPos = WdGutterStyle.wdGutterPosTop;
					break;
				case 2:
					section4.PageSetup.GutterPos = WdGutterStyle.wdGutterPosLeft;
					section4.PageSetup.MirrorMargins = -1;
					break;
				default:
					section4.PageSetup.MirrorMargins = 0;
					break;
				}
				section4.PageSetup.Gutter = Application.CentimetersToPoints(gutter);
			}
			catch
			{
				MessageBox.Show("第" + section4.Index + "节，装订线设置发生错误!", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}
	}

	internal void ResetHeaderFooterStyle([Optional] HeaderFooterFontInfo fontInfo)
	{
		Microsoft.Office.Interop.Word.Application application = Application;
		Styles styles = application.ActiveDocument.Styles;
		object Index = WdBuiltinStyle.wdStyleHeader;
		Style style = styles[ref Index];
		object prop = "";
		style.set_BaseStyle(ref prop);
		Styles styles2 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles2[ref prop].Borders.Enable = 0;
		Styles styles3 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles3[ref prop].ParagraphFormat.LeftIndent = 0f;
		Styles styles4 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles4[ref prop].ParagraphFormat.RightIndent = 0f;
		Styles styles5 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles5[ref prop].ParagraphFormat.FirstLineIndent = 0f;
		Styles styles6 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles6[ref prop].ParagraphFormat.IndentFirstLineCharWidth(0);
		Styles styles7 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles7[ref prop].ParagraphFormat.SpaceAfter = 0f;
		Styles styles8 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles8[ref prop].ParagraphFormat.SpaceBefore = 0f;
		Styles styles9 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles9[ref prop].ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
		Styles styles10 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles10[ref prop].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
		Styles styles11 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleHeader;
		styles11[ref prop].ParagraphFormat.TabStops.ClearAll();
		if (fontInfo.HeaderFontName != null)
		{
			Styles styles12 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleHeader;
			styles12[ref prop].Font.Name = fontInfo.HeaderFontName;
			Styles styles13 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleHeader;
			styles13[ref prop].Font.Size = fontInfo.HeaderFontSize;
			Styles styles14 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleHeader;
			styles14[ref prop].Font.Bold = (fontInfo.HeaderFontBold ? (-1) : 0);
			Styles styles15 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleHeader;
			styles15[ref prop].Font.Italic = (fontInfo.HeaderFontItalic ? (-1) : 0);
		}
		Styles styles16 = application.ActiveDocument.Styles;
		Index = WdBuiltinStyle.wdStyleFooter;
		Style style2 = styles16[ref Index];
		prop = "";
		style2.set_BaseStyle(ref prop);
		Styles styles17 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles17[ref prop].Borders.Enable = 0;
		Styles styles18 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles18[ref prop].ParagraphFormat.LeftIndent = 0f;
		Styles styles19 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles19[ref prop].ParagraphFormat.RightIndent = 0f;
		Styles styles20 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles20[ref prop].ParagraphFormat.FirstLineIndent = 0f;
		Styles styles21 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles21[ref prop].ParagraphFormat.IndentFirstLineCharWidth(0);
		Styles styles22 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles22[ref prop].ParagraphFormat.SpaceAfter = 0f;
		Styles styles23 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles23[ref prop].ParagraphFormat.SpaceBefore = 0f;
		Styles styles24 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles24[ref prop].ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
		Styles styles25 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles25[ref prop].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
		Styles styles26 = application.ActiveDocument.Styles;
		prop = WdBuiltinStyle.wdStyleFooter;
		styles26[ref prop].ParagraphFormat.TabStops.ClearAll();
		if (fontInfo.FooterFontName != null)
		{
			Styles styles27 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleFooter;
			styles27[ref prop].Font.Name = fontInfo.FooterFontName;
			Styles styles28 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleFooter;
			styles28[ref prop].Font.Size = fontInfo.FooterFontSize;
			Styles styles29 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleFooter;
			styles29[ref prop].Font.Bold = (fontInfo.FooterFontBold ? (-1) : 0);
			Styles styles30 = application.ActiveDocument.Styles;
			prop = WdBuiltinStyle.wdStyleFooter;
			styles30[ref prop].Font.Italic = (fontInfo.FooterFontItalic ? (-1) : 0);
		}
	}

	internal void SetHeaderFooter(HeaderFooterTextInfo textInfo, float HeaderHeight, float FooterHeight, HeaderFooterFontInfo fontInfo, bool ClearCurrent, [Optional] object InfoLab)
	{
		Microsoft.Office.Interop.Word.Application application = Application;
		float logoHeight = textInfo.LogoHeight * fontInfo.HeaderFontSize;
		float logoHeight2 = textInfo.LogoHeight * fontInfo.FooterFontSize;
		Range range = application.Selection.Range;
		object Direction = WdCollapseDirection.wdCollapseStart;
		range.Collapse(ref Direction);
		Sections sections = ((textInfo.ApplyModel == 0) ? application.Selection.Sections : application.ActiveDocument.Sections);
		application.ScreenUpdating = false;
		ResetHeaderFooterStyle(fontInfo);
		HeaderFooterTextType[] array;
		HeaderFooterTextType[] array2;
		if (!textInfo.FirstPageDiffrent)
		{
			if (textInfo.OddEvenPageDiffrent)
			{
				array = new HeaderFooterTextType[2];
				array2 = new HeaderFooterTextType[2];
				array[0] = GetTextType(textInfo.PrimaryHeaderText);
				array[1] = GetTextType(textInfo.EvenHeaderText);
				array2[0] = GetTextType(textInfo.PrimaryFooterText);
				array2[1] = GetTextType(textInfo.EvenFooterText);
			}
			else
			{
				array = new HeaderFooterTextType[1];
				array2 = new HeaderFooterTextType[1];
				array[0] = GetTextType(textInfo.PrimaryHeaderText);
				array2[0] = GetTextType(textInfo.PrimaryFooterText);
			}
		}
		else
		{
			array = new HeaderFooterTextType[3];
			array2 = new HeaderFooterTextType[3];
			array[0] = GetTextType(textInfo.PrimaryHeaderText);
			array[1] = GetTextType(textInfo.EvenHeaderText);
			array[2] = GetTextType(textInfo.FirstHeaderText);
			array2[0] = GetTextType(textInfo.PrimaryFooterText);
			array2[1] = GetTextType(textInfo.EvenFooterText);
			array2[2] = GetTextType(textInfo.FirstFooterText);
		}
		if (textInfo.ApplyModel == 0)
		{
			Section section = ((sections.Last.Index < application.ActiveDocument.Sections.Count) ? application.ActiveDocument.Sections[sections.Last.Index + 1] : null);
			if (section != null)
			{
				section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
				section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
				if (textInfo.FirstPageDiffrent)
				{
					section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
					section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
				}
				if (textInfo.OddEvenPageDiffrent)
				{
					section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
					section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
				}
			}
			Section first = sections.First;
			if (textInfo.SameHeaderFooterHeight)
			{
				first.PageSetup.HeaderDistance = application.CentimetersToPoints(HeaderHeight);
				first.PageSetup.FooterDistance = application.CentimetersToPoints(FooterHeight);
			}
			float num = first.PageSetup.PageWidth - first.PageSetup.LeftMargin - first.PageSetup.RightMargin;
			num = ((first.PageSetup.GutterPos == WdGutterStyle.wdGutterPosTop) ? num : (num - first.PageSetup.Gutter));
			HeaderFooter headerFooter;
			HeaderFooter headerFooter2;
			if (textInfo.FirstPageDiffrent)
			{
				first.PageSetup.DifferentFirstPageHeaderFooter = -1;
				if (array[2] != HeaderFooterTextType.None)
				{
					headerFooter = first.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
					headerFooter.LinkToPrevious = false;
					InsertHeaderFooter(ApplyToSection: true, headerFooter, textInfo.FirstHeaderText, array[2], num, textInfo.LogoPath, logoHeight);
				}
				else if (ClearCurrent)
				{
					Range range2 = first.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
					Direction = Type.Missing;
					object Count = Type.Missing;
					range2.Delete(ref Direction, ref Count);
				}
				if (array2[2] != HeaderFooterTextType.None)
				{
					headerFooter2 = first.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
					headerFooter2.LinkToPrevious = false;
					InsertHeaderFooter(ApplyToSection: true, headerFooter2, textInfo.FirstFooterText, array2[2], num, textInfo.LogoPath, logoHeight2);
				}
				else if (ClearCurrent)
				{
					Range range3 = first.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
					object Count = Type.Missing;
					Direction = Type.Missing;
					range3.Delete(ref Count, ref Direction);
				}
			}
			else
			{
				first.PageSetup.DifferentFirstPageHeaderFooter = 0;
			}
			if (textInfo.OddEvenPageDiffrent)
			{
				first.PageSetup.OddAndEvenPagesHeaderFooter = -1;
				headerFooter = first.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
				if (array[1] != HeaderFooterTextType.None)
				{
					headerFooter.LinkToPrevious = false;
					InsertHeaderFooter(ApplyToSection: true, headerFooter, textInfo.EvenHeaderText, array[1], num, textInfo.LogoPath, logoHeight);
				}
				else if (ClearCurrent)
				{
					Range range4 = headerFooter.Range;
					Direction = Type.Missing;
					object Count = Type.Missing;
					range4.Delete(ref Direction, ref Count);
				}
				if (textInfo.HeaderLineType != 5)
				{
					InsertSplitLine(headerFooter, isHeader: true, textInfo.HeaderLineType);
				}
				else
				{
					headerFooter.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
				}
				headerFooter2 = first.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
				if (array2[1] != HeaderFooterTextType.None)
				{
					headerFooter2.LinkToPrevious = false;
					InsertHeaderFooter(ApplyToSection: true, headerFooter2, textInfo.EvenFooterText, array2[1], num, textInfo.LogoPath, logoHeight2);
				}
				else if (ClearCurrent)
				{
					Range range5 = headerFooter2.Range;
					object Count = Type.Missing;
					Direction = Type.Missing;
					range5.Delete(ref Count, ref Direction);
				}
				if (textInfo.FooterLineType != 5)
				{
					InsertSplitLine(headerFooter2, isHeader: false, textInfo.FooterLineType);
				}
				else
				{
					headerFooter2.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
				}
			}
			else
			{
				first.PageSetup.OddAndEvenPagesHeaderFooter = 0;
			}
			headerFooter = first.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
			if (array[0] != HeaderFooterTextType.None)
			{
				headerFooter.LinkToPrevious = false;
				headerFooter.PageNumbers.RestartNumberingAtSection = textInfo.PageNumberStartAtSection;
				headerFooter.PageNumbers.StartingNumber = ((!textInfo.PageNumberStartAtSection) ? 1 : textInfo.StartNumber);
				InsertHeaderFooter(ApplyToSection: true, headerFooter, textInfo.PrimaryHeaderText, array[0], num, textInfo.LogoPath, logoHeight);
			}
			else if (ClearCurrent)
			{
				Range range6 = headerFooter.Range;
				Direction = Type.Missing;
				object Count = Type.Missing;
				range6.Delete(ref Direction, ref Count);
			}
			if (textInfo.HeaderLineType != 5)
			{
				InsertSplitLine(headerFooter, isHeader: true, textInfo.HeaderLineType);
			}
			else
			{
				headerFooter.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
			}
			headerFooter2 = first.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
			if (array2[0] != HeaderFooterTextType.None)
			{
				headerFooter2.LinkToPrevious = false;
				headerFooter2.PageNumbers.RestartNumberingAtSection = textInfo.PageNumberStartAtSection;
				headerFooter2.PageNumbers.StartingNumber = ((!textInfo.PageNumberStartAtSection) ? 1 : textInfo.StartNumber);
				InsertHeaderFooter(ApplyToSection: true, headerFooter2, textInfo.PrimaryFooterText, array2[0], num, textInfo.LogoPath, logoHeight2);
			}
			else if (ClearCurrent)
			{
				Range range7 = headerFooter2.Range;
				object Count = Type.Missing;
				Direction = Type.Missing;
				range7.Delete(ref Count, ref Direction);
			}
			if (textInfo.FooterLineType != 5)
			{
				InsertSplitLine(headerFooter2, isHeader: false, textInfo.FooterLineType);
			}
			else
			{
				headerFooter2.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
			}
			if (InfoLab != null)
			{
				(InfoLab as Label).Text = "设置完成！";
			}
		}
		else
		{
			int num2 = ((textInfo.ApplyModel == 1) ? 1 : application.Selection.Sections.First.Index);
			if (textInfo.FirstPageDiffrent)
			{
				Section section2 = application.ActiveDocument.Sections[num2];
				float num = section2.PageSetup.PageWidth - section2.PageSetup.LeftMargin - section2.PageSetup.RightMargin;
				num = ((section2.PageSetup.GutterPos == WdGutterStyle.wdGutterPosTop) ? num : (num - section2.PageSetup.Gutter));
				section2.PageSetup.DifferentFirstPageHeaderFooter = -1;
				if (array[2] != HeaderFooterTextType.None)
				{
					HeaderFooter headerFooter = section2.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
					headerFooter.LinkToPrevious = false;
					InsertHeaderFooter(ApplyToSection: false, headerFooter, textInfo.FirstHeaderText, array[2], num, textInfo.LogoPath, logoHeight);
				}
				else if (ClearCurrent)
				{
					Range range8 = section2.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
					Direction = Type.Missing;
					object Count = Type.Missing;
					range8.Delete(ref Direction, ref Count);
				}
				if (array2[2] != HeaderFooterTextType.None)
				{
					HeaderFooter headerFooter2 = section2.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
					headerFooter2.LinkToPrevious = false;
					InsertHeaderFooter(ApplyToSection: false, headerFooter2, textInfo.FirstFooterText, array2[2], num, textInfo.LogoPath, logoHeight2);
				}
				else if (ClearCurrent)
				{
					Range range9 = section2.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
					object Count = Type.Missing;
					Direction = Type.Missing;
					range9.Delete(ref Count, ref Direction);
				}
			}
			else
			{
				application.ActiveDocument.Sections[num2].PageSetup.DifferentFirstPageHeaderFooter = 0;
			}
			bool flag = true;
			float num3 = sections[1].PageSetup.PageWidth - sections[1].PageSetup.LeftMargin - sections[1].PageSetup.RightMargin;
			num3 = ((sections[1].PageSetup.GutterPos == WdGutterStyle.wdGutterPosTop) ? num3 : (num3 - sections[1].PageSetup.Gutter));
			foreach (Section item in sections)
			{
				if (item.Index < num2)
				{
					continue;
				}
				if (InfoLab != null)
				{
					(InfoLab as Label).Text = "当前设置第" + item.Index + "节，共" + application.ActiveDocument.Sections.Count + "节。";
				}
				if (textInfo.SameHeaderFooterHeight)
				{
					item.PageSetup.HeaderDistance = application.CentimetersToPoints(HeaderHeight);
					item.PageSetup.FooterDistance = application.CentimetersToPoints(FooterHeight);
				}
				float num = item.PageSetup.PageWidth - item.PageSetup.LeftMargin - item.PageSetup.RightMargin;
				num = ((item.PageSetup.GutterPos == WdGutterStyle.wdGutterPosTop) ? num : (num - item.PageSetup.Gutter));
				if (num == num3 && !flag)
				{
					item.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
					item.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = true;
					if (textInfo.OddEvenPageDiffrent)
					{
						item.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
						item.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = true;
					}
					continue;
				}
				num3 = num;
				HeaderFooter headerFooter;
				HeaderFooter headerFooter2;
				if (textInfo.OddEvenPageDiffrent)
				{
					item.PageSetup.OddAndEvenPagesHeaderFooter = -1;
					headerFooter = item.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
					if (array[1] != HeaderFooterTextType.None)
					{
						headerFooter.LinkToPrevious = false;
						if (flag)
						{
							headerFooter.PageNumbers.RestartNumberingAtSection = textInfo.PageNumberStartAtSection;
							headerFooter.PageNumbers.StartingNumber = ((!textInfo.PageNumberStartAtSection) ? 1 : textInfo.StartNumber);
						}
						InsertHeaderFooter(ApplyToSection: false, headerFooter, textInfo.EvenHeaderText, array[1], num, textInfo.LogoPath, logoHeight);
					}
					else if (ClearCurrent)
					{
						Range range10 = headerFooter.Range;
						Direction = Type.Missing;
						object Count = Type.Missing;
						range10.Delete(ref Direction, ref Count);
					}
					if (textInfo.HeaderLineType != 5)
					{
						InsertSplitLine(headerFooter, isHeader: true, textInfo.HeaderLineType);
					}
					else
					{
						headerFooter.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
					}
					headerFooter2 = item.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
					if (array2[1] != HeaderFooterTextType.None)
					{
						headerFooter2.LinkToPrevious = false;
						if (flag)
						{
							headerFooter2.PageNumbers.RestartNumberingAtSection = textInfo.PageNumberStartAtSection;
							headerFooter2.PageNumbers.StartingNumber = ((!textInfo.PageNumberStartAtSection) ? 1 : textInfo.StartNumber);
						}
						InsertHeaderFooter(ApplyToSection: false, headerFooter2, textInfo.EvenFooterText, array2[1], num, textInfo.LogoPath, logoHeight2);
					}
					else if (ClearCurrent)
					{
						Range range11 = headerFooter2.Range;
						object Count = Type.Missing;
						Direction = Type.Missing;
						range11.Delete(ref Count, ref Direction);
					}
					if (textInfo.FooterLineType != 5)
					{
						InsertSplitLine(headerFooter2, isHeader: false, textInfo.FooterLineType);
					}
					else
					{
						headerFooter2.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
					}
				}
				else
				{
					item.PageSetup.OddAndEvenPagesHeaderFooter = 0;
				}
				headerFooter = item.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
				if (array[0] != HeaderFooterTextType.None)
				{
					headerFooter.LinkToPrevious = false;
					if (flag)
					{
						headerFooter.PageNumbers.RestartNumberingAtSection = textInfo.PageNumberStartAtSection;
						headerFooter.PageNumbers.StartingNumber = ((!textInfo.PageNumberStartAtSection) ? 1 : textInfo.StartNumber);
					}
					InsertHeaderFooter(ApplyToSection: false, headerFooter, textInfo.PrimaryHeaderText, array[0], num, textInfo.LogoPath, logoHeight);
				}
				else if (ClearCurrent)
				{
					Range range12 = headerFooter.Range;
					Direction = Type.Missing;
					object Count = Type.Missing;
					range12.Delete(ref Direction, ref Count);
				}
				if (textInfo.HeaderLineType != 5)
				{
					InsertSplitLine(headerFooter, isHeader: true, textInfo.HeaderLineType);
				}
				else
				{
					headerFooter.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
				}
				headerFooter2 = item.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
				if (array2[0] != HeaderFooterTextType.None)
				{
					headerFooter2.LinkToPrevious = false;
					if (flag)
					{
						headerFooter2.PageNumbers.RestartNumberingAtSection = textInfo.PageNumberStartAtSection;
						headerFooter2.PageNumbers.StartingNumber = ((!textInfo.PageNumberStartAtSection) ? 1 : textInfo.StartNumber);
					}
					InsertHeaderFooter(ApplyToSection: false, headerFooter2, textInfo.PrimaryFooterText, array2[0], num, textInfo.LogoPath, logoHeight2);
				}
				else if (ClearCurrent)
				{
					Range range13 = headerFooter2.Range;
					object Count = Type.Missing;
					Direction = Type.Missing;
					range13.Delete(ref Count, ref Direction);
				}
				if (textInfo.FooterLineType != 5)
				{
					InsertSplitLine(headerFooter2, isHeader: false, textInfo.FooterLineType);
				}
				else
				{
					headerFooter2.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
				}
				flag = false;
			}
		}
		application.ScreenUpdating = true;
		if (InfoLab != null)
		{
			(InfoLab as Label).Text = "设置完成！";
		}
	}

	private HeaderFooterTextType GetTextType(string[] text)
	{
		if (text[0] == "")
		{
			if (text[1] == "")
			{
				if (text[2] == "")
				{
					return HeaderFooterTextType.None;
				}
				return HeaderFooterTextType.Right;
			}
			if (text[2] == "")
			{
				return HeaderFooterTextType.Center;
			}
			return HeaderFooterTextType.CenterRight;
		}
		if (text[1] == "")
		{
			if (text[2] == "")
			{
				return HeaderFooterTextType.Left;
			}
			return HeaderFooterTextType.LeftRight;
		}
		if (text[2] == "")
		{
			return HeaderFooterTextType.LeftCenter;
		}
		return HeaderFooterTextType.All;
	}

	internal void InsertHeaderFooter(bool ApplyToSection, HeaderFooter ThisHeaderFooter, string[] HeaderFooterText, HeaderFooterTextType textType, float innerWidth, string[] LogoPath, float LogoHeight)
	{
		string text = "";
		Range range = ThisHeaderFooter.Range;
		object Unit = Type.Missing;
		object Count = Type.Missing;
		range.Delete(ref Unit, ref Count);
		ThisHeaderFooter.Range.ParagraphFormat.TabStops.ClearAll();
		ThisHeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
		switch (textType)
		{
		case HeaderFooterTextType.Left:
			text = HeaderFooterText[0];
			break;
		case HeaderFooterTextType.Right:
			ThisHeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
			text = HeaderFooterText[2];
			break;
		case HeaderFooterTextType.Center:
			ThisHeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			text = HeaderFooterText[1];
			break;
		case HeaderFooterTextType.CenterRight:
		case HeaderFooterTextType.All:
		{
			TabStops tabStops3 = ThisHeaderFooter.Range.ParagraphFormat.TabStops;
			float position2 = innerWidth / 2f;
			Count = WdTabAlignment.wdAlignTabCenter;
			Unit = Type.Missing;
			tabStops3.Add(position2, ref Count, ref Unit);
			TabStops tabStops4 = ThisHeaderFooter.Range.ParagraphFormat.TabStops;
			Unit = WdTabAlignment.wdAlignTabRight;
			Count = Type.Missing;
			tabStops4.Add(innerWidth, ref Unit, ref Count);
			text = HeaderFooterText[0] + "\t" + HeaderFooterText[1] + "\t" + HeaderFooterText[2];
			break;
		}
		case HeaderFooterTextType.LeftCenter:
		{
			TabStops tabStops2 = ThisHeaderFooter.Range.ParagraphFormat.TabStops;
			float position = innerWidth / 2f;
			Count = WdTabAlignment.wdAlignTabCenter;
			Unit = Type.Missing;
			tabStops2.Add(position, ref Count, ref Unit);
			text = HeaderFooterText[0] + "\t" + HeaderFooterText[1];
			break;
		}
		case HeaderFooterTextType.LeftRight:
		{
			TabStops tabStops = ThisHeaderFooter.Range.ParagraphFormat.TabStops;
			Unit = WdTabAlignment.wdAlignTabRight;
			Count = Type.Missing;
			tabStops.Add(innerWidth, ref Unit, ref Count);
			text = HeaderFooterText[0] + "\t" + HeaderFooterText[2];
			break;
		}
		}
		ThisHeaderFooter.Range.Text = text;
		Range range2;
		object MatchKashida;
		object MatchDiacritics;
		object MatchAlefHamza;
		object MatchControl;
		object Replace;
		object ReplaceWith;
		object Format;
		object Wrap;
		object Forward;
		object MatchAllWordForms;
		object MatchSoundsLike;
		object MatchWildcards;
		object MatchWholeWord;
		if (text.Contains("#"))
		{
			range2 = ThisHeaderFooter.Range;
			Find find = range2.Find;
			Count = "#";
			Unit = true;
			MatchWholeWord = Type.Missing;
			MatchWildcards = Type.Missing;
			MatchSoundsLike = Type.Missing;
			MatchAllWordForms = Type.Missing;
			Forward = Type.Missing;
			Wrap = Type.Missing;
			Format = Type.Missing;
			ReplaceWith = Type.Missing;
			Replace = Type.Missing;
			MatchKashida = Type.Missing;
			MatchDiacritics = Type.Missing;
			MatchAlefHamza = Type.Missing;
			MatchControl = Type.Missing;
			find.Execute(ref Count, ref Unit, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
			while (range2.Find.Found)
			{
				Fields fields = ThisHeaderFooter.Range.Fields;
				Range range3 = range2;
				MatchControl = WdFieldType.wdFieldPage;
				MatchAlefHamza = Type.Missing;
				MatchDiacritics = Type.Missing;
				fields.Add(range3, ref MatchControl, ref MatchAlefHamza, ref MatchDiacritics);
				range2 = ThisHeaderFooter.Range;
				Find find2 = range2.Find;
				MatchDiacritics = "#";
				MatchAlefHamza = true;
				MatchControl = Type.Missing;
				MatchKashida = Type.Missing;
				Replace = Type.Missing;
				ReplaceWith = Type.Missing;
				Format = Type.Missing;
				Wrap = Type.Missing;
				Forward = Type.Missing;
				MatchAllWordForms = Type.Missing;
				MatchSoundsLike = Type.Missing;
				MatchWildcards = Type.Missing;
				MatchWholeWord = Type.Missing;
				Unit = Type.Missing;
				Count = Type.Missing;
				find2.Execute(ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl, ref MatchKashida, ref Replace, ref ReplaceWith, ref Format, ref Wrap, ref Forward, ref MatchAllWordForms, ref MatchSoundsLike, ref MatchWildcards, ref MatchWholeWord, ref Unit, ref Count);
			}
		}
		if (text.Contains("$"))
		{
			WdFieldType wdFieldType = (ApplyToSection ? WdFieldType.wdFieldSectionPages : WdFieldType.wdFieldNumPages);
			range2 = ThisHeaderFooter.Range;
			Find find3 = range2.Find;
			Count = "$";
			Unit = true;
			MatchWholeWord = Type.Missing;
			MatchWildcards = Type.Missing;
			MatchSoundsLike = Type.Missing;
			MatchAllWordForms = Type.Missing;
			Forward = Type.Missing;
			Wrap = Type.Missing;
			Format = Type.Missing;
			ReplaceWith = Type.Missing;
			Replace = Type.Missing;
			MatchKashida = Type.Missing;
			MatchControl = Type.Missing;
			MatchAlefHamza = Type.Missing;
			MatchDiacritics = Type.Missing;
			find3.Execute(ref Count, ref Unit, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchControl, ref MatchAlefHamza, ref MatchDiacritics);
			while (range2.Find.Found)
			{
				Fields fields2 = ThisHeaderFooter.Range.Fields;
				Range range4 = range2;
				MatchDiacritics = wdFieldType;
				MatchAlefHamza = Type.Missing;
				MatchControl = Type.Missing;
				fields2.Add(range4, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
				range2 = ThisHeaderFooter.Range;
				Find find4 = range2.Find;
				MatchControl = "$";
				MatchAlefHamza = true;
				MatchDiacritics = Type.Missing;
				MatchKashida = Type.Missing;
				Replace = Type.Missing;
				ReplaceWith = Type.Missing;
				Format = Type.Missing;
				Wrap = Type.Missing;
				Forward = Type.Missing;
				MatchAllWordForms = Type.Missing;
				MatchSoundsLike = Type.Missing;
				MatchWildcards = Type.Missing;
				MatchWholeWord = Type.Missing;
				Unit = Type.Missing;
				Count = Type.Missing;
				find4.Execute(ref MatchControl, ref MatchAlefHamza, ref MatchDiacritics, ref MatchKashida, ref Replace, ref ReplaceWith, ref Format, ref Wrap, ref Forward, ref MatchAllWordForms, ref MatchSoundsLike, ref MatchWildcards, ref MatchWholeWord, ref Unit, ref Count);
			}
		}
		if (text.Contains("[LOGO1]"))
		{
			range2 = ThisHeaderFooter.Range;
			Find find5 = range2.Find;
			Count = "[LOGO1]";
			Unit = true;
			MatchWholeWord = Type.Missing;
			MatchWildcards = Type.Missing;
			MatchSoundsLike = Type.Missing;
			MatchAllWordForms = Type.Missing;
			Forward = Type.Missing;
			Wrap = Type.Missing;
			Format = Type.Missing;
			ReplaceWith = Type.Missing;
			Replace = Type.Missing;
			MatchKashida = Type.Missing;
			MatchDiacritics = Type.Missing;
			MatchAlefHamza = Type.Missing;
			MatchControl = Type.Missing;
			find5.Execute(ref Count, ref Unit, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
			while (range2.Find.Found)
			{
				range2.Text = "";
				try
				{
					InlineShapes inlineShapes = ThisHeaderFooter.Range.InlineShapes;
					string fileName = LogoPath[0];
					MatchControl = range2;
					MatchAlefHamza = Type.Missing;
					MatchDiacritics = Type.Missing;
					MatchKashida = MatchControl;
					InlineShape inlineShape = inlineShapes.AddPicture(fileName, ref MatchAlefHamza, ref MatchDiacritics, ref MatchKashida);
					inlineShape.LockAspectRatio = MsoTriState.msoCTrue;
					inlineShape.Height = LogoHeight;
					range2 = ThisHeaderFooter.Range;
					Find find6 = range2.Find;
					MatchKashida = "[LOGO1]";
					MatchDiacritics = true;
					MatchAlefHamza = Type.Missing;
					MatchControl = Type.Missing;
					Replace = Type.Missing;
					ReplaceWith = Type.Missing;
					Format = Type.Missing;
					Wrap = Type.Missing;
					Forward = Type.Missing;
					MatchAllWordForms = Type.Missing;
					MatchSoundsLike = Type.Missing;
					MatchWildcards = Type.Missing;
					MatchWholeWord = Type.Missing;
					Unit = Type.Missing;
					Count = Type.Missing;
					find6.Execute(ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl, ref Replace, ref ReplaceWith, ref Format, ref Wrap, ref Forward, ref MatchAllWordForms, ref MatchSoundsLike, ref MatchWildcards, ref MatchWholeWord, ref Unit, ref Count);
				}
				catch
				{
				}
			}
		}
		if (text.Contains("[LOGO2]"))
		{
			range2 = ThisHeaderFooter.Range;
			Find find7 = range2.Find;
			Count = "[LOGO2]";
			Unit = true;
			MatchWholeWord = Type.Missing;
			MatchWildcards = Type.Missing;
			MatchSoundsLike = Type.Missing;
			MatchAllWordForms = Type.Missing;
			Forward = Type.Missing;
			Wrap = Type.Missing;
			Format = Type.Missing;
			ReplaceWith = Type.Missing;
			Replace = Type.Missing;
			MatchControl = Type.Missing;
			MatchAlefHamza = Type.Missing;
			MatchDiacritics = Type.Missing;
			MatchKashida = Type.Missing;
			find7.Execute(ref Count, ref Unit, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchControl, ref MatchAlefHamza, ref MatchDiacritics, ref MatchKashida);
			while (range2.Find.Found)
			{
				range2.Text = "";
				try
				{
					InlineShapes inlineShapes2 = ThisHeaderFooter.Range.InlineShapes;
					string fileName2 = LogoPath[1];
					MatchKashida = range2;
					MatchDiacritics = Type.Missing;
					MatchAlefHamza = Type.Missing;
					MatchControl = MatchKashida;
					InlineShape inlineShape2 = inlineShapes2.AddPicture(fileName2, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
					inlineShape2.LockAspectRatio = MsoTriState.msoCTrue;
					inlineShape2.Height = LogoHeight;
					range2 = ThisHeaderFooter.Range;
					Find find8 = range2.Find;
					MatchControl = "[LOGO2]";
					MatchAlefHamza = true;
					MatchDiacritics = Type.Missing;
					MatchKashida = Type.Missing;
					Replace = Type.Missing;
					ReplaceWith = Type.Missing;
					Format = Type.Missing;
					Wrap = Type.Missing;
					Forward = Type.Missing;
					MatchAllWordForms = Type.Missing;
					MatchSoundsLike = Type.Missing;
					MatchWildcards = Type.Missing;
					MatchWholeWord = Type.Missing;
					Unit = Type.Missing;
					Count = Type.Missing;
					find8.Execute(ref MatchControl, ref MatchAlefHamza, ref MatchDiacritics, ref MatchKashida, ref Replace, ref ReplaceWith, ref Format, ref Wrap, ref Forward, ref MatchAllWordForms, ref MatchSoundsLike, ref MatchWildcards, ref MatchWholeWord, ref Unit, ref Count);
				}
				catch
				{
				}
			}
		}
		if (!text.Contains("[LOGO3]"))
		{
			return;
		}
		range2 = ThisHeaderFooter.Range;
		Find find9 = range2.Find;
		Count = "[LOGO3]";
		Unit = true;
		MatchWholeWord = Type.Missing;
		MatchWildcards = Type.Missing;
		MatchSoundsLike = Type.Missing;
		MatchAllWordForms = Type.Missing;
		Forward = Type.Missing;
		Wrap = Type.Missing;
		Format = Type.Missing;
		ReplaceWith = Type.Missing;
		Replace = Type.Missing;
		MatchKashida = Type.Missing;
		MatchDiacritics = Type.Missing;
		MatchAlefHamza = Type.Missing;
		MatchControl = Type.Missing;
		find9.Execute(ref Count, ref Unit, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
		while (range2.Find.Found)
		{
			range2.Text = "";
			try
			{
				InlineShapes inlineShapes3 = ThisHeaderFooter.Range.InlineShapes;
				string fileName3 = LogoPath[2];
				MatchControl = range2;
				MatchAlefHamza = Type.Missing;
				MatchDiacritics = Type.Missing;
				MatchKashida = MatchControl;
				InlineShape inlineShape3 = inlineShapes3.AddPicture(fileName3, ref MatchAlefHamza, ref MatchDiacritics, ref MatchKashida);
				inlineShape3.LockAspectRatio = MsoTriState.msoCTrue;
				inlineShape3.Height = LogoHeight;
				range2 = ThisHeaderFooter.Range;
				Find find10 = range2.Find;
				MatchKashida = "[LOGO3]";
				MatchDiacritics = true;
				MatchAlefHamza = Type.Missing;
				MatchControl = Type.Missing;
				Replace = Type.Missing;
				ReplaceWith = Type.Missing;
				Format = Type.Missing;
				Wrap = Type.Missing;
				Forward = Type.Missing;
				MatchAllWordForms = Type.Missing;
				MatchSoundsLike = Type.Missing;
				MatchWildcards = Type.Missing;
				MatchWholeWord = Type.Missing;
				Unit = Type.Missing;
				Count = Type.Missing;
				find10.Execute(ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl, ref Replace, ref ReplaceWith, ref Format, ref Wrap, ref Forward, ref MatchAllWordForms, ref MatchSoundsLike, ref MatchWildcards, ref MatchWholeWord, ref Unit, ref Count);
			}
			catch
			{
			}
		}
	}

	internal void InsertSplitLine(HeaderFooter ThisHeaderFooter, bool isHeader, int LineType)
	{
		int num = 2;
		if (defaultValue.HeaderFooter_TextLineGap == 1)
		{
			num = 0;
		}
		if (defaultValue.HeaderFooter_TextLineGap == 2)
		{
			num = 5;
		}
		WdLineWidth lineWidth = WdLineWidth.wdLineWidth150pt;
		WdLineStyle wdLineStyle;
		switch (LineType)
		{
		case 0:
			wdLineStyle = WdLineStyle.wdLineStyleSingle;
			break;
		case 1:
			wdLineStyle = WdLineStyle.wdLineStyleDouble;
			lineWidth = WdLineWidth.wdLineWidth150pt;
			break;
		case 2:
			wdLineStyle = WdLineStyle.wdLineStyleThickThinMedGap;
			lineWidth = WdLineWidth.wdLineWidth300pt;
			break;
		case 3:
			wdLineStyle = WdLineStyle.wdLineStyleThinThickMedGap;
			lineWidth = WdLineWidth.wdLineWidth300pt;
			break;
		default:
			wdLineStyle = WdLineStyle.wdLineStyleSingle;
			lineWidth = WdLineWidth.wdLineWidth300pt;
			break;
		}
		if (isHeader)
		{
			ThisHeaderFooter.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = wdLineStyle;
			ThisHeaderFooter.Range.Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth;
			ThisHeaderFooter.Range.Borders.DistanceFromBottom = num;
		}
		else
		{
			ThisHeaderFooter.Range.Borders[WdBorderType.wdBorderTop].LineStyle = wdLineStyle;
			ThisHeaderFooter.Range.Borders[WdBorderType.wdBorderTop].LineWidth = lineWidth;
			ThisHeaderFooter.Range.Borders.DistanceFromTop = num;
		}
	}

	internal void CreateListTemplate(int ApplyLevel, int[] NumStyle, string[] NumFormat, float[] NumIndent, float[] TextIndent, float[] AfterNumIndent, string[] LinkStyle)
	{
		ListTemplate listTemplate;
		object Index;
		if (Application.Selection.Range.ListFormat.ListTemplate != null)
		{
			listTemplate = Application.Selection.Range.ListFormat.ListTemplate;
		}
		else
		{
			ListTemplates listTemplates = Application.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates;
			Index = 7;
			listTemplate = listTemplates[ref Index];
		}
		ListTemplate listTemplate2 = listTemplate;
		for (int i = 1; i <= ApplyLevel; i++)
		{
			ListLevel listLevel = listTemplate2.ListLevels[i];
			if (NumStyle[i - 1] != -1)
			{
				listLevel.NumberStyle = LevelNumStyle[NumStyle[i - 1]];
			}
			listLevel.NumberFormat = NumFormat[i - 1];
			listLevel.LinkedStyle = ((LinkStyle[i - 1] == "无") ? "" : LinkStyle[i - 1]);
			listLevel.NumberPosition = Application.CentimetersToPoints(NumIndent[i - 1]);
			listLevel.TextPosition = Application.CentimetersToPoints(TextIndent[i - 1]);
			if (AfterNumIndent[i - 1] != 0f)
			{
				listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
				listLevel.TabPosition = Application.CentimetersToPoints(AfterNumIndent[i - 1]);
			}
			listLevel.StartAt = 1;
			listLevel.ResetOnHigher = i - 1;
		}
		for (int j = ApplyLevel + 1; j <= 9; j++)
		{
			ListLevel listLevel2 = listTemplate2.ListLevels[j];
			listLevel2.NumberFormat = "";
			listLevel2.NumberStyle = WdListNumberStyle.wdListNumberStyleNone;
		}
		ListFormat listFormat = Globals.ThisAddIn.Application.Selection.Range.ListFormat;
		Index = false;
		object ApplyTo = WdListApplyTo.wdListApplyToWholeList;
		object DefaultListBehavior = WdDefaultListBehavior.wdWord9ListBehavior;
		object ApplyLevel2 = ApplyLevel;
		listFormat.ApplyListTemplateWithLevel(listTemplate2, ref Index, ref ApplyTo, ref DefaultListBehavior, ref ApplyLevel2);
	}

	internal void AutoCreateLevelList(int Levels, [Optional] float NumIndent, [Optional] float TextIndent, [Optional] float AfterIndent)
	{
		int[] array = new int[Levels];
		string[] array2 = new string[Levels];
		float[] array3 = new float[Levels];
		float[] array4 = new float[Levels];
		float[] array5 = new float[Levels];
		string[] array6 = new string[Levels];
		float num = 0f;
		string text = "";
		object Index;
		for (int i = 0; i < Levels; i++)
		{
			array[i] = 0;
			array2[i] = ((i == 0) ? ("%" + (i + 1)) : (array2[i - 1] + ".%" + (i + 1)));
			array3[i] = NumIndent;
			array6[i] = "标题 " + (i + 1);
			text = ((i == 0) ? "8" : (text + ".8"));
			Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
			Index = array6[i];
			Style style = styles[ref Index];
			if (style != null)
			{
				System.Drawing.Font font = new System.Drawing.Font(new FontFamily(style.Font.Name ?? style.Font.NameFarEast ?? "宋体"), style.Font.Size);
				num = Math.Max(num, TextRenderer.MeasureText(text, font).Width);
			}
		}
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		float pixels = num;
		Index = Type.Missing;
		num = application.PixelsToPoints(pixels, ref Index);
		num = (float)Math.Ceiling(Globals.ThisAddIn.Application.PointsToCentimeters(num) * 10f) / 10f + 0.5f;
		for (int j = 0; j < Levels; j++)
		{
			array4[j] = ((TextIndent == 0f) ? num : TextIndent);
			array5[j] = ((AfterIndent == 0f) ? num : AfterIndent);
		}
		CreateListTemplate(Levels, array, array2, array3, array4, array5, array6);
	}

	internal void ListFormat(List ListApplyTo, int NumberStyle, string NumberForamt, float NumIndent, float TextIndent, float AfterNumIndent)
	{
		if (ListApplyTo == null)
		{
			return;
		}
		WdListType listType = ListApplyTo.ListParagraphs[1].Range.ListFormat.ListType;
		ListTemplate listTemplate = ListApplyTo.ListParagraphs[1].Range.ListFormat.ListTemplate;
		if (listType != WdListType.wdListOutlineNumbering)
		{
			ListApplyTo.ListParagraphs[1].Range.ListFormat.ListLevelNumber = 1;
			if (NumberStyle != -1)
			{
				listTemplate.ListLevels[1].NumberStyle = LevelNumStyle[NumberStyle];
				listTemplate.ListLevels[1].Font.Reset();
			}
			else if (listType == WdListType.wdListNoNumbering)
			{
				listTemplate.ListLevels[1].NumberStyle = LevelNumStyle[0];
			}
			if (NumberForamt != null)
			{
				listTemplate.ListLevels[1].NumberFormat = NumberForamt;
			}
			else if (listType == WdListType.wdListNoNumbering)
			{
				listTemplate.ListLevels[1].NumberFormat = "%1 ";
			}
			listTemplate.ListLevels[1].NumberPosition = Application.CentimetersToPoints(NumIndent);
			listTemplate.ListLevels[1].TextPosition = Application.CentimetersToPoints(TextIndent);
			if (AfterNumIndent > NumIndent)
			{
				listTemplate.ListLevels[1].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
				listTemplate.ListLevels[1].TabPosition = Application.CentimetersToPoints(AfterNumIndent);
			}
			listTemplate.OutlineNumbered = false;
			object ContinuePreviousList = false;
			object DefaultListBehavior = Type.Missing;
			object ApplyLevel = 1;
			ListApplyTo.ApplyListTemplateWithLevel(listTemplate, ref ContinuePreviousList, ref DefaultListBehavior, ref ApplyLevel);
		}
	}

	internal void SetTableShapeAlignment(object alignmentObject, int alignmentType, float leftIndent, int objectType)
	{
		switch (objectType)
		{
		case 1:
		{
			Table table = alignmentObject as Table;
			table.Rows.Alignment = tableAlignmentType[alignmentType];
			if (alignmentType == 0 && leftIndent != 0f)
			{
				table.Rows.LeftIndent = leftIndent;
			}
			else
			{
				table.Rows.LeftIndent = 0f;
			}
			break;
		}
		case 2:
		{
			InlineShape inlineShape = alignmentObject as InlineShape;
			inlineShape.Range.ParagraphFormat.CharacterUnitLeftIndent = 0f;
			inlineShape.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
			inlineShape.Range.ParagraphFormat.FirstLineIndent = 0f;
			inlineShape.Range.ParagraphFormat.Alignment = inlineShapeAlignmentType[alignmentType];
			if (alignmentType == 0 && leftIndent != 0f)
			{
				inlineShape.Range.ParagraphFormat.LeftIndent = leftIndent;
			}
			else
			{
				inlineShape.Range.ParagraphFormat.LeftIndent = 0f;
			}
			break;
		}
		case 3:
		{
			Shape shape2 = alignmentObject as Shape;
			shape2.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
			if (leftIndent != 0f)
			{
				shape2.Left = leftIndent;
			}
			else
			{
				shape2.Left = (float)shapeAlignmentType[alignmentType];
			}
			break;
		}
		case 4:
		{
			Shape shape = alignmentObject as Shape;
			shape.Anchor.ParagraphFormat.FirstLineIndent = 0f;
			shape.Anchor.ParagraphFormat.Alignment = inlineShapeAlignmentType[alignmentType];
			if (alignmentType == 0 && leftIndent != 0f)
			{
				shape.Anchor.ParagraphFormat.LeftIndent = leftIndent;
			}
			else
			{
				shape.Anchor.ParagraphFormat.LeftIndent = 0f;
			}
			break;
		}
		}
	}

	internal void SetTableFormat(Table tableToSet, [Optional] bool firstRowBold, [Optional] bool firstColumnBold, [Optional] bool setInnerMargin, [Optional] int innerMarginType, [Optional] float innerMargin, [Optional] bool removeUselessLine, [Optional] bool setBorder, [Optional] int borderType, [Optional] int borderLineWidth)
	{
		if (setBorder)
		{
			WdLineStyle wdLineStyle = borderType switch
			{
				2 => lineStyle[3], 
				3 => lineStyle[2], 
				_ => lineStyle[borderType], 
			};
			tableToSet.Borders[WdBorderType.wdBorderLeft].LineStyle = wdLineStyle;
			tableToSet.Borders[WdBorderType.wdBorderLeft].LineWidth = tLineWidth[borderLineWidth];
			tableToSet.Borders[WdBorderType.wdBorderRight].LineStyle = lineStyle[borderType];
			tableToSet.Borders[WdBorderType.wdBorderRight].LineWidth = tLineWidth[borderLineWidth];
			tableToSet.Borders[WdBorderType.wdBorderTop].LineStyle = wdLineStyle;
			tableToSet.Borders[WdBorderType.wdBorderTop].LineWidth = tLineWidth[borderLineWidth];
			tableToSet.Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[borderType];
			tableToSet.Borders[WdBorderType.wdBorderBottom].LineWidth = tLineWidth[borderLineWidth];
		}
		if (firstRowBold)
		{
			for (int i = 1; tableToSet.Range.Cells[i].Range.Rows.HeadingFormat == -1 || tableToSet.Range.Cells[i].RowIndex == 1; i++)
			{
				tableToSet.Range.Cells[i].Range.Font.Bold = -1;
			}
		}
		foreach (Cell cell in tableToSet.Range.Cells)
		{
			if (firstColumnBold && cell.ColumnIndex == 1)
			{
				cell.Range.Font.Bold = -1;
			}
			if (setInnerMargin)
			{
				float num = Globals.ThisAddIn.Application.CentimetersToPoints(innerMargin);
				if (innerMarginType == 0 || innerMarginType == 4 || innerMarginType == 6)
				{
					cell.TopPadding = num;
				}
				if (innerMarginType == 1 || innerMarginType == 4 || innerMarginType == 6)
				{
					cell.BottomPadding = num;
				}
				if (innerMarginType == 2 || innerMarginType == 5 || innerMarginType == 6)
				{
					cell.LeftPadding = num;
				}
				if (innerMarginType == 3 || innerMarginType == 5 || innerMarginType == 6)
				{
					cell.RightPadding = num;
				}
			}
			if (removeUselessLine)
			{
				RemoveSpaceLines(cell.Range, keepInner: true);
			}
		}
	}

	internal void SetPictureFormat(object pictureToSet, int ObjectType, [Optional] bool setSingleSpace, [Optional] bool sameWidth, [Optional] float pWidth, [Optional] bool sameHeight, [Optional] float pHeight)
	{
		switch (ObjectType)
		{
		case 0:
		{
			InlineShape inlineShape = (InlineShape)pictureToSet;
			if (setSingleSpace)
			{
				inlineShape.Range.ParagraphFormat.Space1();
			}
			if (sameWidth && sameHeight)
			{
				inlineShape.LockAspectRatio = MsoTriState.msoFalse;
			}
			else
			{
				inlineShape.LockAspectRatio = MsoTriState.msoTrue;
			}
			if (sameWidth)
			{
				inlineShape.Width = Globals.ThisAddIn.Application.CentimetersToPoints(pWidth);
			}
			if (sameHeight)
			{
				inlineShape.Height = Globals.ThisAddIn.Application.CentimetersToPoints(pHeight);
			}
			break;
		}
		case 1:
		{
			Shape shape = (Shape)pictureToSet;
			if (sameWidth && sameHeight)
			{
				shape.LockAspectRatio = MsoTriState.msoFalse;
			}
			else
			{
				shape.LockAspectRatio = MsoTriState.msoTrue;
			}
			if (sameWidth)
			{
				shape.Width = Application.CentimetersToPoints(pWidth);
			}
			if (sameHeight)
			{
				shape.Height = Application.CentimetersToPoints(pHeight);
			}
			break;
		}
		}
	}

	internal void SetGrid(Microsoft.Office.Interop.Word.Document wordDoc, float fontSize)
	{
		float num = 0.0023f * fontSize + 1.38f;
		int num2 = (int)((wordDoc.Sections[1].PageSetup.PageHeight - wordDoc.Sections[1].PageSetup.TopMargin - wordDoc.Sections[1].PageSetup.BottomMargin) / (num * fontSize));
		wordDoc.GridOriginFromMargin = true;
		wordDoc.SnapToGrid = true;
		foreach (Section section in wordDoc.Sections)
		{
			try
			{
				section.PageSetup.LinesPage = num2;
			}
			catch
			{
			}
		}
	}

	internal void CreateNewTable(bool ThreeLine, bool ThreeLineExtra, bool BroadOuterLine, bool TitleRowFilled, bool SummaryRow, bool SummaryColumn, bool SummaryRowFilled, bool Diagonal)
	{
		WdLineWidth[] array = new WdLineWidth[6]
		{
			WdLineWidth.wdLineWidth025pt,
			WdLineWidth.wdLineWidth050pt,
			WdLineWidth.wdLineWidth075pt,
			WdLineWidth.wdLineWidth150pt,
			WdLineWidth.wdLineWidth225pt,
			WdLineWidth.wdLineWidth300pt
		};
		Application.ScreenUpdating = false;
		Selection selection = Application.Selection;
		object Direction = WdCollapseDirection.wdCollapseEnd;
		selection.Collapse(ref Direction);
		object Position;
		object ExcludeLabel;
		if (tableSettings.CaptionLab)
		{
			CaptionLabels captionLabels = Application.CaptionLabels;
			Direction = tableSettings.CaptionTitle;
			CaptionLabel captionLabel = captionLabels[ref Direction];
			captionLabel.NumberStyle = CaptionNumStyle[tableSettings.CaptionNumberStyle];
			if (tableSettings.CaptionIncludeHeadings)
			{
				captionLabel.IncludeChapterNumber = true;
				captionLabel.ChapterStyleLevel = tableSettings.HeadingsLevel;
				switch (tableSettings.LinkChar)
				{
				case 0:
					captionLabel.Separator = WdSeparatorType.wdSeparatorHyphen;
					break;
				case 1:
					captionLabel.Separator = WdSeparatorType.wdSeparatorPeriod;
					break;
				case 2:
					captionLabel.Separator = WdSeparatorType.wdSeparatorColon;
					break;
				case 3:
					captionLabel.Separator = WdSeparatorType.wdSeparatorEnDash;
					break;
				}
			}
			else
			{
				captionLabel.IncludeChapterNumber = false;
			}
			Range range = Application.Selection.Range;
			Direction = tableSettings.CaptionTitle;
			object Title = " " + tableSettings.TableTitle;
			object TitleAutoText = Type.Missing;
			Position = Type.Missing;
			ExcludeLabel = Type.Missing;
			range.InsertCaption(ref Direction, ref Title, ref TitleAutoText, ref Position, ref ExcludeLabel);
		}
		else
		{
			Application.Selection.Range.InsertBefore(tableSettings.TableTitle);
		}
		Application.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
		Application.Selection.Range.Font.Name = tableSettings.FontName;
		Application.Selection.Range.Font.Size = tableSettings.FontSize;
		Selection selection2 = Application.Selection;
		ExcludeLabel = WdUnits.wdLine;
		Position = Type.Missing;
		selection2.Move(ref ExcludeLabel, ref Position);
		Tables tables = Application.ActiveDocument.Tables;
		Range range2 = Application.Selection.Range;
		int rows = tableSettings.Rows;
		int columns = tableSettings.Columns;
		Position = Type.Missing;
		ExcludeLabel = Type.Missing;
		Table table = tables.Add(range2, rows, columns, ref Position, ref ExcludeLabel);
		table.Range.Font.Name = tableSettings.FontName;
		table.Range.Font.Size = tableSettings.FontSize;
		table.Range.Font.Color = (WdColor)RGB(tableSettings.FontColor.R, tableSettings.FontColor.G, tableSettings.FontColor.B);
		if (tableSettings.FixRowHeight)
		{
			float num = (float)Math.Round((0.0023f * tableSettings.FontSize + 1.38f) * tableSettings.FontSize);
			table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
			table.Range.ParagraphFormat.LineSpacing = num;
			table.Rows.SetHeight(num + 6f, WdRowHeightRule.wdRowHeightExactly);
		}
		else
		{
			table.Rows.SetHeight(Application.CentimetersToPoints(0.8f), WdRowHeightRule.wdRowHeightAtLeast);
		}
		table.ApplyStyleHeadingRows = true;
		table.ApplyStyleFirstColumn = true;
		table.Rows[1].Range.Font.Bold = -1;
		table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
		if (SummaryRow)
		{
			table.Rows[tableSettings.Rows].Range.Font.Bold = -1;
			table.Rows[tableSettings.Rows].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
		}
		foreach (Row row in table.Rows)
		{
			foreach (Cell cell in row.Cells)
			{
				cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
				if (cell.RowIndex == 1)
				{
					cell.Range.Text = "列" + (cell.ColumnIndex - 1);
				}
				if (cell.RowIndex > 1 && cell.ColumnIndex == 1)
				{
					cell.Range.Text = "行" + (cell.RowIndex - 1);
				}
			}
		}
		if (SummaryRow)
		{
			table.Cell(tableSettings.Rows, 1).Range.Text = "汇总行";
		}
		if (SummaryColumn)
		{
			table.Cell(1, tableSettings.Columns).Range.Text = "汇总列";
		}
		WdLineWidth lineWidth = ((tableSettings.OuterLineType == 1) ? array[tableSettings.OuterLineWidth] : tLineWidth[tableSettings.OuterLineWidth]);
		WdLineWidth wdLineWidth = ((tableSettings.InnerLineType == 1) ? array[tableSettings.InnerLineWidth] : tLineWidth[tableSettings.InnerLineWidth]);
		WdLineWidth lineWidth2 = ((tableSettings.TitleRowLineType == 1) ? array[tableSettings.TitleRowLineWidth] : tLineWidth[tableSettings.TitleRowLineWidth]);
		int index = ((tableSettings.OuterLineType == 2) ? 3 : ((tableSettings.OuterLineType != 3) ? tableSettings.OuterLineType : 2));
		int color = RGB(tableSettings.OuterLineColor.R, tableSettings.OuterLineColor.G, tableSettings.OuterLineColor.B);
		int num2 = RGB(tableSettings.InnerLineColor.R, tableSettings.InnerLineColor.G, tableSettings.InnerLineColor.B);
		int color2 = RGB(tableSettings.TitleRowLineColor.R, tableSettings.TitleRowLineColor.G, tableSettings.TitleRowLineColor.B);
		if (ThreeLine || ThreeLineExtra)
		{
			if (ThreeLineExtra)
			{
				Rows rows2 = table.Rows;
				ExcludeLabel = table.Rows[1];
				rows2.Add(ref ExcludeLabel);
				if (TitleRowFilled)
				{
					if (tableSettings.FillType != 1)
					{
						table.Rows[1].Shading.BackgroundPatternColor = (WdColor)RGB(tableSettings.BackgrounColor.R, tableSettings.BackgrounColor.G, tableSettings.BackgrounColor.B);
						table.Rows[2].Shading.BackgroundPatternColor = (WdColor)RGB(tableSettings.BackgrounColor.R, tableSettings.BackgrounColor.G, tableSettings.BackgrounColor.B);
					}
					if (tableSettings.FillType != 0)
					{
						table.Rows[1].Shading.Texture = textureIndex[tableSettings.TextureStyle];
						table.Rows[2].Shading.Texture = textureIndex[tableSettings.TextureStyle];
					}
				}
			}
			else if (TitleRowFilled)
			{
				if (tableSettings.FillType != 1)
				{
					table.Rows[1].Shading.BackgroundPatternColor = (WdColor)RGB(tableSettings.BackgrounColor.R, tableSettings.BackgrounColor.G, tableSettings.BackgrounColor.B);
				}
				if (tableSettings.FillType != 0)
				{
					table.Rows[1].Shading.Texture = textureIndex[tableSettings.TextureStyle];
				}
			}
			table.Rows[1].Borders[WdBorderType.wdBorderTop].LineStyle = lineStyle[index];
			table.Rows[1].Borders[WdBorderType.wdBorderTop].LineWidth = lineWidth;
			table.Rows[1].Borders[WdBorderType.wdBorderTop].Color = (WdColor)color;
			table.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[tableSettings.TitleRowLineType];
			table.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth2;
			table.Rows[1].Borders[WdBorderType.wdBorderBottom].Color = (WdColor)color2;
			table.Rows[table.Rows.Count].Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[tableSettings.OuterLineType];
			table.Rows[table.Rows.Count].Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth;
			table.Rows[table.Rows.Count].Borders[WdBorderType.wdBorderBottom].Color = (WdColor)color;
			if (ThreeLineExtra)
			{
				table.Rows[2].Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[tableSettings.TitleRowLineType];
				table.Rows[2].Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth2;
				table.Rows[2].Borders[WdBorderType.wdBorderBottom].Color = (WdColor)color2;
				table.Cell(1, 1).Merge(table.Cell(2, 1));
				while (true)
				{
					try
					{
						table.Cell(1, 2).Merge(table.Cell(1, 3));
					}
					catch
					{
						break;
					}
				}
				table.Cell(1, 2).Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[tableSettings.TitleRowLineType];
				table.Cell(1, 2).Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth2;
				table.Cell(1, 2).Borders[WdBorderType.wdBorderBottom].Color = (WdColor)color2;
				table.Cell(1, 1).Range.Text = "表头1";
				table.Cell(1, 2).Range.Text = "表头2";
			}
			else
			{
				table.Cell(1, 1).Range.Text = "表头";
			}
			if (SummaryRowFilled)
			{
				if (tableSettings.FillType != 1)
				{
					table.Rows[table.Rows.Count].Shading.BackgroundPatternColor = (WdColor)RGB(tableSettings.BackgrounColor.R, tableSettings.BackgrounColor.G, tableSettings.BackgrounColor.B);
				}
				if (tableSettings.FillType != 0)
				{
					table.Rows[table.Rows.Count].Shading.Texture = textureIndex[tableSettings.TextureStyle];
				}
			}
		}
		else
		{
			table.Borders.InsideLineStyle = lineStyle[tableSettings.InnerLineType];
			table.Borders.InsideLineWidth = wdLineWidth;
			table.Borders.InsideColor = (WdColor)num2;
			if (BroadOuterLine)
			{
				table.Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[tableSettings.OuterLineType];
				table.Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth;
				table.Borders[WdBorderType.wdBorderBottom].Color = (WdColor)color;
				table.Borders[WdBorderType.wdBorderRight].LineStyle = lineStyle[tableSettings.OuterLineType];
				table.Borders[WdBorderType.wdBorderRight].LineWidth = lineWidth;
				table.Borders[WdBorderType.wdBorderRight].Color = (WdColor)color;
				table.Borders[WdBorderType.wdBorderTop].LineStyle = lineStyle[index];
				table.Borders[WdBorderType.wdBorderTop].LineWidth = lineWidth;
				table.Borders[WdBorderType.wdBorderTop].Color = (WdColor)color;
				table.Borders[WdBorderType.wdBorderLeft].LineStyle = lineStyle[index];
				table.Borders[WdBorderType.wdBorderLeft].LineWidth = lineWidth;
				table.Borders[WdBorderType.wdBorderLeft].Color = (WdColor)color;
			}
			else
			{
				table.Borders.OutsideLineStyle = lineStyle[tableSettings.InnerLineType];
				table.Borders.OutsideLineWidth = wdLineWidth;
				table.Borders.OutsideColor = (WdColor)num2;
			}
			table.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = lineStyle[tableSettings.TitleRowLineType];
			table.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = lineWidth2;
			table.Rows[1].Borders[WdBorderType.wdBorderBottom].Color = (WdColor)color2;
			if (Diagonal)
			{
				if (tableSettings.FixRowHeight)
				{
					float num3 = (float)Math.Round((0.0023f * tableSettings.FontSize + 1.38f) * tableSettings.FontSize);
					table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
					table.Range.ParagraphFormat.LineSpacing = num3;
					table.Rows[1].Height = num3 * 2f;
				}
				else
				{
					table.Rows[1].Height = Application.CentimetersToPoints(1.2f);
				}
				table.Cell(1, 1).Borders[WdBorderType.wdBorderDiagonalDown].LineStyle = lineStyle[tableSettings.InnerLineType];
				table.Cell(1, 1).Borders[WdBorderType.wdBorderDiagonalDown].LineWidth = wdLineWidth;
				table.Cell(1, 1).Borders[WdBorderType.wdBorderDiagonalDown].Color = (WdColor)num2;
				table.Cell(1, 1).Range.Text = "表头2\r\n表头1";
				table.Cell(1, 1).Range.Paragraphs[1].Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				table.Cell(1, 1).Range.Paragraphs[2].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
			}
			else
			{
				table.Cell(1, 1).Range.Text = "表头";
			}
			if (TitleRowFilled)
			{
				if (tableSettings.FillType != 1)
				{
					table.Rows[1].Shading.BackgroundPatternColor = (WdColor)RGB(tableSettings.BackgrounColor.R, tableSettings.BackgrounColor.G, tableSettings.BackgrounColor.B);
				}
				if (tableSettings.FillType != 0)
				{
					table.Rows[1].Shading.Texture = textureIndex[tableSettings.TextureStyle];
				}
			}
			if (SummaryRowFilled)
			{
				if (tableSettings.FillType != 1)
				{
					table.Rows[tableSettings.Rows].Shading.BackgroundPatternColor = (WdColor)RGB(tableSettings.BackgrounColor.R, tableSettings.BackgrounColor.G, tableSettings.BackgrounColor.B);
				}
				if (tableSettings.FillType != 0)
				{
					table.Rows[tableSettings.Rows].Shading.Texture = textureIndex[tableSettings.TextureStyle];
				}
			}
		}
		Application.ScreenUpdating = true;
	}

	internal void CreateQRCodeImage(string TextInfo)
	{
		Bitmap dataObject = new QRCodeCreator(new QRCodeConfig
		{
			ErrorCorrectionLevel = (QRErrorCorrectionLevel)defaultValue.QRECCLevel,
			ModuleSize = defaultValue.QRModulePixel,
			DrawQuietZone = defaultValue.QRCodeQuitZone,
			DarkColor = defaultValue.QRCodeDarkColor,
			LightColor = defaultValue.QRCodeLightColor
		}).GenerateQrCode(TextInfo);
		Selection selection = Application.Selection;
		object Direction = WdCollapseDirection.wdCollapseEnd;
		selection.Collapse(ref Direction);
		Clipboard.SetDataObject(dataObject);
		Application.Selection.Range.Paste();
	}

	internal static List<char> GetChnCharSet(int setType)
	{
		List<char> list = new List<char>();
		if ((setType & 1) == 1)
		{
			list.AddRange(new _003C_003Ez__ReadOnlyArray<char>(new char[4] { '，', '。', '：', '；' }));
		}
		if (((setType >> 1) & 1) == 1)
		{
			list.AddRange(new _003C_003Ez__ReadOnlyArray<char>(new char[8] { '（', '）', '［', '］', '｛', '｝', '＜', '＞' }));
		}
		if (((setType >> 2) & 1) == 1)
		{
			list.AddRange(new _003C_003Ez__ReadOnlyArray<char>(new char[3] { '？', '！', '～' }));
		}
		return list;
	}

	internal static List<char> GetEngCharSet(int setType)
	{
		List<char> list = new List<char>();
		if ((setType & 1) == 1)
		{
			list.AddRange(new _003C_003Ez__ReadOnlyArray<char>(new char[4] { ',', '.', ':', ';' }));
		}
		if (((setType >> 1) & 1) == 1)
		{
			list.AddRange(new _003C_003Ez__ReadOnlyArray<char>(new char[8] { '(', ')', '[', ']', '{', '}', '<', '>' }));
		}
		if (((setType >> 2) & 1) == 1)
		{
			list.AddRange(new _003C_003Ez__ReadOnlyArray<char>(new char[3] { '?', '!', '~' }));
		}
		return list;
	}

	internal static int RGB(byte r, byte g, byte b)
	{
		return (r & 0xFF) | ((g & 0xFF) << 8) | ((b & 0xFF) << 16);
	}

	internal static Color ColorFormInt(int color)
	{
		byte red = (byte)color;
		byte green = (byte)(color >> 8);
		byte blue = (byte)(color >> 16);
		return Color.FromArgb(255, red, green, blue);
	}

	internal static void SetSuperscriptOrSubscript(Range target, string scriptText, bool superscript = true, bool useRegex = false)
	{
		Match match = Regex.Match(target.Text, useRegex ? scriptText : Regex.Escape(scriptText));
		while (match.Success)
		{
			int num = target.Start + match.Index;
			if (superscript)
			{
				Microsoft.Office.Interop.Word.Document document = target.Document;
				object Start = num;
				object End = num + match.Value.Length;
				document.Range(ref Start, ref End).Font.Superscript = -1;
			}
			else
			{
				Microsoft.Office.Interop.Word.Document document2 = target.Document;
				object End = num;
				object Start = num + match.Value.Length;
				document2.Range(ref End, ref Start).Font.Subscript = -1;
			}
			match = match.NextMatch();
		}
	}

	internal static void PunctuationWidthSwitch(Range target, bool EngToChn = true)
	{
		List<char> list = (EngToChn ? GetEngCharSet(textFormatSet.SetType) : GetChnCharSet(textFormatSet.SetType));
		List<char> list2 = (EngToChn ? GetChnCharSet(textFormatSet.SetType) : GetEngCharSet(textFormatSet.SetType));
		for (int i = 0; i < list.Count; i++)
		{
			Match match = Regex.Match(pattern: (list[i] != '.' || !textFormatSet.IgnoreNumberDot) ? Regex.Escape(list[i].ToString()) : "(?<![0-9])\\.", input: target.Text);
			while (match.Success)
			{
				int start = target.Start;
				Microsoft.Office.Interop.Word.Document document = target.Document;
				object Start = start + match.Index;
				object End = start + match.Index + match.Value.Length;
				document.Range(ref Start, ref End).Text = list2[i].ToString();
				match = match.NextMatch();
			}
		}
	}

	internal static void RemoveWhiteSpace(Range target)
	{
		string pattern = textFormatSet.RemoveSpaceType switch
		{
			1 => "(?<![A-Za-z])[ \\u00A0\\u2009\\u200A\\u3000]+(?![A-Za-z])", 
			2 => "(?<=[\\u0021\\u0022\\u002C\\u002E\\u003A\\u003B\\u003F\\u3001\\u3002\\uFF01\\uFF0C\\uFF1A\\uFF1B\\uFF1F])[ \\u00A0\\u2009\\u200A\\u3000]+", 
			_ => "[ \\u00A0\\u2009\\u200A\\u3000]+", 
		};
		Match match = Regex.Match(target.Text, pattern);
		while (match.Success)
		{
			int start = target.Start;
			Microsoft.Office.Interop.Word.Document document = target.Document;
			object Start = start + match.Index;
			object End = start + match.Index + match.Value.Length;
			document.Range(ref Start, ref End).Text = "";
			match = Regex.Match(target.Text, pattern);
		}
	}

	internal static void AddBrakets(Range target, int braketsIndex, bool replace = false)
	{
		string[] array = new string[7] { "“", "《", "(", "[", "{", "<", "〔" };
		string[] array2 = new string[7] { "”", "》", ")", "]", "}", ">", "〕" };
		if (textFormatSet.FullWidthBracket)
		{
			array[2] = "（";
			array[3] = "［";
			array[4] = "｛";
			array[5] = "＜";
			array2[2] = "）";
			array2[3] = "］";
			array2[4] = "｝";
			array2[5] = "＞";
		}
		if (replace && Regex.IsMatch(target.Text, "^[“《（［｛＜\\(\\[\\{\\<〔]{1}.*[”》）］｝＞\\)\\]\\}\\>〕]$"))
		{
			Microsoft.Office.Interop.Word.Document document = target.Document;
			object Start = target.Start;
			object End = target.Start + 1;
			document.Range(ref Start, ref End).Text = array[braketsIndex];
			Microsoft.Office.Interop.Word.Document document2 = target.Document;
			End = target.End - 1;
			Start = target.End;
			document2.Range(ref End, ref Start).Text = array2[braketsIndex];
		}
		else
		{
			target.InsertAfter(array2[braketsIndex]);
			target.InsertBefore(array[braketsIndex]);
		}
	}

	internal static void RemoveSpaceLines(Range target, bool keepInner = false)
	{
		Match match = Regex.Match(target.Text, "^[ \\u00A0\\u2009\\u200A\\u3000\\r]+\\r(?=.*)");
		if (match.Success)
		{
			Microsoft.Office.Interop.Word.Document document = target.Document;
			object Start = target.Start + match.Index;
			object End = target.Start + match.Index + match.Value.Length;
			Range range = document.Range(ref Start, ref End);
			object Unit = Type.Missing;
			object Count = Type.Missing;
			range.Delete(ref Unit, ref Count);
		}
		match = Regex.Match(target.Text, "(?<=.*)[ \\u00A0\\u2009\\u200A\\u3000\\r]+(\\r|\\r\\a)$");
		if (match.Success)
		{
			if (match.Value.EndsWith("\a"))
			{
				Microsoft.Office.Interop.Word.Document document2 = target.Document;
				object Count = target.Start + match.Index;
				object Unit = target.Start + match.Index + match.Value.Length - 2;
				document2.Range(ref Count, ref Unit).Text = "";
			}
			else
			{
				Microsoft.Office.Interop.Word.Document document3 = target.Document;
				object Unit = target.Start + match.Index;
				object Count = target.Start + match.Index + match.Value.Length - 1;
				document3.Range(ref Unit, ref Count).Text = "";
			}
		}
		if (keepInner)
		{
			return;
		}
		int num = 1;
		while (num < target.Paragraphs.Count)
		{
			match = Regex.Match(target.Paragraphs[num].Range.Text, "^[ \\u00A0\\u2009\\u200A\\u3000]*(\\r|\\r\\a)$");
			if (match.Success)
			{
				if (match.Value.EndsWith("\a"))
				{
					target.Paragraphs[num].Range.Text = "\r\a";
					continue;
				}
				if (target.Paragraphs.Count == 1)
				{
					target.Paragraphs[num].Range.Text = "\r";
					continue;
				}
				Range range2 = target.Paragraphs[num].Range;
				object Count = Type.Missing;
				object Unit = Type.Missing;
				range2.Delete(ref Count, ref Unit);
			}
			else
			{
				num++;
			}
		}
	}

	internal static void SetIndent2CharOrNot(Range target, bool setIndent = true)
	{
		foreach (Paragraph paragraph in target.Paragraphs)
		{
			if (setIndent)
			{
				paragraph.Format.LeftIndent = 0f;
				paragraph.Format.FirstLineIndent = 0f;
				paragraph.Format.IndentFirstLineCharWidth(2);
			}
			else
			{
				paragraph.Format.LeftIndent = 0f;
				paragraph.Format.FirstLineIndent = 0f;
			}
		}
	}

	private void InternalStartup()
	{
		base.Startup += ThisAddIn_Startup;
		base.Shutdown += ThisAddIn_Shutdown;
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider)
		: base((Microsoft.Office.Tools.Factory)factory, serviceProvider, "AddIn", "ThisAddIn")
	{
		Globals.Factory = factory;
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void Initialize()
	{
		base.Initialize();
		Application = GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
		Globals.ThisAddIn = this;
		System.Windows.Forms.Application.EnableVisualStyles();
		InitializeCachedData();
		InitializeControls();
		InitializeComponents();
		InitializeData();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void FinishInitialization()
	{
		InternalStartup();
		OnStartup();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void InitializeDataBindings()
	{
		BeginInitialization();
		BindToData();
		EndInitialization();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeCachedData()
	{
		if (base.DataHost != null && base.DataHost.IsCacheInitialized)
		{
			base.DataHost.FillCachedData(this);
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeData()
	{
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void BindToData()
	{
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private void StartCaching(string MemberName)
	{
		base.DataHost.StartCaching(this, MemberName);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private void StopCaching(string MemberName)
	{
		base.DataHost.StopCaching(this, MemberName);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private bool IsCached(string MemberName)
	{
		return base.DataHost.IsCached(this, MemberName);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void BeginInitialization()
	{
		BeginInit();
		CustomTaskPanes.BeginInit();
		VstoSmartTags.BeginInit();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void EndInitialization()
	{
		VstoSmartTags.EndInit();
		CustomTaskPanes.EndInit();
		EndInit();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeControls()
	{
		CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
		VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void InitializeComponents()
	{
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private bool NeedsFill(string MemberName)
	{
		return base.DataHost.NeedsFill(this, MemberName);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	protected override void OnShutdown()
	{
		VstoSmartTags.Dispose();
		CustomTaskPanes.Dispose();
		base.OnShutdown();
	}
}
