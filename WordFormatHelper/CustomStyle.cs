using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

internal class CustomStyle
{
	private string name;

	private string fontname;

	private float fontsize;

	private bool bold;

	private bool italic;

	private bool underline;

	private int paraalignment;

	private float leftindent;

	private float firstlineindent;

	private int firstlineindentbychar;

	private float linespacing;

	private float beforespacing;

	private bool beforebreak;

	private float afterspacing;

	private int numberstyle;

	private string numberformat;

	private bool userdefined;

	public string Name
	{
		get
		{
			return name;
		}
		set
		{
			name = value;
		}
	}

	public string FontName
	{
		get
		{
			return fontname;
		}
		set
		{
			fontname = value;
		}
	}

	public float FontSize
	{
		get
		{
			return fontsize;
		}
		set
		{
			fontsize = value;
		}
	}

	public bool IsBold
	{
		get
		{
			return bold;
		}
		set
		{
			bold = value;
		}
	}

	public bool IsItalic
	{
		get
		{
			return italic;
		}
		set
		{
			italic = value;
		}
	}

	public bool IsUnderline
	{
		get
		{
			return underline;
		}
		set
		{
			underline = value;
		}
	}

	public int ParagraphAlignment
	{
		get
		{
			return paraalignment;
		}
		set
		{
			paraalignment = value;
		}
	}

	public float LeftIndent
	{
		get
		{
			return leftindent;
		}
		set
		{
			leftindent = value;
		}
	}

	public float FirstLineIndent
	{
		get
		{
			return firstlineindent;
		}
		set
		{
			firstlineindent = value;
			if (firstlineindent != 0f)
			{
				firstlineindentbychar = 0;
			}
		}
	}

	public int FirstLineIndentByChar
	{
		get
		{
			return firstlineindentbychar;
		}
		set
		{
			firstlineindentbychar = value;
			if (firstlineindentbychar != 0)
			{
				firstlineindent = 0f;
			}
		}
	}

	public float LineSpacing
	{
		get
		{
			return linespacing;
		}
		set
		{
			linespacing = value;
		}
	}

	public float BeforeSpacing
	{
		get
		{
			return beforespacing;
		}
		set
		{
			beforespacing = value;
		}
	}

	public bool BeforeBreak
	{
		get
		{
			return beforebreak;
		}
		set
		{
			beforebreak = value;
		}
	}

	public float AfterSpacing
	{
		get
		{
			return afterspacing;
		}
		set
		{
			afterspacing = value;
		}
	}

	public int NumberStyle
	{
		get
		{
			return numberstyle;
		}
		set
		{
			numberstyle = value;
			if (numberstyle != 0 && numberformat == "")
			{
				numberformat = "%1";
			}
			if (numberstyle == 0)
			{
				numberformat = "";
			}
		}
	}

	public string NumberFormat
	{
		get
		{
			return numberformat;
		}
		set
		{
			numberformat = value;
		}
	}

	public bool UserDefined
	{
		get
		{
			return userdefined;
		}
		set
		{
			userdefined = value;
		}
	}

	public CustomStyle(string name, [Optional] string fontname, [Optional] float fontsize, [Optional] bool bold, [Optional] bool italic, [Optional] bool underline, [Optional] int paraalignment, [Optional] float leftindent, [Optional] float firstlineindent, [Optional] int firstlineindentbychar, [Optional] float linespacing, [Optional] bool beforebreak, [Optional] float beforespacing, [Optional] float afterspacing, [Optional] int numberstyle, [Optional] string numberformat, [Optional] bool userdefined)
	{
		this.name = name;
		this.fontname = fontname ?? "宋体";
		this.fontsize = ((fontsize == 0f) ? 10.5f : fontsize);
		this.bold = bold;
		this.italic = italic;
		this.underline = underline;
		this.paraalignment = paraalignment;
		this.leftindent = leftindent;
		this.firstlineindent = firstlineindent;
		this.firstlineindentbychar = firstlineindentbychar;
		this.linespacing = ((linespacing == 0f) ? 1f : linespacing);
		this.beforespacing = beforespacing;
		this.beforebreak = beforebreak;
		this.afterspacing = afterspacing;
		this.numberstyle = numberstyle;
		this.numberformat = numberformat ?? "";
		this.userdefined = userdefined;
		base._002Ector();
	}

	public void SetValue([Optional] string fontname, [Optional] float fontsize, [Optional] bool bold, [Optional] bool italic, [Optional] bool underline, [Optional] int paraalignment, [Optional] float leftindent, [Optional] float firstlineindent, [Optional] int firstlineindentbychar, [Optional] float linespacing, [Optional] bool beforebreak, [Optional] float beforespacing, [Optional] float afterspacing, [Optional] int numberstyle, [Optional] string numberformat, [Optional] bool userdefined)
	{
		this.fontname = fontname ?? "宋体";
		this.fontsize = ((fontsize == 0f) ? 10.5f : fontsize);
		this.bold = bold;
		this.italic = italic;
		this.underline = underline;
		this.paraalignment = paraalignment;
		this.leftindent = leftindent;
		this.firstlineindent = firstlineindent;
		this.firstlineindentbychar = firstlineindentbychar;
		this.linespacing = ((linespacing == 0f) ? 1f : linespacing);
		this.beforespacing = beforespacing;
		this.beforebreak = beforebreak;
		this.afterspacing = afterspacing;
		this.numberformat = numberformat ?? "";
		this.numberstyle = numberstyle;
		this.userdefined = userdefined;
	}

	public string StyleInfo()
	{
		string[] array = new string[5] { "左对齐", "居中对齐", "右对齐", "两端对齐", "分散对齐" };
		string text = name + "，" + fontname + "，" + fontsize.ToString("0.0") + "磅";
		if (bold)
		{
			text += "，粗体";
		}
		if (italic)
		{
			text += "，斜体";
		}
		if (underline)
		{
			text += ",下划线";
		}
		text = text + "，段落" + array[paraalignment];
		if (leftindent != 0f)
		{
			text = text + "，左缩进" + Globals.ThisAddIn.Application.PointsToCentimeters(leftindent).ToString("0.0") + "厘米";
		}
		if (firstlineindentbychar != 0)
		{
			text = text + "，首行缩进" + firstlineindentbychar.ToString("0") + "字符";
		}
		else if (firstlineindent != 0f)
		{
			text = text + "，首行缩进" + Globals.ThisAddIn.Application.PointsToCentimeters(firstlineindent).ToString("0.0") + "厘米";
		}
		text = text + "，段落行距" + linespacing.ToString("0.0") + "行";
		if (beforespacing != 0f)
		{
			text = text + "，段前" + beforespacing.ToString("0.0") + "行";
		}
		if (beforebreak)
		{
			text += "，段前分页";
		}
		if (afterspacing != 0f)
		{
			text = text + "，段后" + afterspacing.ToString("0.0") + "行";
		}
		return text;
	}

	public void SetStyle(Document document)
	{
		List<WdListNumberStyle> list = new List<WdListNumberStyle>(10)
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
		List<WdParagraphAlignment> list2 = new List<WdParagraphAlignment>(5)
		{
			WdParagraphAlignment.wdAlignParagraphLeft,
			WdParagraphAlignment.wdAlignParagraphCenter,
			WdParagraphAlignment.wdAlignParagraphRight,
			WdParagraphAlignment.wdAlignParagraphJustify,
			WdParagraphAlignment.wdAlignParagraphDistribute
		};
		Style style;
		try
		{
			Styles styles = document.Styles;
			object Index = name;
			style = styles[ref Index];
		}
		catch
		{
			Styles styles2 = document.Styles;
			string text = name;
			object Index = WdStyleType.wdStyleTypeParagraph;
			style = styles2.Add(text, ref Index);
		}
		if (name != "正文")
		{
			Style style2 = style;
			object Index = "";
			style2.set_BaseStyle(ref Index);
			Style style3 = style;
			Index = WdBuiltinStyle.wdStyleNormal;
			style3.set_NextParagraphStyle(ref Index);
		}
		style.Font.Name = fontname;
		style.Font.Size = fontsize;
		style.Font.Bold = (bold ? (-1) : 0);
		style.Font.Italic = (italic ? (-1) : 0);
		style.Font.Underline = (underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
		if (Regex.IsMatch(name, "标题 [1-9]|列表段落"))
		{
			if (numberstyle != 0)
			{
				ListTemplates listTemplates = Globals.ThisAddIn.Application.ListGalleries[WdListGalleryType.wdNumberGallery].ListTemplates;
				object Index = 1;
				ListTemplate listTemplate = listTemplates[ref Index];
				listTemplate.ListLevels[1].NumberStyle = list[numberstyle - 1];
				listTemplate.ListLevels[1].NumberFormat = numberformat;
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
		ParagraphFormat obj2 = (ParagraphFormat)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209F4-0000-0000-C000-000000000046")));
		obj2.Alignment = list2[paraalignment];
		obj2.LeftIndent = leftindent;
		ParagraphFormat paragraphFormat = obj2;
		if (firstlineindentbychar != 0)
		{
			paragraphFormat.IndentFirstLineCharWidth((short)firstlineindentbychar);
		}
		else
		{
			paragraphFormat.FirstLineIndent = firstlineindent;
		}
		if (linespacing == 1f)
		{
			paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			paragraphFormat.Space1();
		}
		else
		{
			paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
			paragraphFormat.LineSpacing = Globals.ThisAddIn.Application.LinesToPoints(linespacing);
		}
		paragraphFormat.PageBreakBefore = (beforebreak ? (-1) : 0);
		paragraphFormat.SpaceBefore = Globals.ThisAddIn.Application.LinesToPoints(beforespacing);
		paragraphFormat.SpaceAfter = Globals.ThisAddIn.Application.LinesToPoints(afterspacing);
		style.ParagraphFormat = paragraphFormat;
		style.QuickStyle = true;
	}
}
}