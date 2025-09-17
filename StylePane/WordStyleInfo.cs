using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace WordMan_VSTO.StylePane
{
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

            public static StyleParaValues CreateDefault()
            {
                return new StyleParaValues
                {
                    ChnFontName = "宋体",
                    EngFontName = "宋体",
                    FontSize = "五号",
                    FontColor = Color.Black,
                    Bold = false,
                    Italic = false,
                    Underline = false,
                    HAlignment = "左对齐",
                    LeftIndent = "0.00 厘米",
                    RightIndent = "0.00 厘米",
                    FirstLineIndent = "0.00 磅",
                    LineSpace = "单倍行距",
                    SpaceBefore = "0.00 行",
                    SpaceAfter = "0.00 行",
                    BreakBefore = false,
                    NumberStyle = -1,
                    NumberFormat = ""
                };
            }
        }

        private WdBuiltinStyle buildInName;

        public static readonly string[] HAlignments = new string[5] { "左对齐", "中对齐", "右对齐", "两端对齐", "分散对齐" };

        public static readonly string[] LineSpacings = new string[4] { "单倍行距", "1.5倍行距", "双倍行距", "多倍行距" };

        public static readonly string[] SpaceBeforeValues = new string[6] { "0.00 行", "0.50 行", "1.00 行", "1.50 行", "2.00 行", "3.00 行" };

        public static readonly string[] SpaceAfterValues = new string[6] { "0.00 行", "0.50 行", "1.00 行", "1.50 行", "2.00 行", "3.00 行" };

        public static readonly string[] FontSizes = new string[16]
        {
            "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
            "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号"
        };

        public static readonly WdBuiltinStyle[] BuildInStyleNames = new WdBuiltinStyle[10]
        {
            WdBuiltinStyle.wdStyleHeading1,
            WdBuiltinStyle.wdStyleHeading2,
            WdBuiltinStyle.wdStyleHeading3,
            WdBuiltinStyle.wdStyleHeading4,
            WdBuiltinStyle.wdStyleHeading5,
            WdBuiltinStyle.wdStyleHeading6,
            WdBuiltinStyle.wdStyleHeading7,
            WdBuiltinStyle.wdStyleHeading8,
            WdBuiltinStyle.wdStyleHeading9,
            WdBuiltinStyle.wdStyleNormal
        };

        public static readonly WdListNumberStyle[] ListNumberStyles = new WdListNumberStyle[7]
        {
            WdListNumberStyle.wdListNumberStyleArabic,
            WdListNumberStyle.wdListNumberStyleUppercaseRoman,
            WdListNumberStyle.wdListNumberStyleLowercaseRoman,
            WdListNumberStyle.wdListNumberStyleUppercaseLetter,
            WdListNumberStyle.wdListNumberStyleLowercaseLetter,
            WdListNumberStyle.wdListNumberStyleCardinalText,
            WdListNumberStyle.wdListNumberStyleOrdinalText
        };

        public string StyleName { get; set; }

        public bool BuildInStyle { get; set; }

        public string ChnFontName { get; set; }

        public string EngFontName { get; set; }

        public string FontSize { get; set; }

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public bool Underline { get; set; }

        public Color FontColor { get; set; }

        public string HAlignment { get; set; }

        public string LeftIndent { get; set; }

        public string RightIndent { get; set; }

        public string FirstLineIndent { get; set; }

        public string LineSpace { get; set; }

        public string SpaceBefore { get; set; }

        public string SpaceAfter { get; set; }

        public bool BreakBefore { get; set; }

        public int NumberStyle { get; set; }

        public string NumberFormat { get; set; }

        public WordStyleInfo(Style style, WdBuiltinStyle builtinStyle)
        {
            if (!BuildInStyleNames.Contains(builtinStyle))
            {
                throw new InvalidEnumArgumentException("builtinStyle");
            }
            StyleName = style.NameLocal;
            buildInName = builtinStyle;
            BuildInStyle = true;
            ChnFontName = style.Font.NameFarEast;
            EngFontName = style.Font.NameAscii;
            FontSize = WordAPIHelper.ConvertFontSizeToString(style.Font.Size);
            Bold = style.Font.Bold == -1;
            Italic = style.Font.Italic == -1;
            Underline = style.Font.Underline != WdUnderline.wdUnderlineNone;
            FontColor = ColorFromInt((int)style.Font.Color);

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
                case WdLineSpacing.wdLineSpaceMultiple:
                    LineSpace = "多倍行距";
                    break;
                default:
                    LineSpace = "单倍行距";
                    break;
            }

            SpaceBefore = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.SpaceBefore).ToString("0.00 行");
            SpaceAfter = Globals.ThisAddIn.Application.PointsToLines(style.ParagraphFormat.SpaceAfter).ToString("0.00 行");
            FirstLineIndent = style.ParagraphFormat.FirstLineIndent.ToString("0.00 磅");

            switch (style.ParagraphFormat.Alignment)
            {
                case WdParagraphAlignment.wdAlignParagraphLeft:
                    HAlignment = "左对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphCenter:
                    HAlignment = "中对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphRight:
                    HAlignment = "右对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphJustify:
                    HAlignment = "两端对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphDistribute:
                    HAlignment = "分散对齐";
                    break;
                default:
                    HAlignment = "左对齐";
                    break;
            }

            BreakBefore = style.ParagraphFormat.PageBreakBefore == -1;
            NumberStyle = -1;
            NumberFormat = "";
        }

        public WordStyleInfo(string styleName, StyleParaValues para = default)
        {
            StyleName = styleName;
            buildInName = (WdBuiltinStyle)0;
            BuildInStyle = false;
            SetStyleValue(para);
        }

        public WordStyleInfo(WdBuiltinStyle builtinStyle, StyleParaValues para = default)
        {
            if (!BuildInStyleNames.Contains(builtinStyle))
            {
                throw new InvalidEnumArgumentException("builtinStyle");
            }
            StyleName = builtinStyle.ToString();
            buildInName = builtinStyle;
            BuildInStyle = true;
            SetStyleValue(para);
        }

        public WordStyleInfo()
        {
        }

        public static bool operator ==(WordStyleInfo left, WordStyleInfo right)
        {
            if (left == null && right == null)
            {
                return true;
            }
            if (left == null || right == null)
            {
                return false;
            }
            return left.StyleName == right.StyleName && left.BuildInStyle == right.BuildInStyle;
        }

        public static bool operator !=(WordStyleInfo left, WordStyleInfo right)
        {
            return !(left == right);
        }

        public override bool Equals(object obj)
        {
            if (obj is WordStyleInfo)
            {
                return this == (WordStyleInfo)obj;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return StyleName.GetHashCode() ^ BuildInStyle.GetHashCode();
        }

        public void SetStyleValue(StyleParaValues para)
        {
            if (para.Equals(default(StyleParaValues)))
            {
                para = StyleParaValues.CreateDefault();
            }
            
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
                fontStyle |= FontStyle.Bold;
                text += "加粗；";
            }
            if (Italic)
            {
                fontStyle |= FontStyle.Italic;
                text += "斜体；";
            }
            if (Underline)
            {
                fontStyle |= FontStyle.Underline;
                text += "下划线；";
            }
            text = text + "颜色：" + FontColor.Name + "；";
            text = text + "段落行距：" + LineSpace + "；段前：" + SpaceBefore + "；段后：" + SpaceAfter + "；";
            if (BreakBefore)
            {
                text += "段前分行；";
            }
            font = new System.Drawing.Font(new FontFamily(ChnFontName), 10.5f, fontStyle);
            return text;
        }

        private Color ColorFromInt(int colorInt)
        {
            if (colorInt == -16777216)
            {
                return Color.Black;
            }
            if (colorInt == -16777215)
            {
                return Color.White;
            }
            if (colorInt == -16711681)
            {
                return Color.Red;
            }
            if (colorInt == -16776961)
            {
                return Color.Blue;
            }
            if (colorInt == -16711936)
            {
                return Color.Green;
            }
            if (colorInt == -256)
            {
                return Color.Yellow;
            }
            if (colorInt == -65281)
            {
                return Color.Magenta;
            }
            if (colorInt == -16711681)
            {
                return Color.Cyan;
            }
            return Color.FromArgb(colorInt);
        }

        private MsoThemeColorSchemeIndex GetThemeColorIndex(WdThemeColorIndex index)
        {
            switch (index)
            {
                case WdThemeColorIndex.wdThemeColorMainLight1:
                    return MsoThemeColorSchemeIndex.msoThemeLight1;
                case WdThemeColorIndex.wdThemeColorMainDark2:
                    return MsoThemeColorSchemeIndex.msoThemeDark2;
                case WdThemeColorIndex.wdThemeColorMainLight2:
                    return MsoThemeColorSchemeIndex.msoThemeLight2;
                case WdThemeColorIndex.wdThemeColorAccent1:
                    return MsoThemeColorSchemeIndex.msoThemeAccent1;
                case WdThemeColorIndex.wdThemeColorAccent2:
                    return MsoThemeColorSchemeIndex.msoThemeAccent2;
                case WdThemeColorIndex.wdThemeColorAccent3:
                    return MsoThemeColorSchemeIndex.msoThemeAccent3;
                case WdThemeColorIndex.wdThemeColorAccent4:
                    return MsoThemeColorSchemeIndex.msoThemeAccent4;
                case WdThemeColorIndex.wdThemeColorAccent5:
                    return MsoThemeColorSchemeIndex.msoThemeAccent5;
                case WdThemeColorIndex.wdThemeColorAccent6:
                    return MsoThemeColorSchemeIndex.msoThemeAccent6;
                case WdThemeColorIndex.wdThemeColorHyperlink:
                    return MsoThemeColorSchemeIndex.msoThemeHyperlink;
                case WdThemeColorIndex.wdThemeColorHyperlinkFollowed:
                    return MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink;
                default:
                    return MsoThemeColorSchemeIndex.msoThemeDark1;
            }
        }

        public bool SetStyle(Document docForApply)
        {
            try
            {
                Style style = docForApply.Styles[StyleName];
                if (style == null)
                {
                    return false;
                }

                // 设置字体
                style.Font.NameFarEast = ChnFontName;
                style.Font.NameAscii = EngFontName;
                style.Font.Size = WordAPIHelper.ConvertFontSize(FontSize);
                style.Font.Bold = Bold ? -1 : 0;
                style.Font.Italic = Italic ? -1 : 0;
                style.Font.Underline = Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                style.Font.Color = (WdColor)FontColor.ToArgb();

                // 设置段落格式
                ParagraphFormat paragraphFormat = (ParagraphFormat)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("000209F4-0000-0000-C000-000000000046")));
                ParagraphFormat paragraphFormat2 = paragraphFormat;
                switch (HAlignment)
                {
                    case "左对齐":
                        paragraphFormat2.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                    case "中对齐":
                        paragraphFormat2.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        break;
                    case "右对齐":
                        paragraphFormat2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        break;
                    case "两端对齐":
                        paragraphFormat2.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                        break;
                    case "分散对齐":
                        paragraphFormat2.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute;
                        break;
                    default:
                        paragraphFormat2.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                }

                ParagraphFormat paragraphFormat3 = paragraphFormat;
                if (LeftIndent.EndsWith("厘米"))
                {
                    string s = LeftIndent.TrimEnd(' ', '厘', '米');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat3.LeftIndent = Globals.ThisAddIn.Application.CentimetersToPoints(result);
                    }
                }
                if (RightIndent.EndsWith("厘米"))
                {
                    string s = RightIndent.TrimEnd(' ', '厘', '米');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat3.RightIndent = Globals.ThisAddIn.Application.CentimetersToPoints(result);
                    }
                }
                if (FirstLineIndent.EndsWith("磅"))
                {
                    string s = FirstLineIndent.TrimEnd(' ', '磅');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat3.FirstLineIndent = result;
                    }
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
                    case "多倍行距":
                        paragraphFormat3.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                        paragraphFormat3.LineSpacing = 1.5f;
                        break;
                }

                if (SpaceBefore.EndsWith("行"))
                {
                    string s = SpaceBefore.TrimEnd(' ', '行');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat3.SpaceBefore = Globals.ThisAddIn.Application.LinesToPoints(result);
                    }
                }
                if (SpaceAfter.EndsWith("行"))
                {
                    string s = SpaceAfter.TrimEnd(' ', '行');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat3.SpaceAfter = Globals.ThisAddIn.Application.LinesToPoints(result);
                    }
                }

                paragraphFormat3.PageBreakBefore = BreakBefore ? -1 : 0;

                // 应用段落格式
                object Index = Type.Missing;
                style.ParagraphFormat.Alignment = paragraphFormat2.Alignment;
                style.ParagraphFormat.LeftIndent = paragraphFormat3.LeftIndent;
                style.ParagraphFormat.RightIndent = paragraphFormat3.RightIndent;
                style.ParagraphFormat.FirstLineIndent = paragraphFormat3.FirstLineIndent;
                style.ParagraphFormat.LineSpacingRule = paragraphFormat3.LineSpacingRule;
                style.ParagraphFormat.LineSpacing = paragraphFormat3.LineSpacing;
                style.ParagraphFormat.SpaceBefore = paragraphFormat3.SpaceBefore;
                style.ParagraphFormat.SpaceAfter = paragraphFormat3.SpaceAfter;
                style.ParagraphFormat.PageBreakBefore = paragraphFormat3.PageBreakBefore;

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"设置样式失败：{ex.Message}");
            }
        }
    }
}
