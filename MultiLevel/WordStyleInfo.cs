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

namespace WordMan_VSTO.MultiLevel
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
                    LeftIndent = "0.0 字符",
                    RightIndent = "0.0 字符",
                    FirstLineIndent = "0.0 磅",
                    LineSpace = "单倍行距",
                    SpaceBefore = "0.0 磅",
                    SpaceAfter = "0.0 磅",
                    BreakBefore = false,
                    NumberStyle = -1,
                    NumberFormat = ""
                };
            }
        }

        private WdBuiltinStyle buildInName;

        public static readonly string[] HAlignments = new string[5] { "左对齐", "中对齐", "右对齐", "两端对齐", "分散对齐" };

        public static readonly string[] LineSpacings = new string[4] { "单倍行距", "1.5倍行距", "双倍行距", "多倍行距" };

        public static readonly string[] SpaceBeforeValues = new string[6] { "0.0 磅", "6.0 磅", "12.0 磅", "18.0 磅", "24.0 磅", "36.0 磅" };

        public static readonly string[] SpaceAfterValues = new string[6] { "0.0 磅", "6.0 磅", "12.0 磅", "18.0 磅", "24.0 磅", "36.0 磅" };

        public static readonly string[] FontSizes = new string[32]
        {
            // 中文字号
            "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
            "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
            // 磅值字号
            "8 磅", "9 磅", "10 磅", "11 磅", "12 磅", "14 磅", "16 磅", "18 磅",
            "20 磅", "22 磅", "24 磅", "26 磅", "28 磅", "32 磅", "36 磅", "48 磅"
        };

        public static readonly WdBuiltinStyle[] BuildInStyleNames = new WdBuiltinStyle[9]
        {
            WdBuiltinStyle.wdStyleHeading1,
            WdBuiltinStyle.wdStyleHeading2,
            WdBuiltinStyle.wdStyleHeading3,
            WdBuiltinStyle.wdStyleHeading4,
            WdBuiltinStyle.wdStyleHeading5,
            WdBuiltinStyle.wdStyleHeading6,
            WdBuiltinStyle.wdStyleHeading7,
            WdBuiltinStyle.wdStyleHeading8,
            WdBuiltinStyle.wdStyleHeading9
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
            FontSize = MultiLevelDataManager.ConvertFontSizeToString(style.Font.Size);
            Bold = style.Font.Bold == -1;
            Italic = style.Font.Italic == -1;
            Underline = style.Font.Underline != WdUnderline.wdUnderlineNone;
            FontColor = GetWordFontColor(style.Font, style);

            // 将磅转换为字符单位（1字符 ≈ 0.5厘米 ≈ 14.175磅）
            LeftIndent = (style.ParagraphFormat.LeftIndent / 14.175f).ToString("0.0 字符");
            RightIndent = (style.ParagraphFormat.RightIndent / 14.175f).ToString("0.0 字符");

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
                    // 保存具体的倍数值
                    float multipleValue = MultiLevelDataManager.PointsToLines(style.ParagraphFormat.LineSpacing);
                    LineSpace = $"{multipleValue:0.0} 倍行距";
                    break;
                case WdLineSpacing.wdLineSpaceExactly:
                    // 固定行距，保存磅值
                    LineSpace = $"{style.ParagraphFormat.LineSpacing:0.0} 磅";
                    break;
                default:
                    LineSpace = "单倍行距";
                    break;
            }

            // 直接使用磅值，不转换为行
            SpaceBefore = style.ParagraphFormat.SpaceBefore.ToString("0.0 磅");
            SpaceAfter = style.ParagraphFormat.SpaceAfter.ToString("0.0 磅");
            FirstLineIndent = style.ParagraphFormat.FirstLineIndent.ToString("0.0 磅");

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

        /// <summary>
        /// 正确获取Word字体颜色
        /// </summary>
        private Color GetWordFontColor(Microsoft.Office.Interop.Word.Font font, Style style = null)
        {
            try
            {
                // 检查是否是自动颜色
                if (font.Color == WdColor.wdColorAutomatic)
                {
                    return Color.Black;
                }
                
                // 使用ColorTranslator.FromOle方法，这是处理Word颜色的标准方法
                // 参考StyleSettings.cs中的成功实现
                return ColorTranslator.FromOle((int)font.Color);
            }
            catch
            {
                return Color.Black;
            }
        }

        private Color ColorFromInt(int colorInt)
        {
            // 处理Word的特殊颜色值
            switch (colorInt)
            {
                case -16777216: // 黑色
                    return Color.Black;
                case -16777215: // 白色
                    return Color.White;
                case -16711681: // 红色
                    return Color.Red;
                case -16776961: // 蓝色
                    return Color.Blue;
                case -16711936: // 绿色
                    return Color.Green;
                case -256: // 黄色
                    return Color.Yellow;
                case -65281: // 洋红色
                    return Color.Magenta;
                case -16711680: // 青色
                    return Color.Cyan;
                case -16777214: // 自动颜色（通常显示为黑色）
                    return Color.Black;
                case 0: // 透明或未设置
                    return Color.Black;
                default:
                    // 对于其他颜色值，正确处理Word的BGR格式
                    if (colorInt < 0)
                    {
                        // 处理负数颜色值，Word使用BGR格式
                        // 将负数转换为正数进行处理
                        uint unsignedColor = (uint)(colorInt & 0xFFFFFF);
                        int b = (int)((unsignedColor >> 16) & 0xFF);
                        int g = (int)((unsignedColor >> 8) & 0xFF);
                        int r = (int)(unsignedColor & 0xFF);
                        return Color.FromArgb(r, g, b);
                    }
                    else
                    {
                        // 处理正数颜色值，Word使用BGR格式
                        int b = (colorInt >> 16) & 0xFF;
                        int g = (colorInt >> 8) & 0xFF;
                        int r = colorInt & 0xFF;
                        return Color.FromArgb(r, g, b);
                    }
            }
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
                style.Font.Size = MultiLevelDataManager.ConvertFontSize(FontSize);
                style.Font.Bold = Bold ? -1 : 0;
                style.Font.Italic = Italic ? -1 : 0;
                style.Font.Underline = Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                // 设置字体颜色 - 使用安全的 RGB 方法
                try
                {
                    // 将 System.Drawing.Color 转换为 Word RGB 格式
                    int r = FontColor.R;
                    int g = FontColor.G;
                    int b = FontColor.B;
                    int wordRgb = (b << 16) | (g << 8) | r; // Word 使用 BGR 格式
                    style.Font.Color = (WdColor)wordRgb;
                }
                catch
                {
                    // 如果设置失败，使用自动颜色
                    style.Font.Color = WdColor.wdColorAutomatic;
                }

                // 设置段落格式 - 直接使用样式的段落格式
                ParagraphFormat paragraphFormat = style.ParagraphFormat;
                
                // 设置对齐方式
                switch (HAlignment)
                {
                    case "左对齐":
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                    case "中对齐":
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        break;
                    case "右对齐":
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        break;
                    case "两端对齐":
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                        break;
                    case "分散对齐":
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute;
                        break;
                    default:
                        paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        break;
                }
                
                // 设置缩进
                if (LeftIndent.EndsWith("字符"))
                {
                    string s = LeftIndent.TrimEnd(' ', '字', '符');
                    if (float.TryParse(s, out float result))
                    {
                        // 字符转磅：1字符 ≈ 14.175磅
                        paragraphFormat.LeftIndent = result * 14.175f;
                    }
                }
                else if (LeftIndent.EndsWith("厘米"))
                {
                    string s = LeftIndent.TrimEnd(' ', '厘', '米');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.LeftIndent = MultiLevelDataManager.CentimetersToPoints(result);
                    }
                }
                
                if (RightIndent.EndsWith("字符"))
                {
                    string s = RightIndent.TrimEnd(' ', '字', '符');
                    if (float.TryParse(s, out float result))
                    {
                        // 字符转磅：1字符 ≈ 14.175磅
                        paragraphFormat.RightIndent = result * 14.175f;
                    }
                }
                else if (RightIndent.EndsWith("厘米"))
                {
                    string s = RightIndent.TrimEnd(' ', '厘', '米');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.RightIndent = MultiLevelDataManager.CentimetersToPoints(result);
                    }
                }
                if (FirstLineIndent.EndsWith("磅"))
                {
                    string s = FirstLineIndent.TrimEnd(' ', '磅');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.FirstLineIndent = result;
                    }
                }

                // 设置行距
                if (LineSpace == "单倍行距")
                {
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                }
                else if (LineSpace == "1.5倍行距")
                {
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                }
                else if (LineSpace == "双倍行距")
                {
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                }
                else if (LineSpace.EndsWith("倍行距"))
                {
                    // 处理多倍行距，提取倍数值
                    string valueText = LineSpace.Replace("倍行距", "").Trim();
                    if (float.TryParse(valueText, out float multipleValue))
                    {
                        paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                        paragraphFormat.LineSpacing = MultiLevelDataManager.LinesToPoints(multipleValue);
                    }
                }
                else if (LineSpace.EndsWith("磅"))
                {
                    // 处理固定行距
                    string valueText = LineSpace.Replace("磅", "").Trim();
                    if (float.TryParse(valueText, out float exactValue))
                    {
                        paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                        paragraphFormat.LineSpacing = exactValue;
                    }
                }
                else
                {
                    // 默认单倍行距
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                }

                // 设置段前段后间距
                if (SpaceBefore.EndsWith("磅"))
                {
                    string s = SpaceBefore.TrimEnd(' ', '磅');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.SpaceBefore = result;
                    }
                }
                else if (SpaceBefore.EndsWith("行"))
                {
                    string s = SpaceBefore.TrimEnd(' ', '行');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.SpaceBefore = MultiLevelDataManager.LinesToPoints(result);
                    }
                }
                
                if (SpaceAfter.EndsWith("磅"))
                {
                    string s = SpaceAfter.TrimEnd(' ', '磅');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.SpaceAfter = result;
                    }
                }
                else if (SpaceAfter.EndsWith("行"))
                {
                    string s = SpaceAfter.TrimEnd(' ', '行');
                    if (float.TryParse(s, out float result))
                    {
                        paragraphFormat.SpaceAfter = MultiLevelDataManager.LinesToPoints(result);
                    }
                }

                // 设置分页
                paragraphFormat.PageBreakBefore = BreakBefore ? -1 : 0;

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"设置样式失败：{ex.Message}");
            }
        }
    }
}
