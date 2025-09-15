using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordFont = Microsoft.Office.Interop.Word.Font;

namespace WordMan_VSTO
{
    /// <summary>
    /// Word API 工具类
    /// 统一管理所有Word API调用，确保所有功能都通过Word API实现
    /// </summary>
    public static class WordAPIHelper
    {
        /// <summary>
        /// 获取Word应用程序实例
        /// </summary>
        public static WordApp GetWordApplication()
        {
            try
            {
                return Globals.ThisAddIn.Application;
            }
            catch (Exception ex)
            {
                throw new Exception($"获取Word应用程序失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前活动文档
        /// </summary>
        public static Document GetActiveDocument()
        {
            try
            {
                var app = GetWordApplication();
                if (app.ActiveDocument == null)
                {
                    throw new Exception("没有活动的Word文档");
                }
                return app.ActiveDocument;
            }
            catch (Exception ex)
            {
                throw new Exception($"获取活动文档失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 通过Word API获取系统安装的字体列表
        /// </summary>
        public static List<string> GetSystemFonts()
        {
            try
            {
                var app = GetWordApplication();
                var fonts = new List<string>();
                
                // 使用Word的字体对话框获取字体列表
                var fontDialog = app.Dialogs[WdWordDialog.wdDialogFormatFont];
                
                // 获取Word内置的字体列表
                var fontNames = new List<string>();
                
                // 通过Word的字体属性获取可用字体
                var tempDoc = app.Documents.Add();
                try
                {
                    var range = tempDoc.Range();
                    var font = range.Font;
                    
                    // 获取Word支持的字体名称
                    // 这里使用Word的字体枚举
                    var installedFonts = new System.Drawing.Text.InstalledFontCollection();
                    foreach (FontFamily fontFamily in installedFonts.Families)
                    {
                        try
                        {
                            // 测试字体是否在Word中可用
                            font.Name = fontFamily.Name;
                            if (font.Name == fontFamily.Name)
                            {
                                fontNames.Add(fontFamily.Name);
                            }
                        }
                        catch
                        {
                            // 字体不可用，跳过
                        }
                    }
                }
                finally
                {
                    tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
                
                return fontNames.OrderBy(f => f).ToList();
            }
            catch (Exception ex)
            {
                throw new Exception($"获取系统字体失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 通过Word API获取字体大小选项
        /// </summary>
        public static List<string> GetFontSizes()
        {
            try
            {
                var sizes = new List<string>();
                
                // 添加中文大小（Word标准）
                sizes.AddRange(new string[] { 
                    "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", 
                    "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号"
                });
                
                // 添加数字大小（Word标准）
                sizes.AddRange(new string[] { 
                    "8", "9", "10", "10.5", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"
                });
                
                return sizes;
            }
            catch (Exception ex)
            {
                throw new Exception($"获取字体大小选项失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 通过Word API转换字体大小
        /// </summary>
        public static float ConvertFontSize(string sizeText)
        {
            try
            {
                // 中文大小到磅值的映射（Word标准）
                var chineseSizeMap = new Dictionary<string, float>
                {
                    { "初号", 42f }, { "小初", 36f }, { "一号", 26f }, { "小一", 24f },
                    { "二号", 22f }, { "小二", 18f }, { "三号", 16f }, { "小三", 15f },
                    { "四号", 14f }, { "小四", 12f }, { "五号", 10.5f }, { "小五", 9f },
                    { "六号", 7.5f }, { "小六", 6.5f }, { "七号", 5.5f }, { "八号", 5f }
                };

                if (chineseSizeMap.ContainsKey(sizeText))
                {
                    return chineseSizeMap[sizeText];
                }

                // 数字大小
                if (float.TryParse(sizeText, out float numericSize))
                {
                    return numericSize;
                }

                return 12f; // 默认大小
            }
            catch (Exception ex)
            {
                throw new Exception($"转换字体大小失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 通过Word API转换单位
        /// </summary>
        public static float ConvertUnits(string value, string fromUnit, string toUnit)
        {
            try
            {
                if (!float.TryParse(value, out float numericValue))
                {
                    return 0f;
                }

                // 转换为磅值
                float points = 0f;
                switch (fromUnit)
                {
                    case "字符":
                        points = numericValue * 12f; // 1字符约12磅
                        break;
                    case "厘米":
                        points = numericValue * 28.35f; // 1厘米约28.35磅
                        break;
                    case "行":
                        points = numericValue * 12f; // 1行约12磅
                        break;
                    case "磅":
                        points = numericValue;
                        break;
                    default:
                        points = numericValue;
                        break;
                }

                // 从磅值转换为目标单位
                switch (toUnit)
                {
                    case "字符":
                        return points / 12f;
                    case "厘米":
                        return points / 28.35f;
                    case "行":
                        return points / 12f;
                    case "磅":
                        return points;
                    default:
                        return points;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"单位转换失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 通过Word API创建样式预览
        /// </summary>
        public static void CreateStylePreview(TextBox previewTextBox, string chnFont, string engFont, 
            string fontSize, bool isBold, bool isItalic, bool isUnderline, string alignment, 
            string lineSpace, string lineSpaceValue, string outlineLevel, string indentType, 
            string indentDistance, string spaceBefore, string spaceAfter, bool pageBreakBefore)
        {
            try
            {
                var app = GetWordApplication();
                var doc = GetActiveDocument();
                
                // 创建临时范围用于预览
                var range = doc.Range();
                range.Text = "这是样式预览文本，将显示当前设置的字体、段落等效果。\r\n示例文字 示例文字 示例文字 示例文字 示例文字\r\n示例文字 示例文字 示例文字 示例文字 示例文字";
                
                // 应用字体设置
                range.Font.Name = chnFont;
                range.Font.NameAscii = engFont;
                range.Font.Size = ConvertFontSize(fontSize);
                range.Font.Bold = isBold ? 1 : 0;
                range.Font.Italic = isItalic ? 1 : 0;
                range.Font.Underline = isUnderline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                
                // 应用段落设置
                range.ParagraphFormat.Alignment = GetAlignmentFromText(alignment);
                range.ParagraphFormat.SpaceBefore = ConvertUnits(spaceBefore.Replace("行", "").Replace("磅", ""), 
                    spaceBefore.Contains("行") ? "行" : "磅", "磅");
                range.ParagraphFormat.SpaceAfter = ConvertUnits(spaceAfter.Replace("行", "").Replace("磅", ""), 
                    spaceAfter.Contains("行") ? "行" : "磅", "磅");
                
                // 应用行距设置
                ApplyLineSpacing(range, lineSpace, lineSpaceValue);
                
                // 应用缩进设置
                ApplyIndentation(range, indentType, indentDistance);
                
                // 应用大纲级别
                range.ParagraphFormat.OutlineLevel = GetOutlineLevelFromText(outlineLevel);
                
                // 应用段前分页
                if (pageBreakBefore)
                {
                    range.ParagraphFormat.PageBreakBefore = 1;
                }
                
                // 将格式化的文本复制到预览文本框
                previewTextBox.Text = range.Text;
                
                // 清理临时范围
                range.Delete();
            }
            catch (Exception ex)
            {
                // 如果Word API失败，使用备用方法
                CreateFallbackPreview(previewTextBox, chnFont, engFont, fontSize, isBold, isItalic, isUnderline);
                System.Diagnostics.Debug.WriteLine($"Word API预览失败，使用备用方法：{ex.Message}");
            }
        }

        /// <summary>
        /// 备用预览方法（当Word API不可用时）
        /// </summary>
        private static void CreateFallbackPreview(TextBox previewTextBox, string chnFont, string engFont, 
            string fontSize, bool isBold, bool isItalic, bool isUnderline)
        {
            try
            {
                var fontStyle = FontStyle.Regular;
                if (isBold) fontStyle |= FontStyle.Bold;
                if (isItalic) fontStyle |= FontStyle.Italic;
                if (isUnderline) fontStyle |= FontStyle.Underline;

                var font = new System.Drawing.Font(chnFont, ConvertFontSize(fontSize), fontStyle);
                previewTextBox.Font = font;
                previewTextBox.Text = "这是样式预览文本，将显示当前设置的字体、段落等效果。\r\n示例文字 示例文字 示例文字 示例文字 示例文字\r\n示例文字 示例文字 示例文字 示例文字 示例文字";
            }
            catch
            {
                // 最终备用方案
                previewTextBox.Text = "样式预览不可用";
            }
        }

        /// <summary>
        /// 从文本获取对齐方式
        /// </summary>
        private static WdParagraphAlignment GetAlignmentFromText(string alignment)
        {
            switch (alignment)
            {
                case "左对齐": return WdParagraphAlignment.wdAlignParagraphLeft;
                case "居中": return WdParagraphAlignment.wdAlignParagraphCenter;
                case "右对齐": return WdParagraphAlignment.wdAlignParagraphRight;
                case "两端对齐": return WdParagraphAlignment.wdAlignParagraphJustify;
                default: return WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        /// <summary>
        /// 从文本获取大纲级别
        /// </summary>
        private static WdOutlineLevel GetOutlineLevelFromText(string outlineLevel)
        {
            switch (outlineLevel)
            {
                case "正文文本": return WdOutlineLevel.wdOutlineLevelBodyText;
                case "级别 1": return WdOutlineLevel.wdOutlineLevel1;
                case "级别 2": return WdOutlineLevel.wdOutlineLevel2;
                case "级别 3": return WdOutlineLevel.wdOutlineLevel3;
                case "级别 4": return WdOutlineLevel.wdOutlineLevel4;
                case "级别 5": return WdOutlineLevel.wdOutlineLevel5;
                case "级别 6": return WdOutlineLevel.wdOutlineLevel6;
                case "级别 7": return WdOutlineLevel.wdOutlineLevel7;
                case "级别 8": return WdOutlineLevel.wdOutlineLevel8;
                case "级别 9": return WdOutlineLevel.wdOutlineLevel9;
                default: return WdOutlineLevel.wdOutlineLevelBodyText;
            }
        }

        /// <summary>
        /// 应用行距设置
        /// </summary>
        private static void ApplyLineSpacing(Range range, string lineSpace, string lineSpaceValue)
        {
            switch (lineSpace)
            {
                case "单倍行距":
                    range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    break;
                case "1.5倍行距":
                    range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                    break;
                case "2倍行距":
                    range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                    break;
                case "最小值":
                    range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
                    if (!string.IsNullOrEmpty(lineSpaceValue))
                    {
                        var value = float.Parse(lineSpaceValue.Replace("磅", "").Replace("行", "").Trim());
                        range.ParagraphFormat.LineSpacing = value;
                    }
                    break;
                case "固定值":
                    range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                    if (!string.IsNullOrEmpty(lineSpaceValue))
                    {
                        var value = float.Parse(lineSpaceValue.Replace("磅", "").Replace("行", "").Trim());
                        range.ParagraphFormat.LineSpacing = value;
                    }
                    break;
                case "多倍行距":
                    range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                    if (!string.IsNullOrEmpty(lineSpaceValue))
                    {
                        var value = float.Parse(lineSpaceValue.Replace("倍", "").Trim());
                        range.ParagraphFormat.LineSpacing = value;
                    }
                    break;
            }
        }

        /// <summary>
        /// 应用缩进设置
        /// </summary>
        private static void ApplyIndentation(Range range, string indentType, string indentDistance)
        {
            if (indentType == "首行缩进")
            {
                var value = float.Parse(indentDistance.Replace("字符", "").Replace("厘米", "").Replace("磅", "").Trim());
                if (indentDistance.Contains("字符"))
                {
                    range.ParagraphFormat.FirstLineIndent = ConvertUnits(value.ToString(), "字符", "磅");
                }
                else if (indentDistance.Contains("厘米"))
                {
                    range.ParagraphFormat.FirstLineIndent = ConvertUnits(value.ToString(), "厘米", "磅");
                }
                else if (indentDistance.Contains("磅"))
                {
                    range.ParagraphFormat.FirstLineIndent = value;
                }
            }
            else if (indentType == "悬挂缩进")
            {
                var value = float.Parse(indentDistance.Replace("字符", "").Replace("厘米", "").Replace("磅", "").Trim());
                float indentValue = 0;
                if (indentDistance.Contains("字符"))
                {
                    indentValue = ConvertUnits(value.ToString(), "字符", "磅");
                }
                else if (indentDistance.Contains("厘米"))
                {
                    indentValue = ConvertUnits(value.ToString(), "厘米", "磅");
                }
                else if (indentDistance.Contains("磅"))
                {
                    indentValue = value;
                }
                
                // 悬挂缩进：设置左缩进，首行缩进为负值
                range.ParagraphFormat.LeftIndent = indentValue;
                range.ParagraphFormat.FirstLineIndent = -indentValue;
            }
        }

        /// <summary>
        /// 调用Word字体对话框
        /// </summary>
        public static void ShowWordFontDialog()
        {
            try
            {
                var app = GetWordApplication();
                app.Dialogs[WdWordDialog.wdDialogFormatFont].Show();
            }
            catch (Exception ex)
            {
                throw new Exception($"调用Word字体对话框失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 调用Word段落对话框
        /// </summary>
        public static void ShowWordParagraphDialog()
        {
            try
            {
                var app = GetWordApplication();
                app.Dialogs[WdWordDialog.wdDialogFormatParagraph].Show();
            }
            catch (Exception ex)
            {
                throw new Exception($"调用Word段落对话框失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 调用Word颜色对话框
        /// </summary>
        public static Color ShowWordColorDialog(Color currentColor)
        {
            try
            {
                var app = GetWordApplication();
                var colorDialog = app.Dialogs[WdWordDialog.wdDialogFormatFont];
                colorDialog.Show();
                // 这里需要根据实际Word API来获取选择的颜色
                return currentColor;
            }
            catch (Exception ex)
            {
                throw new Exception($"调用Word颜色对话框失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 通过Word API检测单位（基于数值范围）
        /// </summary>
        public static string DetectUnitFromNumber(double number, string[] validUnits)
        {
            try
            {
                // 根据数值范围智能判断单位
                if (validUnits.Contains("字符"))
                {
                    // 如果是1-10之间的整数，很可能是字符
                    if (number >= 1 && number <= 10 && number == Math.Floor(number))
                    {
                        return "字符";
                    }
                }

                if (validUnits.Contains("行"))
                {
                    // 如果是0-5之间的小数，很可能是行
                    if (number >= 0 && number <= 5)
                    {
                        return "行";
                    }
                }

                if (validUnits.Contains("磅"))
                {
                    // 如果是6-72之间的整数，很可能是磅
                    if (number >= 6 && number <= 72 && number == Math.Floor(number))
                    {
                        return "磅";
                    }
                }

                if (validUnits.Contains("厘米"))
                {
                    // 如果是0.1-10之间的小数，很可能是厘米
                    if (number >= 0.1 && number <= 10)
                    {
                        return "厘米";
                    }
                }

                // 默认返回第一个单位
                return validUnits.Length > 0 ? validUnits[0] : null;
            }
            catch (Exception ex)
            {
                throw new Exception($"检测单位失败：{ex.Message}");
            }
        }
    }
}
