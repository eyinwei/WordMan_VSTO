using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordFont = Microsoft.Office.Interop.Word.Font;
using Point = System.Drawing.Point;
using Font = System.Drawing.Font;
using Color = System.Drawing.Color;

namespace WordMan_VSTO
{
    #region 数据结构定义

    /// <summary>
    /// 级别数据结构
    /// </summary>
    public class LevelData
    {
        public int Level { get; set; }
        public string NumberStyle { get; set; }
        public string NumberFormat { get; set; }
        public decimal NumberIndent { get; set; }
        public decimal TextIndent { get; set; }
        public string AfterNumberType { get; set; } // 编号之后类型：无、空格、制表位
        public decimal TabPosition { get; set; } // 制表位位置
        public string LinkedStyle { get; set; }
    }

    /// <summary>
    /// 级别数据事件参数
    /// </summary>
    public class LevelDataEventArgs : EventArgs
    {
        public LevelData LevelData { get; set; }
        
        public LevelDataEventArgs(LevelData levelData)
        {
            LevelData = levelData;
        }
    }

    #endregion

    /// <summary>
    /// 多级列表数据管理器 - 专门处理数据转换和格式转换
    /// </summary>
    public static class MultiLevelDataManager
    {
        #region 数据转换方法

        /// <summary>
        /// 转换字体大小字符串为数值
        /// </summary>
        public static float ConvertFontSize(string fontSizeText)
        {
            if (string.IsNullOrEmpty(fontSizeText))
                return 12f;

            // 移除单位后缀
            string numberText = fontSizeText.Replace("磅", "").Replace("pt", "").Trim();
            
            // 处理中文字号
            if (fontSizeText.Contains("号"))
            {
                var chineseSizes = new Dictionary<string, float>
                {
                    {"初号", 42f}, {"小初", 36f}, {"一号", 26f}, {"小一", 24f},
                    {"二号", 22f}, {"小二", 18f}, {"三号", 16f}, {"小三", 15f},
                    {"四号", 14f}, {"小四", 12f}, {"五号", 10.5f}, {"小五", 9f},
                    {"六号", 7.5f}, {"小六", 6.5f}, {"七号", 5.5f}, {"八号", 5f}
                };
                
                if (chineseSizes.TryGetValue(fontSizeText, out float chineseSize))
                    return chineseSize;
            }
            
            // 尝试解析数值
            if (float.TryParse(numberText, out float result))
            {
                // 如果数值小于10且不包含"磅"或"pt"，可能是中文字号，需要转换
                if (result < 10 && !fontSizeText.Contains("磅") && !fontSizeText.Contains("pt"))
                {
                    var chineseToPoints = new Dictionary<float, float>
                    {
                        {1f, 42f}, {2f, 26f}, {3f, 16f}, {4f, 14f}, {5f, 10.5f},
                        {6f, 7.5f}, {7f, 5.5f}, {8f, 5f}
                    };
                    
                    if (chineseToPoints.TryGetValue(result, out float convertedSize))
                        return convertedSize;
                }
                
                return result;
            }
            
            return 12f; // 默认值
        }

        /// <summary>
        /// 转换单位
        /// </summary>
        public static float ConvertUnits(string valueText, string fromUnit, string toUnit)
        {
            if (!float.TryParse(valueText, out float value))
                return 0f;

            // 使用统一的UnitConverter进行转换
            return (float)WordMan_VSTO.UnitConverter.UnitConvert(value, fromUnit, toUnit);
        }

        /// <summary>
        /// 转换字体大小为字符串
        /// </summary>
        public static string ConvertFontSizeToString(float fontSize)
        {
            // 中文字号映射
            var pointToChinese = new Dictionary<float, string>
            {
                {42f, "初号"}, {36f, "小初"}, {26f, "一号"}, {24f, "小一"},
                {22f, "二号"}, {18f, "小二"}, {16f, "三号"}, {15f, "小三"},
                {14f, "四号"}, {12f, "小四"}, {10.5f, "五号"}, {9f, "小五"},
                {7.5f, "六号"}, {6.5f, "小六"}, {5.5f, "七号"}, {5f, "八号"}
            };
            
            // 查找最接近的中文字号
            foreach (var kvp in pointToChinese.OrderBy(x => Math.Abs(x.Key - fontSize)))
            {
                if (Math.Abs(kvp.Key - fontSize) < 0.1f)
                    return kvp.Value;
            }
            
            // 对于常见磅值，优先使用标准格式
            var commonSizes = new[] { 8f, 9f, 10f, 11f, 12f, 14f, 16f, 18f, 20f, 22f, 24f, 26f, 28f, 32f, 36f, 48f };
            if (commonSizes.Contains(fontSize))
                return fontSize.ToString("0") + " 磅";
            
            return fontSize.ToString("0") + " 磅";
        }

        #endregion

        #region 格式转换方法

        /// <summary>
        /// 生成编号格式字符串
        /// </summary>
        public static string GenerateNumberFormat(int level)
        {
            if (level <= 0) return "";
            
            var sb = new StringBuilder();
            for (int i = 1; i <= level; i++)
            {
                sb.Append("%" + i);
                if (i < level) sb.Append(".");
            }
            return sb.ToString();
        }

        /// <summary>
        /// 解析编号格式字符串
        /// </summary>
        public static List<int> ParseNumberFormat(string format)
        {
            var levels = new List<int>();
            if (string.IsNullOrEmpty(format)) return levels;
            
            var matches = System.Text.RegularExpressions.Regex.Matches(format, @"%(\d+)");
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                if (int.TryParse(match.Groups[1].Value, out int level))
                {
                    levels.Add(level);
                }
            }
            return levels;
        }

        /// <summary>
        /// 验证编号格式
        /// </summary>
        public static bool ValidateNumberFormat(string format, int maxLevel)
        {
            if (string.IsNullOrEmpty(format)) return true;
            
            var levels = ParseNumberFormat(format);
            return levels.All(level => level >= 1 && level <= maxLevel);
        }

        #endregion

        #region Word API 转换方法

        /// <summary>
        /// 厘米转磅 - 直接使用Word API
        /// </summary>
        public static float CentimetersToPoints(float centimeters)
        {
            return Globals.ThisAddIn.Application.CentimetersToPoints(centimeters);
        }

        /// <summary>
        /// 磅转厘米 - 直接使用Word API
        /// </summary>
        public static float PointsToCentimeters(float points)
        {
            return Globals.ThisAddIn.Application.PointsToCentimeters(points);
        }

        /// <summary>
        /// 磅转行 - 直接使用Word API
        /// </summary>
        public static float PointsToLines(float points)
        {
            return Globals.ThisAddIn.Application.PointsToLines(points);
        }

        /// <summary>
        /// 行转磅 - 直接使用Word API
        /// </summary>
        public static float LinesToPoints(float lines)
        {
            return Globals.ThisAddIn.Application.LinesToPoints(lines);
        }

        #endregion

        #region 数据验证方法

        /// <summary>
        /// 验证级别数据
        /// </summary>
        public static bool ValidateLevelData(LevelData levelData, out string errorMessage)
        {
            errorMessage = string.Empty;
            
            if (!MultiLevelDataManager.ValidationHelper.IsValidNumberStyle(levelData.NumberStyle))
            {
                errorMessage = $"无效的编号样式：{levelData.NumberStyle}";
                return false;
            }
            
            if (!MultiLevelDataManager.ValidationHelper.IsValidAfterNumberType(levelData.AfterNumberType))
            {
                errorMessage = $"无效的编号之后类型：{levelData.AfterNumberType}";
                return false;
            }
            
            if (!MultiLevelDataManager.ValidationHelper.IsValidLinkedStyle(levelData.LinkedStyle))
            {
                errorMessage = $"无效的链接样式：{levelData.LinkedStyle}";
                return false;
            }
            
            if (!MultiLevelDataManager.ValidationHelper.IsValidIndentValue(levelData.NumberIndent))
            {
                errorMessage = $"无效的编号缩进值：{levelData.NumberIndent}";
                return false;
            }
            
            if (!MultiLevelDataManager.ValidationHelper.IsValidIndentValue(levelData.TextIndent))
            {
                errorMessage = $"无效的文本缩进值：{levelData.TextIndent}";
                return false;
            }
            
            if (levelData.AfterNumberType == "制表位" && !MultiLevelDataManager.ValidationHelper.IsValidTabPosition(levelData.TabPosition))
            {
                errorMessage = $"无效的制表位位置：{levelData.TabPosition}";
                return false;
            }
            
            return true;
        }

        /// <summary>
        /// 批量验证级别数据列表
        /// </summary>
        public static bool ValidateLevelDataList(List<LevelData> levelDataList, out List<string> errorMessages)
        {
            errorMessages = new List<string>();
            bool isValid = true;
            
            foreach (var levelData in levelDataList)
            {
                if (!ValidateLevelData(levelData, out string errorMessage))
                {
                    errorMessages.Add($"第{levelData.Level}级：{errorMessage}");
                    isValid = false;
                }
            }
            
            return isValid;
        }

        #endregion

        #region 数据导出方法

        /// <summary>
        /// 生成级别数据摘要
        /// </summary>
        public static string GenerateLevelDataSummary(List<LevelData> levelDataList)
        {
            var summary = new StringBuilder();
            summary.AppendLine($"多级列表配置摘要（共{levelDataList.Count}级）：");
            summary.AppendLine();
            
            foreach (var levelData in levelDataList)
            {
                summary.AppendLine($"第{levelData.Level}级：");
                summary.AppendLine($"  编号样式：{levelData.NumberStyle}");
                summary.AppendLine($"  编号格式：{levelData.NumberFormat}");
                summary.AppendLine($"  编号缩进：{levelData.NumberIndent}厘米");
                summary.AppendLine($"  文本缩进：{levelData.TextIndent}厘米");
                summary.AppendLine($"  编号之后：{levelData.AfterNumberType}");
                if (levelData.AfterNumberType == "制表位")
                {
                    summary.AppendLine($"  制表位位置：{levelData.TabPosition}厘米");
                }
                summary.AppendLine($"  链接样式：{levelData.LinkedStyle}");
                summary.AppendLine();
            }
            
            return summary.ToString();
        }

        /// <summary>
        /// 导出配置到文本文件
        /// </summary>
        public static void ExportConfigurationToText(string filePath, List<LevelData> levelDataList, int currentLevels)
        {
            try
            {
                var content = new StringBuilder();
                content.AppendLine("多级列表配置导出");
                content.AppendLine($"导出时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                content.AppendLine($"配置级别数：{currentLevels}");
                content.AppendLine();
                
                content.AppendLine(GenerateLevelDataSummary(levelDataList));
                
                File.WriteAllText(filePath, content.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                throw new Exception($"导出配置到文本文件失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 从级别数据生成Word多级列表模板
        /// </summary>
        public static string GenerateWordListTemplate(List<LevelData> levelDataList)
        {
            var template = new StringBuilder();
            template.AppendLine("Word多级列表模板代码：");
            template.AppendLine();
            
            foreach (var levelData in levelDataList)
            {
                template.AppendLine($"// 第{levelData.Level}级配置");
                template.AppendLine($"listLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic; // {levelData.NumberStyle}");
                template.AppendLine($"listLevel.NumberFormat = \"{levelData.NumberFormat}\";");
                template.AppendLine($"listLevel.NumberPosition = app.CentimetersToPoints({levelData.NumberIndent}f);");
                template.AppendLine($"listLevel.TextPosition = app.CentimetersToPoints({levelData.TextIndent}f);");
                
                if (levelData.AfterNumberType == "制表位")
                {
                    template.AppendLine($"listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;");
                    template.AppendLine($"listLevel.TabPosition = app.CentimetersToPoints({levelData.TabPosition}f);");
                }
                else if (levelData.AfterNumberType == "空格")
                {
                    template.AppendLine($"listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;");
                }
                else
                {
                    template.AppendLine($"listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;");
                }
                
                if (levelData.LinkedStyle != "无")
                {
                    template.AppendLine($"// 链接到样式：{levelData.LinkedStyle}");
                }
                
                template.AppendLine();
            }
            
            return template.ToString();
        }

        #endregion

        #region Word API 辅助方法

        /// <summary>
        /// 获取系统字体列表
        /// </summary>
        public static List<string> GetSystemFonts()
        {
            var fonts = new List<string>();
            
            // 检查Word Application是否可用（设计时可能为null）
            var app = Globals.ThisAddIn?.Application;
            if (app != null)
            {
                try
                {
                    foreach (Font font in app.FontNames)
                    {
                        fonts.Add(font.Name);
                    }
                }
                catch
                {
                    // 如果Word API调用失败，使用系统字体
                    LoadSystemFonts(fonts);
                }
            }
            else
            {
                // 设计时或Word不可用时，使用系统字体
                LoadSystemFonts(fonts);
            }
            
            return fonts;
        }
        
        /// <summary>
        /// 加载系统字体（设计时使用）
        /// </summary>
        private static void LoadSystemFonts(List<string> fonts)
        {
            try
            {
                var installedFontCollection = new System.Drawing.Text.InstalledFontCollection();
                foreach (var fontFamily in installedFontCollection.Families)
                {
                    fonts.Add(fontFamily.Name);
                }
            }
            catch
            {
                // 如果连系统字体都无法获取，添加默认字体
                fonts.AddRange(new[] { "微软雅黑", "宋体", "黑体", "Arial", "Times New Roman" });
            }
        }

        /// <summary>
        /// 获取字体大小列表
        /// </summary>
        public static string[] GetFontSizes()
        {
            return new string[]
            {
                "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
                "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
                "8 磅", "9 磅", "10 磅", "11 磅", "12 磅", "14 磅", "16 磅", "18 磅", "20 磅", "22 磅", "24 磅", "26 磅", "28 磅", "36 磅", "48 磅", "72 磅"
            };
        }


        /// <summary>
        /// 创建样式预览
        /// </summary>
        public static void CreateStylePreview(TextBox textBox, string chnFont, string engFont, string fontSize, 
            bool bold, bool italic, bool underline, string alignment, string lineSpace, string lineSpaceValue, 
            string outlineLevel, string indentType, string indentDistance, string spaceBefore, string spaceAfter, bool pageBreakBefore)
        {
            try
            {
                // 设置字体
                var font = new Font(chnFont, ConvertFontSize(fontSize), 
                    (bold ? FontStyle.Bold : FontStyle.Regular) | 
                    (italic ? FontStyle.Italic : FontStyle.Regular) | 
                    (underline ? FontStyle.Underline : FontStyle.Regular));
                
                textBox.Font = font;
                textBox.Text = "样式预览\r\n样式预览\r\n样式预览\r\n样式预览\r\n样式预览\r\n\r\n样式预览\r\n样式预览\r\n样式预览\r\n样式预览\r\n样式预览";
                
                // 设置对齐方式
                switch (alignment)
                {
                    case "左对齐":
                        textBox.TextAlign = HorizontalAlignment.Left;
                        break;
                    case "居中":
                        textBox.TextAlign = HorizontalAlignment.Center;
                        break;
                    case "右对齐":
                        textBox.TextAlign = HorizontalAlignment.Right;
                        break;
                }
                
                // 设置段落间距（通过Margin模拟）
                var spaceBeforeValue = ParseSpaceValue(spaceBefore);
                var spaceAfterValue = ParseSpaceValue(spaceAfter);
                
                // 设置上下边距来模拟段落间距
                textBox.Margin = new Padding(textBox.Margin.Left, (int)(spaceBeforeValue * 2), 
                    textBox.Margin.Right, (int)(spaceAfterValue * 2));
                
                // 设置缩进（通过Padding模拟）
                var indentValue = ParseIndentValue(indentDistance);
                if (indentType == "首行缩进")
                {
                    textBox.Padding = new Padding((int)(indentValue * 10), textBox.Padding.Top, 
                        textBox.Padding.Right, textBox.Padding.Bottom);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"创建样式预览失败：{ex.Message}");
            }
        }
        
        /// <summary>
        /// 解析间距值
        /// </summary>
        private static float ParseSpaceValue(string spaceText)
        {
            if (string.IsNullOrEmpty(spaceText)) return 0f;
            var cleanText = spaceText.Replace("行", "").Trim();
            return float.TryParse(cleanText, out float value) ? value : 0f;
        }
        
        /// <summary>
        /// 解析缩进值
        /// </summary>
        private static float ParseIndentValue(string indentText)
        {
            if (string.IsNullOrEmpty(indentText)) return 0f;
            var cleanText = indentText.Replace("字符", "").Replace("厘米", "").Trim();
            return float.TryParse(cleanText, out float value) ? value : 0f;
        }

        /// <summary>
        /// 检测单位类型
        /// </summary>
        public static string DetectUnitFromNumber(double number, string[] validUnits)
        {
            if (number > 0 && number < 10)
                return "字符";
            if (number >= 10 && number < 100)
                return "磅";
            return "厘米";
        }

        #endregion

        #region 验证常量和验证方法

        /// <summary>
        /// 验证常量
        /// </summary>
        public static class ValidationConstants
        {
            /// <summary>
            /// 有效的编号样式
            /// </summary>
            public static readonly string[] ValidNumberStyles = new string[]
            {
                "1,2,3...",
                "01,02,03...",
                "A,B,C...",
                "a,b,c...",
                "I,II,III...",
                "i,ii,iii...",
                "一,二,三...",
                "壹,贰,叁...",
                "甲,乙,丙...",
                "正规编号"
            };

            /// <summary>
            /// 有效的编号之后类型
            /// </summary>
            public static readonly string[] ValidAfterNumberTypes = new string[]
            {
                "无",
                "空格",
                "制表位"
            };

            /// <summary>
            /// 有效的链接样式
            /// </summary>
            public static readonly string[] ValidLinkedStyles = new string[]
            {
                "无",
                "标题 1",
                "标题 2",
                "标题 3",
                "标题 4",
                "标题 5",
                "标题 6",
                "标题 7",
                "标题 8",
                "标题 9"
            };
        }

        /// <summary>
        /// 验证辅助类
        /// </summary>
        public static class ValidationHelper
        {
            /// <summary>
            /// 验证编号样式
            /// </summary>
            public static bool IsValidNumberStyle(string numberStyle)
            {
                return MultiLevelDataManager.ValidationConstants.ValidNumberStyles.Contains(numberStyle);
            }

            /// <summary>
            /// 验证编号之后类型
            /// </summary>
            public static bool IsValidAfterNumberType(string afterNumberType)
            {
                return MultiLevelDataManager.ValidationConstants.ValidAfterNumberTypes.Contains(afterNumberType);
            }

            /// <summary>
            /// 验证链接样式
            /// </summary>
            public static bool IsValidLinkedStyle(string linkedStyle)
            {
                return MultiLevelDataManager.ValidationConstants.ValidLinkedStyles.Contains(linkedStyle);
            }

            /// <summary>
            /// 验证缩进值
            /// </summary>
            public static bool IsValidIndentValue(decimal value)
            {
                return value >= 0 && value <= 50; // 0-50厘米
            }

            /// <summary>
            /// 验证制表位位置
            /// </summary>
            public static bool IsValidTabPosition(decimal value)
            {
                return value >= 0 && value <= 50; // 0-50厘米
            }
        }

        #endregion
    }
}
