using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordMan_VSTO.MultiLevel;
using Color = System.Drawing.Color;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式预设管理器 - 管理各种风格预设样式
    /// </summary>
    public static class StylePresetManager
    {
        /// <summary>
        /// 获取所有可用的预设样式名称
        /// </summary>
        public static string[] GetPresetStyleNames()
        {
            return new[] { "公文风格", "论文风格", "报告风格" };
        }

        /// <summary>
        /// 应用预设样式集合到样式列表
        /// </summary>
        public static List<CustomStyle> GetPresetStyles(string presetName)
        {
            switch (presetName)
            {
                case "公文风格":
                    return GetOfficialDocumentStyles();
                case "论文风格":
                    return GetThesisStyles();
                case "报告风格":
                    return GetReportStyles();
                default:
                    throw new ArgumentException($"未知的预设样式：{presetName}");
            }
        }

        /// <summary>
        /// 应用预设样式到控件 - 从样式集合中获取正文样式
        /// </summary>
        public static void ApplyPresetStyle(string presetName, StyleSettingsControls controls)
        {
            try
            {
                // 获取预设样式集合
                var styles = GetPresetStyles(presetName);
                
                // 查找正文样式
                var bodyStyle = styles.FirstOrDefault(s => s.Name == "正文");
                if (bodyStyle == null)
                {
                    throw new InvalidOperationException($"预设样式 '{presetName}' 中未找到正文样式");
                }
                
                // 应用正文样式到控件
                ApplyStyleToControls(bodyStyle, controls);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"应用预设样式失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取公文风格样式集合 - 参考国家标准
        /// </summary>
        private static List<CustomStyle> GetOfficialDocumentStyles()
        {
            return new List<CustomStyle>
            {
                // 正文 - 仿宋三号，1.5倍行距，首行缩进2字符
                new CustomStyle(name: "正文", fontName: "仿宋", engFontName: "仿宋", fontSize: 16f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 3, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 2f, 
                    firstLineIndentByChar: 2, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题1 - 仿宋二号，加粗，1.5倍行距
                new CustomStyle(name: "标题 1", fontName: "仿宋", engFontName: "仿宋", fontSize: 22f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题2 - 仿宋三号，加粗，1.5倍行距
                new CustomStyle(name: "标题 2", fontName: "仿宋", engFontName: "仿宋", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题3 - 仿宋三号，加粗，1.5倍行距
                new CustomStyle(name: "标题 3", fontName: "仿宋", engFontName: "仿宋", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题4 - 仿宋三号，加粗，1.5倍行距
                new CustomStyle(name: "标题 4", fontName: "仿宋", engFontName: "仿宋", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题5 - 仿宋三号，加粗，1.5倍行距
                new CustomStyle(name: "标题 5", fontName: "仿宋", engFontName: "仿宋", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题6 - 仿宋三号，加粗，1.5倍行距
                new CustomStyle(name: "标题 6", fontName: "仿宋", engFontName: "仿宋", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 题注 - 仿宋小四号，居中
                new CustomStyle(name: "题注", fontName: "仿宋", engFontName: "仿宋", fontSize: 12f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.5f, beforeBreak: false, beforeSpacing: 6f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 表内文字 - 仿宋小四号
                new CustomStyle(name: "表内文字", fontName: "仿宋", engFontName: "仿宋", fontSize: 12f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, 
                    numberStyle: 0, numberFormat: null, userDefined: false)
            };
        }


        /// <summary>
        /// 获取论文风格样式集合 - 参考GB/T 7713.2-2022标准
        /// </summary>
        private static List<CustomStyle> GetThesisStyles()
        {
            return new List<CustomStyle>
            {
                // 正文 - 宋体小四号，1.25倍行距，首行缩进2字符
                new CustomStyle(name: "正文", fontName: "宋体", engFontName: "宋体", fontSize: 12f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 3, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 2f, 
                    firstLineIndentByChar: 2, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题1 - 宋体三号，加粗，1.25倍行距
                new CustomStyle(name: "标题 1", fontName: "宋体", engFontName: "宋体", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题2 - 宋体小三号，加粗，1.25倍行距
                new CustomStyle(name: "标题 2", fontName: "宋体", engFontName: "宋体", fontSize: 15f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题3 - 宋体四号，加粗，1.25倍行距
                new CustomStyle(name: "标题 3", fontName: "宋体", engFontName: "宋体", fontSize: 14f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题4 - 宋体小四号，加粗，1.25倍行距
                new CustomStyle(name: "标题 4", fontName: "宋体", engFontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题5 - 宋体小四号，加粗，1.25倍行距
                new CustomStyle(name: "标题 5", fontName: "宋体", engFontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题6 - 宋体小四号，加粗，1.25倍行距
                new CustomStyle(name: "标题 6", fontName: "宋体", engFontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 题注 - 宋体五号，居中
                new CustomStyle(name: "题注", fontName: "宋体", engFontName: "宋体", fontSize: 10.5f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.25f, beforeBreak: false, beforeSpacing: 6f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 表内文字 - 宋体五号
                new CustomStyle(name: "表内文字", fontName: "宋体", engFontName: "宋体", fontSize: 10.5f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, 
                    numberStyle: 0, numberFormat: null, userDefined: false)
            };
        }


        /// <summary>
        /// 获取报告风格样式集合 - 参考GJB 5100A-2017标准
        /// </summary>
        private static List<CustomStyle> GetReportStyles()
        {
            return new List<CustomStyle>
            {
                // 正文 - 微软雅黑四号，1.2倍行距，首行缩进2字符
                new CustomStyle(name: "正文", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 14f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 3, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 2f, 
                    firstLineIndentByChar: 2, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 6f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题1 - 微软雅黑二号，加粗，1.2倍行距
                new CustomStyle(name: "标题 1", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 22f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题2 - 微软雅黑三号，加粗，1.2倍行距
                new CustomStyle(name: "标题 2", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 16f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题3 - 微软雅黑小三号，加粗，1.2倍行距
                new CustomStyle(name: "标题 3", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 15f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题4 - 微软雅黑四号，加粗，1.2倍行距
                new CustomStyle(name: "标题 4", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 14f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题5 - 微软雅黑小四号，加粗，1.2倍行距
                new CustomStyle(name: "标题 5", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 12f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 标题6 - 微软雅黑小四号，加粗，1.2倍行距
                new CustomStyle(name: "标题 6", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 12f, bold: true, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 题注 - 微软雅黑小四号，居中
                new CustomStyle(name: "题注", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 12f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.2f, beforeBreak: false, beforeSpacing: 6f, afterSpacing: 6f, 
                    numberStyle: 0, numberFormat: null, userDefined: false),
                
                // 表内文字 - 微软雅黑小四号
                new CustomStyle(name: "表内文字", fontName: "微软雅黑", engFontName: "微软雅黑", fontSize: 12f, bold: false, italic: false, underline: false, 
                    fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, 
                    firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, 
                    numberStyle: 0, numberFormat: null, userDefined: false)
            };
        }



        /// <summary>
        /// 将CustomStyle应用到控件
        /// </summary>
        private static void ApplyStyleToControls(CustomStyle style, StyleSettingsControls controls)
        {
            // 设置字体 - 分别设置中文字体和英文字体
            SetComboBoxSelection(controls.Cmb_ChnFontName, style.FontName, MultiLevelDataManager.GetSystemFonts());
            SetComboBoxSelection(controls.Cmb_EngFontName, style.EngFontName, MultiLevelDataManager.GetSystemFonts());
            
            // 设置字号 - 将磅值转换为中文字号
            string fontSizeText = MultiLevelDataManager.ConvertFontSizeToString(style.FontSize);
            SetComboBoxSelection(controls.Cmb_FontSize, fontSizeText, MultiLevelDataManager.GetFontSizes());
            
            // 设置字体样式
            controls.Btn_Bold.Pressed = style.Bold;
            controls.Btn_Italic.Pressed = style.Italic;
            controls.Btn_UnderLine.Pressed = style.Underline;
            
            // 设置段落对齐
            string alignment = GetAlignmentText(style.ParaAlignment);
            SetComboBoxSelection(controls.Cmb_ParaAligment, alignment, WordStyleInfo.HAlignments);
            
            // 设置行距
            string lineSpacing = GetLineSpacingText(style.LineSpacing);
            SetComboBoxSelection(controls.Cmb_LineSpacing, lineSpacing, WordStyleInfo.LineSpacings);
            
            // 设置缩进
            controls.Nud_LeftIndent.Value = (decimal)style.LeftIndent;
            controls.Nud_FirstLineIndent.Value = (decimal)style.FirstLineIndent;
            
            // 设置段间距 - 直接设置文本值
            controls.Cmb_BefreSpacing.Text = $"{style.BeforeSpacing:F1} 磅";
            controls.Cmb_AfterSpacing.Text = $"{style.AfterSpacing:F1} 磅";
        }


        /// <summary>
        /// 获取对齐方式文本
        /// </summary>
        private static string GetAlignmentText(int alignment)
        {
            switch (alignment)
            {
                case 0: return "左对齐";
                case 1: return "中对齐";
                case 2: return "右对齐";
                case 3: return "两端对齐";
                case 4: return "分散对齐";
                default: return "左对齐";
            }
        }

        /// <summary>
        /// 获取行距文本
        /// </summary>
        private static string GetLineSpacingText(float lineSpacing)
        {
            if (Math.Abs(lineSpacing - 1.0f) < 0.01f)
                return "单倍行距";
            else if (Math.Abs(lineSpacing - 1.2f) < 0.01f)
                return "1.2倍行距";
            else if (Math.Abs(lineSpacing - 1.25f) < 0.01f)
                return "1.25倍行距";
            else if (Math.Abs(lineSpacing - 1.5f) < 0.01f)
                return "1.5倍行距";
            else if (Math.Abs(lineSpacing - 2.0f) < 0.01f)
                return "双倍行距";
            else
                return $"{lineSpacing:0.0}倍行距";
        }

        /// <summary>
        /// 通用下拉框选择设置方法
        /// </summary>
        private static void SetComboBoxSelection(StandardComboBox comboBox, string value, IEnumerable<string> items)
        {
            var itemList = items.ToList();
            int selectedIndex = itemList.IndexOf(value);
            if (selectedIndex >= 0)
            {
                comboBox.SelectedIndex = selectedIndex;
            }
            else
            {
                comboBox.SelectedIndex = -1;
            }
        }
    }

    /// <summary>
    /// 样式设置控件集合 - 用于传递控件引用
    /// </summary>
    public class StyleSettingsControls
    {
        public StandardComboBox Cmb_ChnFontName { get; set; }
        public StandardComboBox Cmb_EngFontName { get; set; }
        public StandardComboBox Cmb_FontSize { get; set; }
        public ToggleButton Btn_Bold { get; set; }
        public ToggleButton Btn_Italic { get; set; }
        public ToggleButton Btn_UnderLine { get; set; }
        public StandardComboBox Cmb_ParaAligment { get; set; }
        public StandardComboBox Cmb_LineSpacing { get; set; }
        public StandardNumericUpDown Nud_LeftIndent { get; set; }
        public StandardNumericUpDown Nud_FirstLineIndent { get; set; }
        public StandardComboBox Cmb_BefreSpacing { get; set; }
        public StandardComboBox Cmb_AfterSpacing { get; set; }
    }
}
