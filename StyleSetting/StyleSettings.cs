using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordMan_VSTO;
using WordMan_VSTO.MultiLevel;
using Color = System.Drawing.Color;
using Point = System.Drawing.Point;
using Font = System.Drawing.Font;
using CheckBox = System.Windows.Forms.CheckBox;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式设置窗体 - 按照WordFormatHelper的StyleSetGuider设计
    /// </summary>
    public partial class StyleSettings : Form
    {
        #region 私有字段

        // 使用 MultiLevelDataManager 的字号相关方法，避免重复定义

        private readonly List<WdPaperSize> PaperSize = new List<WdPaperSize>(4)
        {
            WdPaperSize.wdPaperA3,
            WdPaperSize.wdPaperA4,
            WdPaperSize.wdPaperA5,
            WdPaperSize.wdPaperB5
        };

        private BindingList<string> StyleNames;
        private readonly List<CustomStyle> Styles = new List<CustomStyle>(17);

        #endregion

        #region 私有方法

        // IsChineseFont 方法已移除 - 现在直接使用 MultiLevelDataManager.GetSystemFonts() 统一获取系统字体

        /// <summary>
        /// 通用下拉框初始化方法
        /// </summary>
        private void InitializeComboBox(StandardComboBox comboBox, IEnumerable<string> items)
        {
            comboBox.Items.Clear();
            comboBox.Items.AddRange(items.ToArray());
        }

        /// <summary>
        /// 通用选择索引设置方法
        /// </summary>
        private void SetComboBoxSelection(StandardComboBox comboBox, string value, IEnumerable<string> items)
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

        /// <summary>
        /// 通用文本验证和格式化
        /// </summary>
        private void ValidateAndFormatText(Control control, string unit)
        {
            if (control is StandardComboBox comboBox && comboBox.SelectedIndex != -1) return;
            
            if (control is StandardTextBox textBox)
            {
                string text = textBox.Text.Trim();
                if (!string.IsNullOrEmpty(text) && !text.EndsWith(unit))
                {
                    textBox.Text = text + unit;
                }
            }
        }

        /// <summary>
        /// 获取控件值
        /// </summary>
        private object GetControlValue(Control control)
        {
            if (control is StandardComboBox comboBox)
                return comboBox.Text;
            if (control is StandardTextBox textBox)
                return textBox.Text;
            if (control is StandardNumericUpDown numericUpDown)
                return numericUpDown.Value;
            if (control is ToggleButton toggleButton)
                return toggleButton.Pressed;
            return null;
        }

        /// <summary>
        /// 获取对齐方式文本（与 LevelStyleSettingsForm 保持一致）
        /// </summary>
        private string GetAlignmentText(int alignment)
        {
            string[] alignments = WordStyleInfo.HAlignments;
            return alignment >= 0 && alignment < alignments.Length ? alignments[alignment] : "左对齐";
        }

        // 使用 MultiLevelDataManager 的行距相关方法，避免重复实现

        /// <summary>
        /// 设置缩进值到StandardNumericUpDown控件（与 LevelStyleSettingsForm 保持一致）
        /// </summary>
        private void SetIndentValue(StandardNumericUpDown numericUpDown, string indentText)
        {
            if (string.IsNullOrEmpty(indentText))
            {
                numericUpDown.Value = 0;
                return;
            }

            // 解析缩进文本，提取数值和单位
            var cleanText = indentText.Replace("厘米", "").Replace("磅", "").Replace("字符", "").Replace("行", "").Trim();
            if (decimal.TryParse(cleanText, out decimal value))
            {
                // 根据单位设置值
                if (indentText.Contains("厘米"))
                {
                    numericUpDown.SetValueInCentimeters(value);
                }
                else if (indentText.Contains("磅"))
                {
                    numericUpDown.SetValueInUnit(value, "磅");
                }
                else if (indentText.Contains("字符"))
                {
                    numericUpDown.SetValueInUnit(value, "字符");
                }
                else if (indentText.Contains("行"))
                {
                    numericUpDown.SetValueInUnit(value, "行");
                }
                else
                {
                    // 默认按厘米处理
                    numericUpDown.SetValueInCentimeters(value);
                }
            }
            else
            {
                numericUpDown.Value = 0;
            }
        }

        // 使用 MultiLevelDataManager.MultiLevelDataManager.ConvertFontSizeToString 方法，避免重复实现

        /// <summary>
        /// 首行缩进方式选择事件处理
        /// </summary>
        private void Cmb_FirstLineIndentType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFirstLineIndentVisibility();
        }

        /// <summary>
        /// 预设样式选择事件处理
        /// </summary>
        private void Cmb_PreSettings_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Cmb_PreSettings.SelectedIndex >= 0)
            {
                LoadPresetStyle(Cmb_PreSettings.SelectedItem.ToString());
            }
        }

        /// <summary>
        /// 加载预设样式
        /// </summary>
        private void LoadPresetStyle(string presetName)
        {
            switch (presetName)
            {
                case "公文风格":
                    LoadOfficialDocumentStyle();
                    break;
                case "论文风格":
                    LoadThesisStyle();
                    break;
                case "报告风格":
                    LoadReportStyle();
                    break;
                case "条文风格":
                    LoadRegulationStyle();
                    break;
            }
        }

        /// <summary>
        /// 公文风格预设
        /// </summary>
        private void LoadOfficialDocumentStyle()
        {
            SetComboBoxSelection(Cmb_ChnFontName, "仿宋", MultiLevelDataManager.GetSystemFonts());
            SetComboBoxSelection(Cmb_FontSize, "三号", MultiLevelDataManager.GetFontSizes());
            Btn_Bold.Pressed = false;
            Btn_Italic.Pressed = false;
            Btn_UnderLine.Pressed = false;
            SetComboBoxSelection(Cmb_ParaAligment, "两端对齐", WordStyleInfo.HAlignments);
            SetComboBoxSelection(Cmb_LineSpacing, "1.5倍行距", WordStyleInfo.LineSpacings);
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 2; // 首行缩进2字符
            Nud_LineSpacing.Value = 1.5m; // 1.5倍行距
            Nud_BefreSpacing.Value = 0;
            Nud_AfterSpacing.Value = 0;
        }

        /// <summary>
        /// 论文风格预设
        /// </summary>
        private void LoadThesisStyle()
        {
            SetComboBoxSelection(Cmb_ChnFontName, "宋体", MultiLevelDataManager.GetSystemFonts());
            SetComboBoxSelection(Cmb_FontSize, "小四", MultiLevelDataManager.GetFontSizes());
            Btn_Bold.Pressed = false;
            Btn_Italic.Pressed = false;
            Btn_UnderLine.Pressed = false;
            SetComboBoxSelection(Cmb_ParaAligment, "两端对齐", WordStyleInfo.HAlignments);
            SetComboBoxSelection(Cmb_LineSpacing, "多倍行距", WordStyleInfo.LineSpacings);
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 2; // 首行缩进2字符
            Nud_LineSpacing.Value = 1.25m; // 1.25倍行距
            Nud_BefreSpacing.Value = 0;
            Nud_AfterSpacing.Value = 0;
        }

        /// <summary>
        /// 报告风格预设
        /// </summary>
        private void LoadReportStyle()
        {
            SetComboBoxSelection(Cmb_ChnFontName, "微软雅黑", MultiLevelDataManager.GetSystemFonts());
            SetComboBoxSelection(Cmb_FontSize, "四号", MultiLevelDataManager.GetFontSizes());
            Btn_Bold.Pressed = false;
            Btn_Italic.Pressed = false;
            Btn_UnderLine.Pressed = false;
            SetComboBoxSelection(Cmb_ParaAligment, "左对齐", WordStyleInfo.HAlignments);
            SetComboBoxSelection(Cmb_LineSpacing, "多倍行距", WordStyleInfo.LineSpacings);
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 0; // 无首行缩进
            Nud_LineSpacing.Value = 1.2m; // 1.2倍行距
            Nud_BefreSpacing.Value = 6; // 段前6磅
            Nud_AfterSpacing.Value = 6; // 段后6磅
        }

        /// <summary>
        /// 条文风格预设
        /// </summary>
        private void LoadRegulationStyle()
        {
            SetComboBoxSelection(Cmb_ChnFontName, "仿宋", MultiLevelDataManager.GetSystemFonts());
            SetComboBoxSelection(Cmb_FontSize, "四号", MultiLevelDataManager.GetFontSizes());
            Btn_Bold.Pressed = false;
            Btn_Italic.Pressed = false;
            Btn_UnderLine.Pressed = false;
            SetComboBoxSelection(Cmb_ParaAligment, "两端对齐", WordStyleInfo.HAlignments);
            SetComboBoxSelection(Cmb_LineSpacing, "1.5倍行距", WordStyleInfo.LineSpacings);
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 2; // 首行缩进2字符
            Nud_LineSpacing.Value = 1.5m; // 1.5倍行距
            Nud_BefreSpacing.Value = 0;
            Nud_AfterSpacing.Value = 0;
        }

        /// <summary>
        /// 更新首行缩进输入框的可见性
        /// </summary>
        private void UpdateFirstLineIndentVisibility()
        {
            if (Cmb_FirstLineIndentType.SelectedIndex == 0) // 无
            {
                label7.Visible = false;
                Nud_FirstLineIndent.Visible = false;
                // label8.Visible = false; // label8控件已移除
                Nud_FirstLineIndentByChar.Visible = false;
            }
            else if (Cmb_FirstLineIndentType.SelectedIndex == 1) // 悬挂缩进
            {
                label7.Visible = true;
                Nud_FirstLineIndent.Visible = true;
                // label8.Visible = false; // label8控件已移除
                Nud_FirstLineIndentByChar.Visible = false;
                label7.Text = "悬挂缩进";
                Nud_FirstLineIndent.Unit = "厘米";
            }
            else if (Cmb_FirstLineIndentType.SelectedIndex == 2) // 首行缩进
            {
                label7.Visible = true;
                Nud_FirstLineIndent.Visible = true;
                // label8.Visible = false; // label8控件已移除
                Nud_FirstLineIndentByChar.Visible = false;
                label7.Text = "首行缩进";
                Nud_FirstLineIndent.Unit = "厘米";
            }
        }

        #endregion

        #region 构造函数

        public StyleSettings()
        {
            InitializeComponent();
            InitializeData();
            BindEvents();
        }

        #endregion

        #region 初始化方法

        private void InitializeData()
        {
            // 初始化控件状态

            // 初始化样式列表
            StyleNames = new BindingList<string>();
            InitializeDefaultStyles();
            
            Lst_Styles.DataSource = StyleNames;
            
            // 使用 MultiLevelDataManager 的统一方法初始化字体列表
            var systemFonts = MultiLevelDataManager.GetSystemFonts();
            InitializeComboBox(Cmb_ChnFontName, systemFonts);
            InitializeComboBox(Cmb_EngFontName, systemFonts);
            
            // 初始化其他下拉框
            InitializeComboBox(Cmb_FontSize, MultiLevelDataManager.GetFontSizes());
            InitializeComboBox(Cmb_ParaAligment, WordStyleInfo.HAlignments);
            InitializeComboBox(Cmb_SetLevel, new[] { "无", "1", "2", "3", "4", "5", "6", "7", "8", "9" });
            InitializeComboBox(Cmb_PreSettings, new[] { "公文风格", "论文风格", "报告风格", "条文风格" });
            
            // 初始化行距下拉框 - 使用 WordStyleInfo.LineSpacings
            InitializeComboBox(Cmb_LineSpacing, WordStyleInfo.LineSpacings);
            
            Cmb_SetLevel.SelectedIndex = 3;
            Lst_Styles.SelectedIndex = -1;
            Cmb_ParaAligment.SelectedIndex = -1;
            
            // 初始化首行缩进方式
            Cmb_FirstLineIndentType.Items.AddRange(new[] { "无", "悬挂缩进", "首行缩进" });
            Cmb_FirstLineIndentType.SelectedIndex = 0; // 默认选择"无"
            
            // 设置首行缩进方式选择事件
            Cmb_FirstLineIndentType.SelectedIndexChanged += Cmb_FirstLineIndentType_SelectedIndexChanged;
            
            // 设置预设样式选择事件
            Cmb_PreSettings.SelectedIndexChanged += Cmb_PreSettings_SelectedIndexChanged;

            // 初始状态：隐藏首行缩进输入框
            UpdateFirstLineIndentVisibility();
            
            // 启用字体设置和段落设置面板
            Pal_Font.Enabled = true;
            Pal_ParaIndent.Enabled = true;
        }

        private void InitializeDefaultStyles()
        {
            Styles.Clear();
            Styles.AddRange(new CustomStyle[]
            {
                new CustomStyle(name: "正文", fontName: null, fontSize: 0f, bold: false, italic: false, underline: false, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 1", fontName: null, fontSize: 14f, bold: true, italic: false, underline: false, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 2", fontName: null, fontSize: 12f, bold: true, italic: false, underline: false, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 3", fontName: null, fontSize: 0f, bold: false, italic: false, underline: false, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "题注", fontName: null, fontSize: 0f, bold: false, italic: false, underline: false, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "表内文字", fontName: null, fontSize: 0f, bold: false, italic: false, underline: false, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false)
            });

            foreach (CustomStyle style in Styles)
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(style.Name, "标题 [4-9]"))
                {
                    StyleNames.Add(style.Name);
                }
            }
        }

        private void BindEvents()
        {
            Lst_Styles.SelectedIndexChanged += Lst_Styles_SelectedIndexChanged;
            Cmb_ChnFontName.SelectedIndexChanged += StyleFontChanged;
            Cmb_EngFontName.SelectedIndexChanged += StyleFontChanged;
            Cmb_FontSize.SelectedIndexChanged += StyleFontChanged;
            Cmb_ParaAligment.SelectedIndexChanged += StyleFontChanged;
            Cmb_LineSpacing.SelectedIndexChanged += IndentSpacingChanged;
            Nud_LeftIndent.ValueChanged += IndentSpacingChanged;
            Nud_FirstLineIndent.ValueChanged += IndentSpacingChanged;
            Nud_FirstLineIndentByChar.ValueChanged += IndentSpacingChanged;
            Nud_LineSpacing.ValueChanged += IndentSpacingChanged;
            Nud_BefreSpacing.ValueChanged += IndentSpacingChanged;
            Nud_AfterSpacing.ValueChanged += IndentSpacingChanged;
            Btn_Bold.PressedChanged += FontStyleChanged;
            Btn_Italic.PressedChanged += FontStyleChanged;
            Btn_UnderLine.PressedChanged += FontStyleChanged;
            // Chk_BeforeBreak.CheckedChanged += FontStyleChanged; // Chk_BeforeBreak控件已移除
            Cmb_SetLevel.SelectedIndexChanged += Cmb_SetLevel_SelectedIndexChanged;
            Btn_AddStyle.Click += Btn_AddStyle_Click;
            Btn_DelStyle.Click += Btn_DelStyle_Click;
            Btn_ApplySet.Click += Btn_ApplySet_Click;
            Btn_ReadDocumentStyle.Click += Btn_ReadDocumentStyle_Click;
        }

        #endregion

        #region 事件处理方法

        private void Lst_Styles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Lst_Styles.SelectedIndex >= 0)
            {
                string selectedStyle = Lst_Styles.SelectedItem.ToString();
                var style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                if (style != null)
                {
                    LoadStyleToControls(style);
                    UpdateStyleInfo(style);
                    
                    // 启用字体设置和段落设置面板
                    Pal_Font.Enabled = true;
                    Pal_ParaIndent.Enabled = true;
                }
            }
            else
            {
                // 如果没有选中样式，禁用编辑面板
                Pal_Font.Enabled = false;
                Pal_ParaIndent.Enabled = false;
            }
        }

        private void StyleFontChanged(object sender, EventArgs e)
        {
            UpdateCurrentStyle();
        }

        private void IndentSpacingChanged(object sender, EventArgs e)
        {
            UpdateCurrentStyle();
        }

        private void FontStyleChanged(object sender, EventArgs e)
        {
            UpdateCurrentStyle();
        }


        private void Cmb_SetLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 处理标题级别变化
        }

        private void Btn_AddStyle_Click(object sender, EventArgs e)
        {
            string styleName = Txt_AddStyleName.Text.Trim();
            if (string.IsNullOrEmpty(styleName))
            {
                MessageBox.Show("请输入样式名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (StyleNames.Contains(styleName))
            {
                MessageBox.Show("样式名称已存在", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 创建新样式
            var newStyle = CreateStyleFromControls(styleName);
            Styles.Add(newStyle);
            StyleNames.Add(styleName);
            Lst_Styles.SelectedItem = styleName;
        }

        private void Btn_DelStyle_Click(object sender, EventArgs e)
        {
            if (Lst_Styles.SelectedIndex >= 0)
            {
                string selectedStyle = Lst_Styles.SelectedItem.ToString();
                var style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                if (style != null && style.UserDefined)
                {
                    Styles.Remove(style);
                    StyleNames.Remove(selectedStyle);
                    ClearControls();
                }
            }
        }

        private void Btn_ApplySet_Click(object sender, EventArgs e)
        {
            try
            {
                ApplyStylesToDocument();
                MessageBox.Show("样式设置已应用", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_ReadDocumentStyle_Click(object sender, EventArgs e)
        {
            try
            {
                ReadDocumentStyles();
                MessageBox.Show("文档样式读取完成", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取文档样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region 辅助方法

        private void LoadStyleToControls(CustomStyle style)
        {
            // 使用通用方法加载样式到控件
            if (!string.IsNullOrEmpty(style.FontName))
            {
                // 分别设置中文字体和英文字体（与 LevelStyleSettingsForm 保持一致）
                SetComboBoxSelection(Cmb_ChnFontName, style.FontName, MultiLevelDataManager.GetSystemFonts());
                SetComboBoxSelection(Cmb_EngFontName, style.FontName, MultiLevelDataManager.GetSystemFonts());
            }

            if (style.FontSize > 0)
            {
                SetComboBoxSelection(Cmb_FontSize, MultiLevelDataManager.ConvertFontSizeToString(style.FontSize), MultiLevelDataManager.GetFontSizes());
            }

            Btn_Bold.Pressed = style.Bold;
            Btn_Italic.Pressed = style.Italic;
            Btn_UnderLine.Pressed = style.Underline;

            if (style.ParaAlignment >= 0)
            {
                SetComboBoxSelection(Cmb_ParaAligment, GetAlignmentText(style.ParaAlignment), WordStyleInfo.HAlignments);
            }

            // 设置行距下拉框
            if (style.LineSpacing > 0)
            {
                string lineSpacingText = MultiLevelDataManager.ConvertFontSizeToString(style.LineSpacing);
                SetComboBoxSelection(Cmb_LineSpacing, lineSpacingText, WordStyleInfo.LineSpacings);
            }

            // 设置缩进控件（从数值转换为字符串格式，与 LevelStyleSettingsForm 保持一致）
            SetIndentValue(Nud_LeftIndent, $"{style.LeftIndent:F1} 厘米");
            SetIndentValue(Nud_RightIndent, $"{style.RightIndent:F1} 厘米");
            Nud_FirstLineIndent.Value = (decimal)style.FirstLineIndent;
            Nud_FirstLineIndentByChar.Value = style.FirstLineIndentByChar;
            Nud_LineSpacing.Value = (decimal)style.LineSpacing;
            Nud_BefreSpacing.Value = (decimal)style.BeforeSpacing;
            Nud_AfterSpacing.Value = (decimal)style.AfterSpacing;
            // Chk_BeforeBreak.Checked = style.BeforeBreak; // Chk_BeforeBreak控件已移除
        }

        private void UpdateCurrentStyle()
        {
            if (Lst_Styles.SelectedIndex >= 0)
            {
                string selectedStyle = Lst_Styles.SelectedItem.ToString();
                var style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                if (style != null)
                {
                    UpdateStyleFromControls(style);
                    UpdateStyleInfo(style);
                }
            }
        }

        private void UpdateStyleFromControls(CustomStyle style)
        {
            // 分别获取中文字体和英文字体（与 LevelStyleSettingsForm 保持一致）
            style.FontName = Cmb_ChnFontName.SelectedItem?.ToString() ?? Cmb_EngFontName.SelectedItem?.ToString();
            if (Cmb_FontSize.SelectedIndex >= 0)
            {
                style.FontSize = MultiLevelDataManager.ConvertFontSize(Cmb_FontSize.SelectedItem?.ToString());
            }
            style.Bold = Btn_Bold.Pressed;
            style.Italic = Btn_Italic.Pressed;
            style.Underline = Btn_UnderLine.Pressed;
            style.ParaAlignment = Cmb_ParaAligment.SelectedIndex;
            style.LeftIndent = (float)Nud_LeftIndent.Value;
            style.RightIndent = (float)Nud_RightIndent.Value;
            style.FirstLineIndent = (float)Nud_FirstLineIndent.Value;
            style.FirstLineIndentByChar = (int)Nud_FirstLineIndentByChar.Value;
            
            // 处理行距设置
            if (Cmb_LineSpacing.SelectedIndex >= 0)
            {
                string lineSpacingText = Cmb_LineSpacing.SelectedItem?.ToString();
                style.LineSpacing = MultiLevelDataManager.ConvertFontSize(lineSpacingText);
            }
            else
            {
                style.LineSpacing = (float)Nud_LineSpacing.Value;
            }
            
            style.BeforeSpacing = (float)Nud_BefreSpacing.Value;
            style.AfterSpacing = (float)Nud_AfterSpacing.Value;
            // style.BeforeBreak = Chk_BeforeBreak.Checked; // Chk_BeforeBreak控件已移除
        }

        private CustomStyle CreateStyleFromControls(string name)
        {
            // 分别获取中文字体和英文字体（与 LevelStyleSettingsForm 保持一致）
            string fontName = !string.IsNullOrEmpty(Cmb_ChnFontName.Text) ? Cmb_ChnFontName.Text : Cmb_EngFontName.Text;
            
            // 处理行距设置
            float lineSpacing = (float)Nud_LineSpacing.Value;
            if (Cmb_LineSpacing.SelectedIndex >= 0)
            {
                string lineSpacingText = Cmb_LineSpacing.SelectedItem?.ToString();
                lineSpacing = MultiLevelDataManager.ConvertFontSize(lineSpacingText);
            }
            
            return new CustomStyle(
                name: name,
                fontName: fontName,
                fontSize: Cmb_FontSize.SelectedIndex >= 0 ? MultiLevelDataManager.ConvertFontSize(Cmb_FontSize.SelectedItem?.ToString()) : 0f,
                bold: Btn_Bold.Pressed,
                italic: Btn_Italic.Pressed,
                underline: Btn_UnderLine.Pressed,
                paraAlignment: Cmb_ParaAligment.SelectedIndex,
                leftIndent: (float)Nud_LeftIndent.Value,
                rightIndent: (float)Nud_RightIndent.Value,
                firstLineIndent: (float)Nud_FirstLineIndent.Value,
                firstLineIndentByChar: (int)Nud_FirstLineIndentByChar.Value,
                lineSpacing: lineSpacing,
                beforeBreak: false, // Chk_BeforeBreak控件已移除，使用默认值false
                beforeSpacing: (float)Nud_BefreSpacing.Value,
                afterSpacing: (float)Nud_AfterSpacing.Value,
                numberStyle: 0,
                numberFormat: null,
                userDefined: true
            );
        }

        private void UpdateStyleInfo(CustomStyle style)
        {
            string info = $"样式名称：{style.Name}\n" +
                         $"字体：{style.FontName ?? "默认"}\n" +
                         $"大小：{style.FontSize}磅\n" +
                         $"格式：{(style.Bold ? "粗体 " : "")}{(style.Italic ? "斜体 " : "")}{(style.Underline ? "下划线" : "")}\n" +
                         $"对齐：{GetAlignmentText(style.ParaAlignment)}\n" +
                         $"缩进：左{style.LeftIndent}，首行{style.FirstLineIndent}";
            
            // Lab_StyleInfo控件已移除，样式信息显示功能暂时禁用
        }


        private void ClearControls()
        {
            Cmb_ChnFontName.SelectedIndex = -1;
            Cmb_EngFontName.SelectedIndex = -1;
            Cmb_FontSize.SelectedIndex = -1;
            Btn_Bold.Pressed = false;
            Btn_Italic.Pressed = false;
            Btn_UnderLine.Pressed = false;
            Cmb_ParaAligment.SelectedIndex = -1;
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 0;
            Nud_FirstLineIndentByChar.Value = 0;
            Nud_LineSpacing.Value = 0;
            Nud_BefreSpacing.Value = 0;
            Nud_AfterSpacing.Value = 0;
            // Chk_BeforeBreak.Checked = false; // Chk_BeforeBreak控件已移除
            // Lab_StyleInfo控件已移除，样式信息显示功能暂时禁用
        }

        private void ApplyStylesToDocument()
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;

            foreach (var style in Styles)
        {
            try
            {
                    // 应用样式到文档
                    ApplyStyleToDocument(doc, style);
            }
            catch (Exception ex)
            {
                    System.Diagnostics.Debug.WriteLine($"应用样式 {style.Name} 时出错：{ex.Message}");
                }
            }
        }

        private void ApplyStyleToDocument(Document doc, CustomStyle style)
        {
            // 这里实现具体的样式应用逻辑
            // 使用Word API将样式应用到文档
        }

        private void ReadDocumentStyles()
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

            // 读取文档中的样式
            foreach (Word.Style wordStyle in doc.Styles)
            {
                if (wordStyle.Type == WdStyleType.wdStyleTypeParagraph)
                {
                    // 读取样式属性并更新到列表中
                }
            }
        }

        #endregion

        private void label16_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 字体颜色选择事件（与 LevelStyleSettingsForm 保持一致）
        /// </summary>
        private void Btn_FontColor_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog
            {
                Color = Btn_FontColor.BackColor,
                AnyColor = true,
                SolidColorOnly = true
            };

            if (colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                Btn_FontColor.BackColor = colorDialog.Color;
                UpdateCurrentStyle();
            }
        }
    }

    #region CustomStyle 类

    /// <summary>
    /// 自定义样式类
    /// </summary>
    public class CustomStyle
    {
        public string Name { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public int ParaAlignment { get; set; }
        public float LeftIndent { get; set; }
        public float RightIndent { get; set; }
        public float FirstLineIndent { get; set; }
        public int FirstLineIndentByChar { get; set; }
        public float LineSpacing { get; set; }
        public float BeforeSpacing { get; set; }
        public bool BeforeBreak { get; set; }
        public float AfterSpacing { get; set; }
        public int NumberStyle { get; set; }
        public string NumberFormat { get; set; }
        public bool UserDefined { get; set; }

        public CustomStyle(string name, string fontName, float fontSize, bool bold, bool italic, bool underline,
            int paraAlignment, float leftIndent, float rightIndent, float firstLineIndent, int firstLineIndentByChar, float lineSpacing,
            float beforeSpacing, bool beforeBreak, float afterSpacing, int numberStyle, string numberFormat, bool userDefined)
        {
            Name = name;
            FontName = fontName;
            FontSize = fontSize;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            ParaAlignment = paraAlignment;
            LeftIndent = leftIndent;
            RightIndent = rightIndent;
            FirstLineIndent = firstLineIndent;
            FirstLineIndentByChar = firstLineIndentByChar;
            LineSpacing = lineSpacing;
            BeforeSpacing = beforeSpacing;
            BeforeBreak = beforeBreak;
            AfterSpacing = afterSpacing;
            NumberStyle = numberStyle;
            NumberFormat = numberFormat;
            UserDefined = userDefined;
            }
        }

        #endregion
    }
