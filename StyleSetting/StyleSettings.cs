using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
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
            {
                // 根据控件名称确定单位
                if (control.Name == "Nud_LeftIndent" || control.Name == "Nud_RightIndent" || control.Name == "Nud_FirstLineIndent")
                {
                    return $"{numericUpDown.GetValueInCentimeters():0.0} 厘米";
                }
                else
                {
                    return $"{numericUpDown.GetValueInCentimeters():0.0} 厘米";
                }
            }
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
                if (indentText.Contains("字符"))
                {
                    numericUpDown.SetValueInUnit(value, "字符");
                }
                else if (indentText.Contains("厘米"))
                {
                    numericUpDown.SetValueInCentimeters(value);
                }
                else if (indentText.Contains("磅"))
                {
                    numericUpDown.SetValueInUnit(value, "磅");
                }
                else if (indentText.Contains("行"))
                {
                    numericUpDown.SetValueInUnit(value, "行");
                }
                else
                {
                    // 默认按字符处理（与多级段落设置保持一致）
                    numericUpDown.SetValueInUnit(value, "字符");
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
        private void 首行缩进方式下拉框_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFirstLineIndentVisibility();
        }

        /// <summary>
        /// 预设样式选择事件处理
        /// </summary>
        private void 风格下拉框_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (风格下拉框.SelectedIndex >= 0)
            {
                LoadPresetStyle(风格下拉框.SelectedItem.ToString());
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
            SetComboBoxSelection(Cmb_BefreSpacing, "0.0 磅", WordStyleInfo.SpaceBeforeValues);
            SetComboBoxSelection(Cmb_AfterSpacing, "0.0 磅", WordStyleInfo.SpaceAfterValues);
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
            SetComboBoxSelection(Cmb_LineSpacing, "1.25 倍行距", WordStyleInfo.LineSpacings);
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 2; // 首行缩进2字符
            SetComboBoxSelection(Cmb_BefreSpacing, "0.0 磅", WordStyleInfo.SpaceBeforeValues);
            SetComboBoxSelection(Cmb_AfterSpacing, "0.0 磅", WordStyleInfo.SpaceAfterValues);
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
            SetComboBoxSelection(Cmb_LineSpacing, "1.2 倍行距", WordStyleInfo.LineSpacings);
            Nud_LeftIndent.Value = 0;
            Nud_FirstLineIndent.Value = 0; // 无首行缩进
            SetComboBoxSelection(Cmb_BefreSpacing, "6.0 磅", WordStyleInfo.SpaceBeforeValues); // 段前6磅
            SetComboBoxSelection(Cmb_AfterSpacing, "6.0 磅", WordStyleInfo.SpaceAfterValues); // 段后6磅
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
            SetComboBoxSelection(Cmb_BefreSpacing, "0.0 磅", WordStyleInfo.SpaceBeforeValues);
            SetComboBoxSelection(Cmb_AfterSpacing, "0.0 磅", WordStyleInfo.SpaceAfterValues);
        }

        /// <summary>
        /// 更新首行缩进输入框的可见性（与多级段落设置保持一致）
        /// </summary>
        private void UpdateFirstLineIndentVisibility()
        {
            if (首行缩进方式下拉框.SelectedIndex == 0) // 无
            {
                label7.Visible = false;
                Nud_FirstLineIndent.Visible = false;
            }
            else if (首行缩进方式下拉框.SelectedIndex == 1) // 悬挂缩进
            {
                label7.Visible = true;
                Nud_FirstLineIndent.Visible = true;
                label7.Text = "悬挂缩进";
                Nud_FirstLineIndent.Unit = "字符"; // 与多级段落设置保持一致，使用字符单位
            }
            else if (首行缩进方式下拉框.SelectedIndex == 2) // 首行缩进
            {
                label7.Visible = true;
                Nud_FirstLineIndent.Visible = true;
                label7.Text = "首行缩进";
                Nud_FirstLineIndent.Unit = "字符"; // 与多级段落设置保持一致，使用字符单位
            }
        }

        #endregion

        #region 构造函数

        public StyleSettings()
        {
            InitializeComponent();
            InitializeData();
            BindEvents();
            
            // 自动读取文档样式
            try
            {
                ReadDocumentStyles();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"自动读取文档样式时出错：{ex.Message}");
                // 如果读取文档样式失败，确保至少有一些默认样式
                if (Styles.Count == 0)
                {
                    InitializeDefaultStyles();
                    // 重新绑定样式名称列表
                    StyleNames.Clear();
                    foreach (var style in Styles)
                    {
                        StyleNames.Add(style.Name);
                    }
                    Lst_Styles.DataSource = null;
                    Lst_Styles.DataSource = StyleNames;
                }
            }
        }

        #endregion

        #region 初始化方法

        private void InitializeData()
        {
            // 初始化控件状态

            // 初始化样式列表
            StyleNames = new BindingList<string>();
            Lst_Styles.DataSource = StyleNames;
            
            // 使用 MultiLevelDataManager 的统一方法初始化字体列表
            var systemFonts = MultiLevelDataManager.GetSystemFonts();
            InitializeComboBox(Cmb_ChnFontName, systemFonts);
            InitializeComboBox(Cmb_EngFontName, systemFonts);
            
            // 初始化其他下拉框
            InitializeComboBox(Cmb_FontSize, MultiLevelDataManager.GetFontSizes());
            InitializeComboBox(Cmb_ParaAligment, WordStyleInfo.HAlignments);
            InitializeComboBox(风格下拉框, new[] { "公文风格", "论文风格", "报告风格", "条文风格" });
            
            
            // 初始化行距下拉框 - 使用 WordStyleInfo.LineSpacings
            InitializeComboBox(Cmb_LineSpacing, WordStyleInfo.LineSpacings);
            
            // 初始化段落间距下拉框 - 使用 WordStyleInfo 的预设值
            InitializeComboBox(Cmb_BefreSpacing, WordStyleInfo.SpaceBeforeValues);
            InitializeComboBox(Cmb_AfterSpacing, WordStyleInfo.SpaceAfterValues);
            
            // 设置显示标题数下拉框
            InitializeComboBox(显示标题数下拉框, new[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" });
            显示标题数下拉框.SelectedIndex = 4; // 默认显示4级标题
            显示标题数下拉框.SelectedIndexChanged += 显示标题数下拉框_SelectedIndexChanged;
            
            Lst_Styles.SelectedIndex = -1;
            Cmb_ParaAligment.SelectedIndex = -1;
            
            // 初始化首行缩进方式
            首行缩进方式下拉框.Items.Clear(); // 先清空避免重复
            首行缩进方式下拉框.Items.AddRange(new[] { "无", "悬挂缩进", "首行缩进" });
            首行缩进方式下拉框.SelectedIndex = 0; // 默认选择"无"
            
            // 设置首行缩进方式选择事件
            首行缩进方式下拉框.SelectedIndexChanged += 首行缩进方式下拉框_SelectedIndexChanged;
            
            // 设置预设样式选择事件
            风格下拉框.SelectedIndexChanged += 风格下拉框_SelectedIndexChanged;

            // 初始化样式名称输入框
            InitializeStyleNameTextBox();

            // 初始状态：隐藏首行缩进输入框
            UpdateFirstLineIndentVisibility();
            
            // 根据选择的标题数过滤样式
            FilterStylesByTitleCount();
            
            // 启用字体设置和段落设置面板
            Pal_Font.Enabled = true;
            Pal_ParaIndent.Enabled = true;
            
            // 启用按钮
            加载.Enabled = true;
            添加.Enabled = true;
            删除.Enabled = true;
        }

        /// <summary>
        /// 初始化样式名称输入框
        /// </summary>
        private void InitializeStyleNameTextBox()
        {
            Txt_AddStyleName.Text = "请输入需要增加的样式名称";
            Txt_AddStyleName.ForeColor = Color.Gray;
            
            // 添加焦点事件处理
            Txt_AddStyleName.Enter += Txt_AddStyleName_Enter;
            Txt_AddStyleName.Leave += Txt_AddStyleName_Leave;
        }

        /// <summary>
        /// 样式名称输入框获得焦点事件
        /// </summary>
        private void Txt_AddStyleName_Enter(object sender, EventArgs e)
        {
            if (Txt_AddStyleName.Text == "请输入需要增加的样式名称")
            {
                Txt_AddStyleName.Text = "";
                Txt_AddStyleName.ForeColor = Color.Black;
            }
        }

        /// <summary>
        /// 样式名称输入框失去焦点事件
        /// </summary>
        private void Txt_AddStyleName_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Txt_AddStyleName.Text))
            {
                Txt_AddStyleName.Text = "请输入需要增加的样式名称";
                Txt_AddStyleName.ForeColor = Color.Gray;
            }
        }

        private void InitializeDefaultStyles()
        {
            Styles.Clear();
            Styles.AddRange(new CustomStyle[]
            {
                new CustomStyle(name: "正文", fontName: "宋体", fontSize: 10.5f, bold: false, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 2f, firstLineIndentByChar: 2, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 1", fontName: "宋体", fontSize: 16f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 2", fontName: "宋体", fontSize: 14f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 3", fontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 4", fontName: "宋体", fontSize: 12f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 5", fontName: "宋体", fontSize: 10.5f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "标题 6", fontName: "宋体", fontSize: 10.5f, bold: true, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 12f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "题注", fontName: "宋体", fontSize: 9f, bold: false, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 1, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 6f, afterSpacing: 6f, numberStyle: 0, numberFormat: null, userDefined: false),
                new CustomStyle(name: "表内文字", fontName: "宋体", fontSize: 9f, bold: false, italic: false, underline: false, fontColor: Color.Black, paraAlignment: 0, leftIndent: 0f, rightIndent: 0f, firstLineIndent: 0f, firstLineIndentByChar: 0, lineSpacing: 1.0f, beforeBreak: false, beforeSpacing: 0f, afterSpacing: 0f, numberStyle: 0, numberFormat: null, userDefined: false)
            });

            foreach (CustomStyle style in Styles)
            {
                StyleNames.Add(style.Name);
            }
        }

        private void BindEvents()
        {
            Lst_Styles.SelectedIndexChanged += Lst_Styles_SelectedIndexChanged;
            // 移除实时更新事件，样式设置只在点击应用设置按钮时生效
            // Cmb_ChnFontName.SelectedIndexChanged += StyleFontChanged;
            // Cmb_EngFontName.SelectedIndexChanged += StyleFontChanged;
            // Cmb_FontSize.SelectedIndexChanged += StyleFontChanged;
            // Cmb_ParaAligment.SelectedIndexChanged += StyleFontChanged;
            // Cmb_LineSpacing.SelectedIndexChanged += IndentSpacingChanged;
            Cmb_LineSpacing.TextChanged += Cmb_LineSpacing_TextChanged;
            Cmb_LineSpacing.Validated += Cmb_LineSpacing_Validated;
            // Nud_LeftIndent.ValueChanged += IndentSpacingChanged;
            // Nud_FirstLineIndent.ValueChanged += IndentSpacingChanged;
            // Nud_LineSpacing.ValueChanged += IndentSpacingChanged;
            // Nud_BefreSpacing.ValueChanged += IndentSpacingChanged;
            // Nud_AfterSpacing.ValueChanged += IndentSpacingChanged;
            // Btn_Bold.PressedChanged += FontStyleChanged;
            // Btn_Italic.PressedChanged += FontStyleChanged;
            // Btn_UnderLine.PressedChanged += FontStyleChanged;
            Btn_FontColor.Click += Btn_FontColor_Click;
            // Chk_BeforeBreak.CheckedChanged += FontStyleChanged; // Chk_BeforeBreak控件已移除
            添加.Click += 添加_Click;
            删除.Click += 删除_Click;
            Btn_ApplySet.Click += Btn_ApplySet_Click;
            //Btn_ReadDocumentStyle.Click += Btn_ReadDocumentStyle_Click;
            this.关闭.Click += 关闭_Click;
            this.加载.Click += 加载_Click;
            
            // 添加导入导出按钮事件
            导入.Click += Btn_Import_Click;
            导出.Click += Btn_Export_Click;
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



        /// <summary>
        /// 显示标题数下拉框选择变化事件
        /// </summary>
        private void 显示标题数下拉框_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 根据选择的标题数过滤样式列表
            FilterStylesByTitleCount();
        }

        /// <summary>
        /// 根据选择的标题数过滤样式列表
        /// </summary>
        private void FilterStylesByTitleCount()
        {
            int titleCount = 显示标题数下拉框.SelectedIndex; // 0=0级, 1=1级, 2=2级, ...
            
            // 清空当前样式列表
            StyleNames.Clear();
            
            foreach (var style in Styles)
            {
                bool shouldShow = false;
                
                if (titleCount == 0) // 不显示标题样式，只显示其他样式
                {
                    if (!style.Name.StartsWith("标题"))
                    {
                        shouldShow = true;
                    }
                }
                else if (style.Name.StartsWith("标题"))
                {
                    // 提取标题级别，只显示指定数量及以下的标题
                    if (style.Name == "标题 1" && titleCount >= 1) shouldShow = true;
                    else if (style.Name == "标题 2" && titleCount >= 2) shouldShow = true;
                    else if (style.Name == "标题 3" && titleCount >= 3) shouldShow = true;
                    else if (style.Name == "标题 4" && titleCount >= 4) shouldShow = true;
                    else if (style.Name == "标题 5" && titleCount >= 5) shouldShow = true;
                    else if (style.Name == "标题 6" && titleCount >= 6) shouldShow = true;
                    else if (style.Name == "标题 7" && titleCount >= 7) shouldShow = true;
                    else if (style.Name == "标题 8" && titleCount >= 8) shouldShow = true;
                    else if (style.Name == "标题 9" && titleCount >= 9) shouldShow = true;
                }
                else if (style.Name == "正文" || style.Name == "题注" || style.Name == "表内文字")
                {
                    // 始终显示正文、题注和表内文字
                    shouldShow = true;
                }
                
                if (shouldShow)
                {
                    StyleNames.Add(style.Name);
                }
            }
            
            // 刷新列表显示
            Lst_Styles.DataSource = null;
            Lst_Styles.DataSource = StyleNames;
        }


        private void 添加_Click(object sender, EventArgs e)
        {
            string styleName = Txt_AddStyleName.Text.Trim();
            
            // 检查是否为空或提示文本
            if (string.IsNullOrEmpty(styleName) || styleName == "请输入需要增加的样式名称")
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
            
            // 清空输入框并恢复提示文本
            Txt_AddStyleName.Text = "请输入需要增加的样式名称";
            Txt_AddStyleName.ForeColor = Color.Gray;
        }

        private void 删除_Click(object sender, EventArgs e)
        {
            if (Lst_Styles.SelectedIndex >= 0)
            {
                string selectedStyle = Lst_Styles.SelectedItem.ToString();
                var style = Styles.FirstOrDefault(s => s.Name == selectedStyle);
                
                if (style != null)
                {
                    // 检查是否为内置样式
                    if (!style.UserDefined)
                    {
                        MessageBox.Show("不能删除内置样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    // 确认删除
                    var result = MessageBox.Show($"确定要删除样式 '{selectedStyle}' 吗？", "确认删除", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    
                    if (result == DialogResult.Yes)
                    {
                        Styles.Remove(style);
                        StyleNames.Remove(selectedStyle);
                        ClearControls();
                        
                        // 刷新样式列表显示
                        Lst_Styles.DataSource = null;
                        Lst_Styles.DataSource = StyleNames;
                        
                        MessageBox.Show("样式删除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选择要删除的样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Btn_ApplySet_Click(object sender, EventArgs e)
        {
            try
            {
                // 先更新当前选中的样式（如果有选中的话）
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
                
                // 然后应用样式到文档
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

        private void 关闭_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void 加载_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取文档中的所有样式
                var documentStyles = GetDocumentStyles();
                
                if (documentStyles.Count == 0)
                {
                    MessageBox.Show("文档中没有找到样式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // 创建样式选择窗体
                using (var styleSelectionForm = new StyleSelectionForm())
                {
                    // 初始化可用样式
                    styleSelectionForm.InitializeStyles(documentStyles);
                    
                    // 设置当前已选择的样式
                    var currentSelectedStyles = StyleNames.ToList();
                    styleSelectionForm.SetSelectedStyles(currentSelectedStyles);
                    
                    // 显示窗体
                    if (styleSelectionForm.ShowDialog() == DialogResult.OK)
                    {
                        // 更新样式列表
                        StyleNames.Clear();
                        foreach (var selectedStyle in styleSelectionForm.SelectedStyles)
                        {
                            StyleNames.Add(selectedStyle);
                        }
                        
                        // 刷新样式列表显示
                        Lst_Styles.DataSource = null;
                        Lst_Styles.DataSource = StyleNames;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出样式设置
        /// </summary>
        private void Btn_Export_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = StyleFileManager.ShowSaveFileDialog("样式设置", "XML文件|*.xml|所有文件|*.*");
                if (!string.IsNullOrEmpty(filePath))
                {
                    // 根据当前显示的样式列表进行导出，而不是导出所有样式
                    var stylesToExport = new List<CustomStyle>();
                    foreach (var styleName in StyleNames)
                    {
                        var style = Styles.FirstOrDefault(s => s.Name == styleName);
                        if (style != null)
                        {
                            stylesToExport.Add(style);
                        }
                    }
                    
                    StyleFileManager.SerializeListToXml(stylesToExport, filePath);
                    MessageBox.Show($"样式设置导出成功，共导出 {stylesToExport.Count} 个样式", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出样式设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导入样式设置
        /// </summary>
        private void Btn_Import_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = StyleFileManager.ShowOpenFileDialog("XML文件|*.xml|所有文件|*.*");
                if (!string.IsNullOrEmpty(filePath))
                {
                    var importedStyles = StyleFileManager.DeserializeListFromXml<CustomStyle>(filePath);
                    if (importedStyles != null && importedStyles.Count > 0)
                    {
                        // 清空现有样式
                        Styles.Clear();
                        StyleNames.Clear();
                        
                        // 加载导入的样式
                        Styles.AddRange(importedStyles);
                        foreach (var style in importedStyles)
                        {
                            StyleNames.Add(style.Name);
                        }
                        
                        // 刷新显示
                        Lst_Styles.DataSource = null;
                        Lst_Styles.DataSource = StyleNames;
                        
                        // 自动应用导入的样式到文档
                        try
                        {
                            ApplyStylesToDocument();
                            MessageBox.Show($"成功导入并应用 {importedStyles.Count} 个样式到文档", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception applyEx)
                        {
                            MessageBox.Show($"成功导入 {importedStyles.Count} 个样式，但应用样式时出错：{applyEx.Message}", "部分成功", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入样式设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Cmb_LineSpacing_TextChanged(object sender, EventArgs e)
        {
            // 移除实时更新，样式设置只在点击应用设置按钮时生效
            // UpdateCurrentStyle();
        }

        private void Cmb_LineSpacing_Validated(object sender, EventArgs e)
        {
            ValidateAndFormatText(Cmb_LineSpacing, Cmb_LineSpacing.Text.EndsWith("行") ? "行" : "磅");
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
            Btn_FontColor.BackColor = style.FontColor;

            if (style.ParaAlignment >= 0)
            {
                SetComboBoxSelection(Cmb_ParaAligment, GetAlignmentText(style.ParaAlignment), WordStyleInfo.HAlignments);
            }

            // 设置行距下拉框
            if (style.LineSpacing > 0)
            {
                string lineSpacingText = ConvertLineSpacingToString(style.LineSpacing);
                SetComboBoxSelection(Cmb_LineSpacing, lineSpacingText, WordStyleInfo.LineSpacings);
            }

            // 设置缩进控件（使用厘米单位显示）
            SetIndentValue(Nud_LeftIndent, $"{style.LeftIndent:F1} 厘米");
            SetIndentValue(Nud_RightIndent, $"{style.RightIndent:F1} 厘米");
            
            // 根据首行缩进值设置首行缩进方式和控件
            if (style.FirstLineIndent == 0f)
            {
                首行缩进方式下拉框.SelectedIndex = 0; // 无
                UpdateFirstLineIndentVisibility();
            }
            else
            {
                首行缩进方式下拉框.SelectedIndex = 2; // 首行缩进
                UpdateFirstLineIndentVisibility();
                SetIndentValue(Nud_FirstLineIndent, $"{style.FirstLineIndent:F1} 厘米");
            }
            // 段落间距使用磅为单位显示
            SetComboBoxSelection(Cmb_BefreSpacing, $"{style.BeforeSpacing:F1} 磅", WordStyleInfo.SpaceBeforeValues);
            SetComboBoxSelection(Cmb_AfterSpacing, $"{style.AfterSpacing:F1} 磅", WordStyleInfo.SpaceAfterValues);
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
            style.FontColor = Btn_FontColor.BackColor;
            style.ParaAlignment = Cmb_ParaAligment.SelectedIndex;
            // 从控件读取缩进值，直接使用厘米单位
            style.LeftIndent = (float)Nud_LeftIndent.GetValueInCentimeters();
            style.RightIndent = (float)Nud_RightIndent.GetValueInCentimeters();
            
            // 根据首行缩进方式设置首行缩进值
            if (首行缩进方式下拉框.SelectedIndex == 0) // 无
            {
                style.FirstLineIndent = 0f;
            }
            else
            {
                style.FirstLineIndent = (float)Nud_FirstLineIndent.GetValueInCentimeters();
            }
            
            // 处理行距设置
            if (!string.IsNullOrEmpty(Cmb_LineSpacing.Text))
            {
                style.LineSpacing = ConvertLineSpacingToFloat(Cmb_LineSpacing.Text);
            }
            else
            {
                style.LineSpacing = 1.0f; // 默认单倍行距
            }
            
            // 从控件读取段落间距值，需要转换为厘米单位
            style.BeforeSpacing = ConvertSpacingToCentimeters(Cmb_BefreSpacing.Text);
            style.AfterSpacing = ConvertSpacingToCentimeters(Cmb_AfterSpacing.Text);
            // style.BeforeBreak = Chk_BeforeBreak.Checked; // Chk_BeforeBreak控件已移除
        }

        private CustomStyle CreateStyleFromControls(string name)
        {
            // 分别获取中文字体和英文字体（与 LevelStyleSettingsForm 保持一致）
            string fontName = !string.IsNullOrEmpty(Cmb_ChnFontName.Text) ? Cmb_ChnFontName.Text : Cmb_EngFontName.Text;
            
            // 处理行距设置
            float lineSpacing = 1.0f; // 默认单倍行距
            if (!string.IsNullOrEmpty(Cmb_LineSpacing.Text))
            {
                lineSpacing = ConvertLineSpacingToFloat(Cmb_LineSpacing.Text);
            }
            
            return new CustomStyle(
                name: name,
                fontName: fontName,
                fontSize: Cmb_FontSize.SelectedIndex >= 0 ? MultiLevelDataManager.ConvertFontSize(Cmb_FontSize.SelectedItem?.ToString()) : 0f,
                bold: Btn_Bold.Pressed,
                italic: Btn_Italic.Pressed,
                underline: Btn_UnderLine.Pressed,
                fontColor: Btn_FontColor.BackColor,
                paraAlignment: Cmb_ParaAligment.SelectedIndex,
                leftIndent: (float)Nud_LeftIndent.GetValueInCentimeters(),
                rightIndent: (float)Nud_RightIndent.GetValueInCentimeters(),
                firstLineIndent: 首行缩进方式下拉框.SelectedIndex == 0 ? 0f : (float)Nud_FirstLineIndent.GetValueInCentimeters(),
                firstLineIndentByChar: 0, // 使用默认值0，因为Nud_FirstLineIndentByChar控件已移除
                lineSpacing: lineSpacing,
                beforeBreak: false, // Chk_BeforeBreak控件已移除，使用默认值false
                beforeSpacing: ConvertSpacingToCentimeters(Cmb_BefreSpacing.Text),
                afterSpacing: ConvertSpacingToCentimeters(Cmb_AfterSpacing.Text),
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
            SetComboBoxSelection(Cmb_BefreSpacing, "0.0 磅", WordStyleInfo.SpaceBeforeValues);
            SetComboBoxSelection(Cmb_AfterSpacing, "0.0 磅", WordStyleInfo.SpaceAfterValues);
            // Chk_BeforeBreak.Checked = false; // Chk_BeforeBreak控件已移除
            // Lab_StyleInfo控件已移除，样式信息显示功能暂时禁用
        }

        /// <summary>
        /// 将行距文本转换为数值（与多级段落设置保持一致）
        /// </summary>
        private float ConvertLineSpacingToFloat(string lineSpacingText)
        {
            if (string.IsNullOrEmpty(lineSpacingText))
                return 1.0f;

            try
            {
                if (lineSpacingText == "单倍行距")
                {
                    return 1.0f;
                }
                else if (lineSpacingText == "1.5倍行距")
                {
                    return 1.5f;
                }
                else if (lineSpacingText == "双倍行距")
                {
                    return 2.0f;
                }
                else if (lineSpacingText.EndsWith("倍行距"))
                {
                    // 处理多倍行距，提取倍数值
                    string valueText = lineSpacingText.Replace("倍行距", "").Trim();
                    if (float.TryParse(valueText, out float multipleValue))
                    {
                        return multipleValue;
                    }
                }
                else if (lineSpacingText.EndsWith("磅"))
                {
                    // 处理固定行距，转换为倍行距
                    string valueText = lineSpacingText.Replace("磅", "").Trim();
                    if (float.TryParse(valueText, out float exactValue))
                    {
                        // 使用默认字体大小12磅进行计算
                        return exactValue / 12f;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换行距时出错：{ex.Message}");
            }

            return 1.0f; // 默认单倍行距
        }

        /// <summary>
        /// 将行距数值转换为文本（与多级段落设置保持一致）
        /// </summary>
        private string ConvertLineSpacingToString(float lineSpacing)
        {
            try
            {
                if (lineSpacing == 1.0f)
                {
                    return "单倍行距";
                }
                else if (lineSpacing == 1.5f)
                {
                    return "1.5倍行距";
                }
                else if (lineSpacing == 2.0f)
                {
                    return "双倍行距";
                }
                else
                {
                    return $"{lineSpacing:F1} 倍行距";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换行距文本时出错：{ex.Message}");
                return "单倍行距";
            }
        }

        /// <summary>
        /// 将段落间距文本转换为厘米单位
        /// </summary>
        private float ConvertSpacingToCentimeters(string spacingText)
        {
            if (string.IsNullOrEmpty(spacingText))
                return 0f;

            try
            {
                if (spacingText.EndsWith("磅"))
                {
                    string valueText = spacingText.TrimEnd(' ', '磅');
                    if (float.TryParse(valueText, out float points))
                    {
                        var app = Globals.ThisAddIn.Application;
                        return (float)app.PointsToCentimeters(points);
                    }
                }
                else if (spacingText.EndsWith("行"))
                {
                    string valueText = spacingText.TrimEnd(' ', '行');
                    if (float.TryParse(valueText, out float lines))
                    {
                        var app = Globals.ThisAddIn.Application;
                        return (float)app.PointsToCentimeters(MultiLevelDataManager.LinesToPoints(lines));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"转换段落间距时出错：{ex.Message}");
            }

            return 0f;
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
            try
            {
                // 检查样式是否存在，如果不存在则创建
                Style wordStyle = null;
                try
                {
                    wordStyle = doc.Styles[style.Name];
                }
                catch
                {
                    // 样式不存在，创建新样式
                    wordStyle = doc.Styles.Add(style.Name, WdStyleType.wdStyleTypeParagraph);
                }

                if (wordStyle != null)
                {
                    // 设置字体
                    if (!string.IsNullOrEmpty(style.FontName))
                    {
                        wordStyle.Font.NameFarEast = style.FontName;
                        wordStyle.Font.NameAscii = style.FontName;
                    }
                    
                    if (style.FontSize > 0)
                    {
                        wordStyle.Font.Size = style.FontSize;
                    }
                    
                    wordStyle.Font.Bold = style.Bold ? -1 : 0;
                    wordStyle.Font.Italic = style.Italic ? -1 : 0;
                    wordStyle.Font.Underline = style.Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                    
                    // 设置字体颜色
                    try
                    {
                        int r = style.FontColor.R;
                        int g = style.FontColor.G;
                        int b = style.FontColor.B;
                        int wordRgb = (b << 16) | (g << 8) | r; // Word 使用 BGR 格式
                        wordStyle.Font.Color = (WdColor)wordRgb;
                    }
                    catch
                    {
                        wordStyle.Font.Color = WdColor.wdColorAutomatic;
                    }

                    // 设置段落格式
                    var paragraphFormat = wordStyle.ParagraphFormat;
                    
                    // 设置对齐方式
                    switch (style.ParaAlignment)
                    {
                        case 0: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                        case 1: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; break;
                        case 2: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; break;
                        case 3: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify; break;
                        case 4: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute; break;
                        default: paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                    }
                    
                    // 设置缩进（转换为磅）
                    var app = Globals.ThisAddIn.Application;
                    paragraphFormat.LeftIndent = app.CentimetersToPoints(style.LeftIndent);
                    paragraphFormat.RightIndent = app.CentimetersToPoints(style.RightIndent);
                    paragraphFormat.FirstLineIndent = app.CentimetersToPoints(style.FirstLineIndent);
                    
                    // 设置行距
                    if (style.LineSpacing > 0)
                    {
                        if (style.LineSpacing == 1.0f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                        }
                        else if (style.LineSpacing == 1.5f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                        }
                        else if (style.LineSpacing == 2.0f)
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
                        }
                        else
                        {
                            paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                            paragraphFormat.LineSpacing = app.LinesToPoints(style.LineSpacing);
                        }
                    }
                    
                    // 设置段前段后间距
                    paragraphFormat.SpaceBefore = app.CentimetersToPoints(style.BeforeSpacing);
                    paragraphFormat.SpaceAfter = app.CentimetersToPoints(style.AfterSpacing);
                    
                    // 设置分页
                    paragraphFormat.PageBreakBefore = style.BeforeBreak ? -1 : 0;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"应用样式 {style.Name} 失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档中的所有样式名称
        /// </summary>
        private List<string> GetDocumentStyles()
        {
            var styles = new List<string>();
            
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app?.ActiveDocument?.Styles != null)
                {
                    // 保存当前选择状态
                    var originalSelection = app.Selection;
                    var originalRange = originalSelection.Range;
                    
                    try
                    {
                        // 遍历样式列表，但不影响文档选择
                        var stylesCollection = app.ActiveDocument.Styles;
                        for (int i = 1; i <= stylesCollection.Count; i++)
                        {
                            try
                            {
                                var style = stylesCollection[i];
                                // 只添加用户定义的样式和内置样式
                                if (style.BuiltIn || style.InUse)
                                {
                                    styles.Add(style.NameLocal);
                                }
                            }
                            catch
                            {
                                // 忽略无法访问的样式
                                continue;
                            }
                        }
                    }
                    finally
                    {
                        // 恢复原始选择状态
                        try
                        {
                            originalRange.Select();
                        }
                        catch
                        {
                            // 如果无法恢复选择，忽略错误
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取文档样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            return styles;
        }

        private void ReadDocumentStyles()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

                // 清空现有样式列表
                Styles.Clear();
                StyleNames.Clear();

                // 定义要加载的样式名称（参考WordFormatHelper的实现）
                var targetStyleNames = new[] { "正文", "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "题注", "表内文字" };

            // 读取指定的样式
            foreach (var styleName in targetStyleNames)
            {
                try
                {
                    var wordStyle = doc.Styles[styleName];
                    if (wordStyle != null && wordStyle.Type == WdStyleType.wdStyleTypeParagraph)
                    {
                        // 创建自定义样式对象
                        var customStyle = new CustomStyle(
                            name: wordStyle.NameLocal,
                            fontName: wordStyle.Font.NameFarEast ?? wordStyle.Font.Name,
                            fontSize: wordStyle.Font.Size,
                            bold: wordStyle.Font.Bold == (int)WdConstants.wdToggle,
                            italic: wordStyle.Font.Italic == (int)WdConstants.wdToggle,
                            underline: wordStyle.Font.Underline != WdUnderline.wdUnderlineNone,
                            fontColor: GetWordFontColor(wordStyle.Font),
                            paraAlignment: (int)wordStyle.ParagraphFormat.Alignment,
                            leftIndent: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.LeftIndent),
                            rightIndent: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.RightIndent),
                            firstLineIndent: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.FirstLineIndent),
                            firstLineIndentByChar: 0, // Word API不直接提供字符单位
                            lineSpacing: ConvertWordLineSpacingToFloat(wordStyle.ParagraphFormat),
                            beforeSpacing: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.SpaceBefore),
                            beforeBreak: wordStyle.ParagraphFormat.PageBreakBefore != 0,
                            afterSpacing: (float)app.PointsToCentimeters(wordStyle.ParagraphFormat.SpaceAfter),
                            numberStyle: 0,
                            numberFormat: null,
                            userDefined: !wordStyle.BuiltIn
                        );

                        Styles.Add(customStyle);
                        StyleNames.Add(customStyle.Name);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"读取样式 {styleName} 时出错：{ex.Message}");
                }
            }

                // 刷新样式列表显示
                Lst_Styles.DataSource = null;
                Lst_Styles.DataSource = StyleNames;
                
                System.Diagnostics.Debug.WriteLine($"成功加载了 {Styles.Count} 个文档样式");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"读取文档样式时出错：{ex.Message}");
                // 不显示错误消息，避免影响用户体验
            }
        }

        /// <summary>
        /// 将Word行距格式转换为统一的行距值
        /// </summary>
        private float ConvertWordLineSpacingToFloat(ParagraphFormat paragraphFormat)
        {
            try
            {
                var lineSpacingRule = paragraphFormat.LineSpacingRule;
                var lineSpacing = paragraphFormat.LineSpacing;
                
                switch (lineSpacingRule)
                {
                    case WdLineSpacing.wdLineSpaceSingle:
                        return 1.0f;
                    case WdLineSpacing.wdLineSpace1pt5:
                        return 1.5f;
                    case WdLineSpacing.wdLineSpaceDouble:
                        return 2.0f;
                    case WdLineSpacing.wdLineSpaceMultiple:
                        // 多倍行距，需要转换
                        var app = Globals.ThisAddIn.Application;
                        return (float)app.PointsToLines(lineSpacing);
                    case WdLineSpacing.wdLineSpaceExactly:
                        // 固定行距，转换为倍行距
                        // 使用默认字体大小12磅进行计算
                        return lineSpacing / 12f;
                    case WdLineSpacing.wdLineSpaceAtLeast:
                        // 最小行距，转换为倍行距
                        // 使用默认字体大小12磅进行计算
                        return lineSpacing / 12f;
                    default:
                        return 1.0f; // 默认单倍行距
                }
            }
            catch
            {
                return 1.0f; // 出错时返回默认单倍行距
            }
        }

        /// <summary>
        /// 获取Word字体颜色（与WordStyleInfo保持一致）
        /// </summary>
        private Color GetWordFontColor(Microsoft.Office.Interop.Word.Font font)
        {
            try
            {
                // 检查是否是自动颜色
                if (font.Color == WdColor.wdColorAutomatic)
                {
                    return Color.Black;
                }
                
                // 使用ColorTranslator.FromOle方法，这是处理Word颜色的标准方法
                return ColorTranslator.FromOle((int)font.Color);
            }
            catch
            {
                return Color.Black;
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
                SolidColorOnly = true,
                AllowFullOpen = true,
                FullOpen = true
            };

            if (colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                Btn_FontColor.BackColor = colorDialog.Color;
                
                // 移除实时更新，样式设置只在点击应用设置按钮时生效
                // 字体颜色设置将在点击应用设置按钮时生效
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
        [XmlIgnore]
        public Color FontColor { get; set; }
        
        /// <summary>
        /// 用于XML序列化的字体颜色属性
        /// </summary>
        [XmlElement("FontColor")]
        public string FontColorString
        {
            get { return FontColor.ToArgb().ToString(); }
            set { FontColor = Color.FromArgb(int.Parse(value)); }
        }
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

        /// <summary>
        /// 无参数构造函数，用于XML序列化
        /// </summary>
        public CustomStyle()
        {
            // 设置默认值
            Name = "";
            FontName = "宋体";
            FontSize = 10.5f;
            Bold = false;
            Italic = false;
            Underline = false;
            FontColor = Color.Black;
            ParaAlignment = 0; // 左对齐
            LeftIndent = 0f;
            RightIndent = 0f;
            FirstLineIndent = 0f;
            FirstLineIndentByChar = 0;
            LineSpacing = 1f;
            BeforeSpacing = 0f;
            BeforeBreak = false;
            AfterSpacing = 0f;
            NumberStyle = 0;
            NumberFormat = "";
            UserDefined = false;
        }

        /// <summary>
        /// 带参数的构造函数
        /// </summary>
        public CustomStyle(string name, string fontName, float fontSize, bool bold, bool italic, bool underline, Color fontColor,
            int paraAlignment, float leftIndent, float rightIndent, float firstLineIndent, int firstLineIndentByChar, float lineSpacing,
            float beforeSpacing, bool beforeBreak, float afterSpacing, int numberStyle, string numberFormat, bool userDefined)
        {
            Name = name;
            FontName = fontName;
            FontSize = fontSize;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            FontColor = fontColor;
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
