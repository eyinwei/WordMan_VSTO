using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;

namespace WordMan_VSTO.MultiLevel
{
    public partial class LevelStyleSettingsForm : Form
    {
        private readonly List<WordStyleInfo> LevelStyles = new List<WordStyleInfo>();
        private readonly List<string> FontNames = new List<string>();
        private bool userChange;
        private int maxLevel = 9;

        /// <summary>
        /// 获取当前样式设置
        /// </summary>
        public List<WordStyleInfo> GetLevelStyles()
        {
            return LevelStyles;
        }

        /// <summary>
        /// 加载现有的样式设置
        /// </summary>
        public void LoadExistingStyles(List<WordStyleInfo> existingStyles)
        {
            if (existingStyles == null || existingStyles.Count == 0)
                return;

            // 清空现有样式
            LevelStyles.Clear();
            
            // 加载现有样式
            LevelStyles.AddRange(existingStyles);
            
            // 重新绑定数据源
            if (Dta_StyleList != null)
            {
                Dta_StyleList.DataSource = null;
                Dta_StyleList.DataSource = LevelStyles;
            }
        }

        public LevelStyleSettingsForm(int maxLevel = 9)
        {
            this.maxLevel = maxLevel;
            InitializeComponent();
            InitializeForm();
        }

        private void InitializeForm()
        {
            userChange = false;
            
            // 获取系统字体
            LoadSystemFonts();
            
            // 初始化级别样式（只有在没有现有样式时才初始化）
            if (LevelStyles.Count == 0)
            {
                InitializeLevelStyles();
            }
            
            // 初始化控件
            InitializeControls();
            
            // 初始化事件处理
            InitializeEventHandlers();
            
            userChange = true;
        }

        #region 通用方法

        /// <summary>
        /// 加载系统字体
        /// </summary>
        private void LoadSystemFonts()
        {
            var installedFontCollection = new InstalledFontCollection();
            FontNames.AddRange(installedFontCollection.Families.Select(fontFamily => fontFamily.Name));
        }

        /// <summary>
        /// 通用下拉框初始化方法
        /// </summary>
        private void InitializeComboBox(StandardComboBox comboBox, IEnumerable<string> items)
        {
            comboBox.Items.Clear();
            comboBox.Items.AddRange(items.ToArray());
        }

        /// <summary>
        /// 通用DataGridView更新方法
        /// </summary>
        private void UpdateDataGridView(string columnName, int rowIndex)
        {
            if (Dta_StyleList.Columns.Contains(columnName))
            {
                int columnIndex = Dta_StyleList.Columns[columnName].Index;
                Dta_StyleList.UpdateCellValue(columnIndex, rowIndex);
                
                // 对于颜色列，需要强制刷新显示
                if (columnName == "Col_FontColor")
                {
                    Dta_StyleList.InvalidateCell(columnIndex, rowIndex);
                }
            }
        }

        /// <summary>
        /// 通用选择索引设置方法
        /// </summary>
        private void SetComboBoxSelection(StandardComboBox comboBox, string value, IEnumerable<string> items)
        {
            var itemList = items.ToList();
            int selectedIndex = itemList.IndexOf(value);
            if (selectedIndex != -1)
            {
                comboBox.SelectedIndex = selectedIndex;
            }
            else
            {
                comboBox.Text = value;
            }
        }

        /// <summary>
        /// 通用文本验证和格式化
        /// </summary>
        private void ValidateAndFormatText(Control control, string unit)
        {
            if (control is StandardComboBox comboBox && comboBox.SelectedIndex != -1) return;
            
            var text = control.Text.TrimEnd(' ', '磅', '厘', '米', '行');
            if (float.TryParse(text, out float value))
            {
                control.Text = $"{value:0.0} {unit}";
            }
            else
            {
                // 设置默认值
                if (control is StandardComboBox cb) cb.SelectedIndex = 0;
                else control.Text = $"0.0 {unit}";
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
                return $"{numericUpDown.GetValueInCentimeters():0.0} 厘米";
            if (control is ToggleButton toggleButton)
                return toggleButton.Pressed;
            if (control is Button button && button.Name == "Btn_FontColor")
                return button.BackColor;
            return null;
        }

        /// <summary>
        /// 设置缩进值到StandardNumericUpDown控件
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

        #endregion

        /// <summary>
        /// 设置所有控件字体为微软雅黑
        /// </summary>
        private void SetAllControlsFont()
        {
            // 设置窗体字体
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            
            // 设置GroupBox字体
            Grp_SetSelectedStyle.Font = new Font("Microsoft YaHei", 10F, FontStyle.Bold);
            Grp_SetSelectedStyle.ForeColor = Color.FromArgb(51, 51, 51);
            
            // 设置DataGridView字体
            Dta_StyleList.Font = new Font("Microsoft YaHei", 9F);
            Dta_StyleList.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold);
            Dta_StyleList.DefaultCellStyle.Font = new Font("Microsoft YaHei", 9F);
            
            // 设置所有标签字体
            SetControlFont(label1, 9F, FontStyle.Bold);
            SetControlFont(label2, 9F, FontStyle.Bold);
            SetControlFont(label3, 9F, FontStyle.Bold);
            SetControlFont(label4, 9F, FontStyle.Bold);
            SetControlFont(label5, 9F, FontStyle.Bold);
            SetControlFont(label6, 9F, FontStyle.Bold);
            SetControlFont(label7, 9F, FontStyle.Bold);
            SetControlFont(label8, 9F, FontStyle.Bold);
            SetControlFont(label9, 9F, FontStyle.Bold);
            SetControlFont(label10, 9F, FontStyle.Bold);
            SetControlFont(label11, 9F, FontStyle.Bold);
            SetControlFont(label12, 9F, FontStyle.Bold);
            SetControlFont(label13, 9F, FontStyle.Bold);
            SetControlFont(label14, 9F, FontStyle.Bold);
            
            // 设置所有下拉框字体
            SetControlFont(Cmb_ChnFontName, 9F);
            SetControlFont(Cmb_EngFontName, 9F);
            SetControlFont(Cmb_FontSize, 9F);
            SetControlFont(Cmb_Alignment, 9F);
            SetControlFont(Cmb_LineSpacing, 9F);
            SetControlFont(Cmb_SpaceBefore, 9F);
            SetControlFont(Cmb_SpaceAfter, 9F);
            
            // 设置所有数值输入框字体
            SetControlFont(Txt_LeftIndent, 9F);
            SetControlFont(Txt_RightIndent, 9F);
            
            // 设置所有按钮字体
            SetControlFont(Btn_FontColor, 9F, FontStyle.Bold);
            SetControlFont(Btn_Bold, 8F, FontStyle.Bold);
            SetControlFont(Btn_Italic, 8F, FontStyle.Bold);
            SetControlFont(Btn_Underline, 8F, FontStyle.Bold);
            SetControlFont(Btn_BreakBefore, 8F, FontStyle.Bold);
            SetControlFont(Btn_SetStyles, 9F, FontStyle.Bold);
            SetControlFont(Btn_Cancel, 9F, FontStyle.Bold);
        }

        /// <summary>
        /// 设置控件字体
        /// </summary>
        private void SetControlFont(Control control, float size, FontStyle style = FontStyle.Regular)
        {
            if (control != null)
            {
                control.Font = new Font("Microsoft YaHei", size, style);
            }
        }

        private void InitializeLevelStyles()
        {
            try
            {
                var headingStyles = new[]
                {
                    WdBuiltinStyle.wdStyleHeading1, WdBuiltinStyle.wdStyleHeading2, WdBuiltinStyle.wdStyleHeading3,
                    WdBuiltinStyle.wdStyleHeading4, WdBuiltinStyle.wdStyleHeading5, WdBuiltinStyle.wdStyleHeading6,
                    WdBuiltinStyle.wdStyleHeading7, WdBuiltinStyle.wdStyleHeading8, WdBuiltinStyle.wdStyleHeading9
                };

                var styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
                
                for (int i = 0; i < maxLevel && i < headingStyles.Length; i++)
                {
                    object index = headingStyles[i];
                    Style style = styles[ref index];
                    
                    // 创建样式对象，从Word文档样式读取所有属性包括颜色
                    var wordStyleInfo = new WordStyleInfo(style, headingStyles[i]);
                    
                    // 注释掉强制设置黑色，让颜色从文档样式自动读取
                    // wordStyleInfo.FontColor = Color.Black;
                    
                    LevelStyles.Add(wordStyleInfo);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化级别样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeControls()
        {
            // 初始化DataGridView
            InitializeDataGridView();
            
            // 设置所有控件字体为微软雅黑
            SetAllControlsFont();
            
            // 使用通用方法初始化下拉框
            InitializeComboBox(Cmb_ChnFontName, FontNames);
            InitializeComboBox(Cmb_EngFontName, FontNames);
            InitializeComboBox(Cmb_FontSize, WordStyleInfo.FontSizes);
            InitializeComboBox(Cmb_Alignment, WordStyleInfo.HAlignments);
            InitializeComboBox(Cmb_LineSpacing, WordStyleInfo.LineSpacings);
            InitializeComboBox(Cmb_SpaceBefore, WordStyleInfo.SpaceBeforeValues);
            InitializeComboBox(Cmb_SpaceAfter, WordStyleInfo.SpaceAfterValues);
        }

        /// <summary>
        /// 初始化事件处理
        /// </summary>
        private void InitializeEventHandlers()
        {
            // 字体相关控件
            Cmb_ChnFontName.SelectedIndexChanged += (s, e) => OnControlValueChanged(s, "ChnFontName", "Col_ChnFontName");
            Cmb_EngFontName.SelectedIndexChanged += (s, e) => OnControlValueChanged(s, "EngFontName", "Col_EngFontName");
            Cmb_FontSize.TextChanged += (s, e) => OnControlValueChanged(s, "FontSize", "Col_FontSize");
            
            // 格式相关控件
            Btn_Bold.PressedChanged += (s, e) => OnControlValueChanged(s, "Bold", "Col_FontBold");
            Btn_Italic.PressedChanged += (s, e) => OnControlValueChanged(s, "Italic", "Col_FontItalic");
            Btn_Underline.PressedChanged += (s, e) => OnControlValueChanged(s, "Underline", "Col_FontUnderline");
            
            // 缩进相关控件
            Txt_LeftIndent.ValueChanged += (s, e) => OnControlValueChanged(s, "LeftIndent", "Col_LeftIndent");
            Txt_RightIndent.ValueChanged += (s, e) => OnControlValueChanged(s, "RightIndent", "Col_RightIndent");
            
            // 间距相关控件
            Cmb_LineSpacing.TextChanged += (s, e) => OnControlValueChanged(s, "LineSpace", "Col_LineSpace");
            Cmb_SpaceBefore.TextChanged += (s, e) => OnControlValueChanged(s, "SpaceBefore", "Col_SpaceBefore");
            Cmb_SpaceAfter.TextChanged += (s, e) => OnControlValueChanged(s, "SpaceAfter", "Col_SpaceAfter");
            
            // 对齐相关控件
            Cmb_Alignment.SelectedIndexChanged += (s, e) => OnControlValueChanged(s, "HAlignment", "Col_HAlignment");
            
            // 分页相关控件
            Btn_BreakBefore.PressedChanged += (s, e) => OnControlValueChanged(s, "BreakBefore", "Col_BreakBefore");
            
            // 颜色控件
            Btn_FontColor.BackColorChanged += (s, e) => OnControlValueChanged(s, "FontColor", "Col_FontColor");
        }

        /// <summary>
        /// 通用控件值变化事件处理
        /// </summary>
        private void OnControlValueChanged(object sender, string propertyName, string columnName)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0) return;

            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                var style = LevelStyles[selectedRow.Index];
                var control = sender as Control;
                
                // 使用反射设置属性值
                var property = typeof(WordStyleInfo).GetProperty(propertyName);
                if (property != null)
                {
                    var newValue = GetControlValue(control);
                    property.SetValue(style, newValue, null);
                    
                    // 对于颜色属性，确保数据绑定正确更新
                    if (propertyName == "FontColor")
                    {
                        // 强制刷新整个行
                        Dta_StyleList.Refresh();
                    }
                    else
                    {
                        UpdateDataGridView(columnName, selectedRow.Index);
                    }
                }
            }
        }

        private void InitializeDataGridView()
        {
            Dta_StyleList.Columns.Clear();
            Dta_StyleList.ReadOnly = true;
            Dta_StyleList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Dta_StyleList.AutoGenerateColumns = false;
            
            // 添加列
            Dta_StyleList.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_StyleName",
                    DataPropertyName = "StyleName",
                    Frozen = true,
                    HeaderText = "样式名",
                    ReadOnly = true,
                    Width = 80
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_ChnFontName",
                    DataPropertyName = "ChnFontName",
                    HeaderText = "中文字体",
                    Width = 90
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_EngFontName",
                    DataPropertyName = "EngFontName",
                    HeaderText = "西文字体",
                    Width = 90
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_FontSize",
                    DataPropertyName = "FontSize",
                    HeaderText = "字体大小",
                    Width = 90
                },
                new DataGridViewImageColumn
                {
                    Name = "Col_FontColor",
                    DataPropertyName = "FontColor",
                    HeaderText = "颜色",
                    ImageLayout = DataGridViewImageCellLayout.Normal,
                    Width = 50
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontBold",
                    DataPropertyName = "Bold",
                    HeaderText = "粗体",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 50
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontItalic",
                    DataPropertyName = "Italic",
                    HeaderText = "斜体",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 50
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontUnderline",
                    DataPropertyName = "Underline",
                    HeaderText = "下划线",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 60
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_LeftIndent",
                    DataPropertyName = "LeftIndent",
                    HeaderText = "左缩进",
                    Width = 80
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_RightIndent",
                    DataPropertyName = "RightIndent",
                    HeaderText = "右缩进",
                    Width = 80
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_LineSpace",
                    DataPropertyName = "LineSpace",
                    HeaderText = "行距",
                    Width = 80
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_SpaceBefore",
                    DataPropertyName = "SpaceBefore",
                    HeaderText = "段前行距",
                    Width = 100
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_SpaceAfter",
                    DataPropertyName = "SpaceAfter",
                    HeaderText = "段后行距",
                    Width = 100
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_HAlignment",
                    DataPropertyName = "HAlignment",
                    HeaderText = "水平对齐",
                    Width = 100
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_BreakBefore",
                    DataPropertyName = "BreakBefore",
                    HeaderText = "段前分页",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 50
                }
            });
            
            Dta_StyleList.DataSource = LevelStyles;
            Dta_StyleList.CellFormatting += Dta_StyleList_CellFormatting;
            Dta_StyleList.SelectionChanged += Dta_StyleList_SelectionChanged;
        }

        private void Dta_StyleList_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex != Dta_StyleList.Columns["Col_FontColor"].Index || e.RowIndex < 0 || !(Dta_StyleList.Rows[e.RowIndex].DataBoundItem is WordStyleInfo wordStyleInfo))
            {
                return;
            }
            using (Bitmap bitmap = new Bitmap(16, 16))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(wordStyleInfo.FontColor);
                e.Value = new Bitmap(bitmap);
            }
        }

        private void Dta_StyleList_SelectionChanged(object sender, EventArgs e)
        {
            if (Dta_StyleList.SelectedRows.Count > 0)
            {
                SetValueByStyle(LevelStyles[Dta_StyleList.SelectedRows[0].Index]);
            }
        }

        private void SetValueByStyle(WordStyleInfo style)
        {
            userChange = false;
            
            // 使用通用方法设置控件值
            SetComboBoxSelection(Cmb_ChnFontName, style.ChnFontName, FontNames);
            SetComboBoxSelection(Cmb_EngFontName, style.EngFontName, FontNames);
            SetComboBoxSelection(Cmb_FontSize, style.FontSize, WordStyleInfo.FontSizes);
            SetComboBoxSelection(Cmb_LineSpacing, style.LineSpace, WordStyleInfo.LineSpacings);
            SetComboBoxSelection(Cmb_SpaceBefore, style.SpaceBefore, WordStyleInfo.SpaceBeforeValues);
            SetComboBoxSelection(Cmb_SpaceAfter, style.SpaceAfter, WordStyleInfo.SpaceAfterValues);
            SetComboBoxSelection(Cmb_Alignment, style.HAlignment, WordStyleInfo.HAlignments);
            
            // 设置其他控件
            Btn_FontColor.BackColor = style.FontColor;
            Btn_Bold.Pressed = style.Bold;
            Btn_Italic.Pressed = style.Italic;
            Btn_Underline.Pressed = style.Underline;
            Btn_BreakBefore.Pressed = style.BreakBefore;
            
            // 设置缩进控件（从字符串解析数值）
            SetIndentValue(Txt_LeftIndent, style.LeftIndent);
            SetIndentValue(Txt_RightIndent, style.RightIndent);
            
            userChange = true;
        }

        private void ToggleButton_PressedChanged(object sender, EventArgs e)
        {
            if (sender is ToggleButton toggleButton)
            {
                toggleButton.Text = toggleButton.Pressed ? "是" : "否";
            }
        }

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
                
                // 立即更新当前选中的样式
                if (Dta_StyleList.SelectedRows.Count > 0)
                {
                    foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
                    {
                        var style = LevelStyles[selectedRow.Index];
                        style.FontColor = colorDialog.Color;
                        
                        // 强制刷新 DataGridView 显示
                        Dta_StyleList.InvalidateRow(selectedRow.Index);
                    }
                }
            }
        }


        private void Cmb_FontSize_Validated(object sender, EventArgs e)
        {
            ValidateAndFormatText(Cmb_FontSize, "磅");
        }

        private void Txt_Indent_ValueChanged(object sender, EventArgs e)
        {
            // StandardNumericUpDown的ValueChanged事件由OnControlValueChanged处理
            // 这里保留方法以保持设计器兼容性
        }

        private void Cmb_LineSpace_Validated(object sender, EventArgs e)
        {
            ValidateAndFormatText(Cmb_LineSpacing, Cmb_LineSpacing.Text.EndsWith("行") ? "行" : "磅");
        }

        private void Cmb_SpaceValue_Validated(object sender, EventArgs e)
        {
            if (sender is StandardComboBox comboBox)
            {
                ValidateAndFormatText(comboBox, comboBox.Text.EndsWith("行") ? "行" : "磅");
            }
        }

        private void Btn_SetStyles_Click(object sender, EventArgs e)
        {
            // 只保存样式设置，不立即应用到文档
            // 样式将在应用多级列表时生效
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        #region 设计器事件处理方法（保持兼容性）


        private void Cmb_SpaceValue_TextChanged(object sender, EventArgs e)
        {
            // 这个方法现在由OnControlValueChanged处理，但保留以保持设计器兼容性
        }

        private void Btn_FontColor_BackColorChanged(object sender, EventArgs e)
        {
            // 这个方法现在由OnControlValueChanged处理，但保留以保持设计器兼容性
        }

        private void Cmb_FontNameAndHV_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 这个方法现在由OnControlValueChanged处理，但保留以保持设计器兼容性
        }

        private void Cmb_FontSize_TextChanged(object sender, EventArgs e)
        {
            // 这个方法现在由OnControlValueChanged处理，但保留以保持设计器兼容性
        }

        private void Cmb_LineSpace_TextChanged(object sender, EventArgs e)
        {
            // 这个方法现在由OnControlValueChanged处理，但保留以保持设计器兼容性
        }

        #endregion

        private void label9_Click(object sender, EventArgs e)
        {
            // 段落行距标签点击事件 - 可以添加帮助信息或重置功能
        }

    }
}
