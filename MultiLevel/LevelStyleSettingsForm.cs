using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordMan_VSTO.StylePane;
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
            InstalledFontCollection installedFontCollection = new InstalledFontCollection();
            FontFamily[] families = installedFontCollection.Families;
            foreach (FontFamily fontFamily in families)
            {
                FontNames.Add(fontFamily.Name);
            }

            // 初始化级别样式
            InitializeLevelStyles();
            
            // 初始化控件
            InitializeControls();
            
            userChange = true;
        }

        private void InitializeLevelStyles()
        {
            try
            {
                // 获取Word内置标题样式
                for (int i = 1; i <= maxLevel; i++)
                {
                    WdBuiltinStyle wdBuiltinStyle;
                    switch (i)
                    {
                        case 1: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading1; break;
                        case 2: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading2; break;
                        case 3: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading3; break;
                        case 4: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading4; break;
                        case 5: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading5; break;
                        case 6: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading6; break;
                        case 7: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading7; break;
                        case 8: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading8; break;
                        case 9: wdBuiltinStyle = WdBuiltinStyle.wdStyleHeading9; break;
                        default: wdBuiltinStyle = (WdBuiltinStyle)0; break;
                    }
                    
                    Styles styles = WordAPIHelper.GetWordApplication().ActiveDocument.Styles;
                    object Index = wdBuiltinStyle;
                    Style style = styles[ref Index];
                    LevelStyles.Add(new WordStyleInfo(style, wdBuiltinStyle));
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
            
            // 初始化字体下拉框
            Cmb_ChnFontName.Items.Clear();
            Cmb_EngFontName.Items.Clear();
            foreach (string fontName in FontNames)
            {
                Cmb_ChnFontName.Items.Add(fontName);
                Cmb_EngFontName.Items.Add(fontName);
            }

            // 初始化字体大小下拉框
            Cmb_FontSize.Items.Clear();
            foreach (string fontSize in WordStyleInfo.FontSizes)
            {
                Cmb_FontSize.Items.Add(fontSize);
            }

            // 初始化对齐方式下拉框
            Cmb_Alignment.Items.Clear();
            foreach (string alignment in WordStyleInfo.HAlignments)
            {
                Cmb_Alignment.Items.Add(alignment);
            }

            // 初始化行距下拉框
            Cmb_LineSpacing.Items.Clear();
            foreach (string lineSpacing in WordStyleInfo.LineSpacings)
            {
                Cmb_LineSpacing.Items.Add(lineSpacing);
            }

            // 初始化段前距下拉框
            Cmb_SpaceBefore.Items.Clear();
            foreach (string spaceBefore in WordStyleInfo.SpaceBeforeValues)
            {
                Cmb_SpaceBefore.Items.Add(spaceBefore);
            }

            // 初始化段后距下拉框
            Cmb_SpaceAfter.Items.Clear();
            foreach (string spaceAfter in WordStyleInfo.SpaceAfterValues)
            {
                Cmb_SpaceAfter.Items.Add(spaceAfter);
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
                    Width = 100
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_ChnFontName",
                    DataPropertyName = "ChnFontName",
                    HeaderText = "中文字体",
                    Width = 120
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_EngFontName",
                    DataPropertyName = "EngFontName",
                    HeaderText = "西文字体",
                    Width = 120
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_FontSize",
                    DataPropertyName = "FontSize",
                    HeaderText = "字体大小",
                    Width = 80
                },
                new DataGridViewImageColumn
                {
                    Name = "Col_FontColor",
                    DataPropertyName = "FontColor",
                    HeaderText = "颜色",
                    ImageLayout = DataGridViewImageCellLayout.Normal,
                    Width = 60
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontBold",
                    DataPropertyName = "Bold",
                    HeaderText = "粗体",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 60
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_FontItalic",
                    DataPropertyName = "Italic",
                    HeaderText = "斜体",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 60
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
                    Width = 80
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_SpaceAfter",
                    DataPropertyName = "SpaceAfter",
                    HeaderText = "段后行距",
                    Width = 80
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Col_HAlignment",
                    DataPropertyName = "HAlignment",
                    HeaderText = "水平对齐",
                    Width = 80
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "Col_BreakBefore",
                    DataPropertyName = "BreakBefore",
                    HeaderText = "段前分页",
                    FalseValue = false,
                    TrueValue = true,
                    Width = 80
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
            int selectedIndex = FontNames.IndexOf(style.ChnFontName);
            Cmb_ChnFontName.SelectedIndex = selectedIndex;
            selectedIndex = FontNames.IndexOf(style.EngFontName);
            Cmb_EngFontName.SelectedIndex = selectedIndex;
            selectedIndex = WordStyleInfo.FontSizes.ToList().IndexOf(style.FontSize);
            if (selectedIndex != -1)
            {
                Cmb_FontSize.SelectedIndex = selectedIndex;
            }
            else
            {
                Cmb_FontSize.Text = style.FontSize;
            }
            Btn_FontColor.BackColor = style.FontColor;
            Btn_Bold.Pressed = style.Bold;
            Btn_Italic.Pressed = style.Italic;
            Btn_Underline.Pressed = style.Underline;
            Txt_LeftIndent.Text = style.LeftIndent;
            Txt_RightIndent.Text = style.RightIndent;
            selectedIndex = WordStyleInfo.LineSpacings.ToList().IndexOf(style.LineSpace);
            if (selectedIndex != -1)
            {
                Cmb_LineSpacing.SelectedIndex = selectedIndex;
            }
            else
            {
                Cmb_LineSpacing.Text = style.LineSpace;
            }
            selectedIndex = WordStyleInfo.SpaceBeforeValues.ToList().IndexOf(style.SpaceBefore);
            if (selectedIndex != -1)
            {
                Cmb_SpaceBefore.SelectedIndex = selectedIndex;
            }
            else
            {
                Cmb_SpaceBefore.Text = style.SpaceBefore;
            }
            selectedIndex = WordStyleInfo.SpaceAfterValues.ToList().IndexOf(style.SpaceAfter);
            if (selectedIndex != -1)
            {
                Cmb_SpaceAfter.SelectedIndex = selectedIndex;
            }
            else
            {
                Cmb_SpaceAfter.Text = style.SpaceAfter;
            }
            selectedIndex = WordStyleInfo.HAlignments.ToList().IndexOf(style.HAlignment);
            if (selectedIndex != -1)
            {
                Cmb_Alignment.SelectedIndex = selectedIndex;
            }
            else
            {
                Cmb_Alignment.SelectedIndex = 0;
            }
            Btn_BreakBefore.Pressed = style.BreakBefore;
            userChange = true;
        }

        private void ToggleButton_PressedChanged(object sender, EventArgs e)
        {
            if (!(sender is ToggleButton toggleButton))
            {
                return;
            }
            toggleButton.Text = (toggleButton.Pressed ? "是" : "否");
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            string columnName = "";
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                switch (toggleButton.Name)
                {
                    case "Btn_Bold":
                        LevelStyles[selectedRow.Index].Bold = toggleButton.Pressed;
                        columnName = "Col_FontBold";
                        break;
                    case "Btn_Italic":
                        LevelStyles[selectedRow.Index].Italic = toggleButton.Pressed;
                        columnName = "Col_FontItalic";
                        break;
                    case "Btn_Underline":
                        LevelStyles[selectedRow.Index].Underline = toggleButton.Pressed;
                        columnName = "Col_FontUnderline";
                        break;
                    case "Btn_BreakBefore":
                        LevelStyles[selectedRow.Index].BreakBefore = toggleButton.Pressed;
                        columnName = "Col_BreakBefore";
                        break;
                }
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
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
            }
        }

        private void Cmb_FontNameAndHV_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            string columnName = "";
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                switch ((sender as ComboBox).Name)
                {
                    case "Cmb_ChnFontName":
                        LevelStyles[selectedRow.Index].ChnFontName = (sender as ComboBox).SelectedItem.ToString();
                        columnName = "Col_ChnFontName";
                        break;
                    case "Cmb_EngFontName":
                        LevelStyles[selectedRow.Index].EngFontName = (sender as ComboBox).SelectedItem.ToString();
                        columnName = "Col_EngFontName";
                        break;
                    case "Cmb_Alignment":
                        LevelStyles[selectedRow.Index].HAlignment = (sender as ComboBox).SelectedItem.ToString();
                        columnName = "Col_HAlignment";
                        break;
                }
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
            }
        }

        private void Cmb_FontSize_TextChanged(object sender, EventArgs e)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                LevelStyles[selectedRow.Index].FontSize = Cmb_FontSize.Text;
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns["Col_FontSize"].Index, selectedRow.Index);
            }
        }

        private void Cmb_FontSize_Validated(object sender, EventArgs e)
        {
            int num = WordStyleInfo.FontSizes.ToList().IndexOf(Cmb_FontSize.Text);
            if (num != -1)
            {
                Cmb_FontSize.SelectedIndex = num;
            }
            else if (Regex.IsMatch(Cmb_FontSize.Text, "^\\d+(?:\\.(?:0|5))?(?:\\s+)?磅?$"))
            {
                string text = Cmb_FontSize.Text.TrimEnd(' ', '磅');
                if (float.TryParse(text, out float fontSize))
                {
                    Cmb_FontSize.Text = text + " 磅";
                }
            }
            else
            {
                Cmb_FontSize.SelectedIndex = 10; // 默认五号
            }
        }

        private void Btn_FontColor_BackColorChanged(object sender, EventArgs e)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                LevelStyles[selectedRow.Index].FontColor = Btn_FontColor.BackColor;
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns["Col_FontColor"].Index, selectedRow.Index);
            }
        }

        private void Txt_Indent_TextChanged(object sender, EventArgs e)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            string columnName = "";
            TextBox textBox = sender as TextBox;
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                string name = textBox.Name;
                if (name == "Txt_LeftIndent")
                {
                    LevelStyles[selectedRow.Index].LeftIndent = textBox.Text;
                    columnName = "Col_LeftIndent";
                }
                else if (name == "Txt_RightIndent")
                {
                    LevelStyles[selectedRow.Index].RightIndent = textBox.Text;
                    columnName = "Col_RightIndent";
                }
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
            }
        }

        private void Txt_Indent_Validated(object sender, EventArgs e)
        {
            TextBox textBox = sender as TextBox;
            string s = textBox.Text.TrimEnd(' ', '磅', '厘', '米');
            try
            {
                float num = float.Parse(s);
                if (textBox.Text.EndsWith("厘米"))
                {
                    textBox.Text = num.ToString("0.00 厘米");
                }
                else
                {
                    textBox.Text = num.ToString("0.00 磅");
                }
            }
            catch
            {
                textBox.Text = "0.00 厘米";
            }
        }

        private void Cmb_LineSpace_Validated(object sender, EventArgs e)
        {
            if (Cmb_LineSpacing.SelectedIndex != -1)
            {
                return;
            }
            string s = Cmb_LineSpacing.Text.TrimEnd(' ', '磅', '行');
            try
            {
                float num = float.Parse(s);
                if (Cmb_LineSpacing.Text.EndsWith("行"))
                {
                    Cmb_LineSpacing.Text = num.ToString("0.00 行");
                }
                else
                {
                    Cmb_LineSpacing.Text = num.ToString("0.00 磅");
                }
            }
            catch
            {
                Cmb_LineSpacing.SelectedIndex = 0; // 默认单倍行距
            }
        }

        private void Cmb_LineSpace_TextChanged(object sender, EventArgs e)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                LevelStyles[selectedRow.Index].LineSpace = Cmb_LineSpacing.Text;
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns["Col_LineSpace"].Index, selectedRow.Index);
            }
        }

        private void Cmb_SpaceValue_Validated(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex != -1)
            {
                return;
            }
            string s = comboBox.Text.TrimEnd(' ', '磅', '行');
            try
            {
                float num = float.Parse(s);
                if (comboBox.Text.EndsWith("行"))
                {
                    comboBox.Text = num.ToString("0.00 行");
                }
                else
                {
                    comboBox.Text = num.ToString("0.00 磅");
                }
            }
            catch
            {
                comboBox.SelectedIndex = 0; // 默认0.00行
            }
        }

        private void Cmb_SpaceValue_TextChanged(object sender, EventArgs e)
        {
            if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
            {
                return;
            }
            ComboBox comboBox = sender as ComboBox;
            string columnName = "";
            foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
            {
                string name = comboBox.Name;
                if (name == "Cmb_SpaceBefore")
                {
                    LevelStyles[selectedRow.Index].SpaceBefore = comboBox.Text;
                    columnName = "Col_SpaceBefore";
                }
                else if (name == "Cmb_SpaceAfter")
                {
                    LevelStyles[selectedRow.Index].SpaceAfter = comboBox.Text;
                    columnName = "Col_SpaceAfter";
                }
                Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
            }
        }

        private void Btn_SetStyles_Click(object sender, EventArgs e)
        {
            string text = string.Empty;
            foreach (WordStyleInfo style in LevelStyles)
            {
                if (!style.SetStyle(WordAPIHelper.GetWordApplication().ActiveDocument))
                {
                    text = text + style.StyleName + ";";
                }
            }
            if (!string.IsNullOrEmpty(text))
            {
                MessageBox.Show("样式：" + text.TrimEnd(';') + " 引用设置时出现错误，请检查设置值是否正确！", "多级段落设置", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                MessageBox.Show("样式设置已应用到文档！", "多级段落设置", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Close();
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
