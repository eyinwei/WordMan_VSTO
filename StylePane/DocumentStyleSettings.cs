using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordMan_VSTO.StylePane
{
    public partial class DocumentStyleSettings : UserControl
    {
        private readonly List<WordStyleInfo> Styles = new List<WordStyleInfo>();
        private readonly List<string> FontNames = new List<string>();
        private bool userChange;

        public DocumentStyleSettings()
        {
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

            // 初始化样式列表
            InitializeStyles();
            
            // 初始化控件
            InitializeControls();
            
            userChange = true;
        }

        private void InitializeStyles()
        {
            try
            {
                // 获取Word内置样式
                for (int i = 1; i <= 9; i++)
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
                    
                    Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
                    object Index = wdBuiltinStyle;
                    Style style = styles[ref Index];
                    Styles.Add(new WordStyleInfo(style, wdBuiltinStyle));
                }

                // 添加正文样式
                Styles styles2 = Globals.ThisAddIn.Application.ActiveDocument.Styles;
                object Index2 = WdBuiltinStyle.wdStyleNormal;
                Style normalStyle = styles2[ref Index2];
                Styles.Add(new WordStyleInfo(normalStyle, WdBuiltinStyle.wdStyleNormal));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeControls()
        {
            // 初始化样式列表
            Lst_Styles.Items.Clear();
            foreach (WordStyleInfo style in Styles)
            {
                Lst_Styles.Items.Add(style.StyleName);
            }
            if (Lst_Styles.Items.Count > 0)
            {
                Lst_Styles.SelectedIndex = 0;
                LoadStyleForEditing(0);
            }

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

        private void LoadStyleForEditing(int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= Styles.Count)
                return;

            WordStyleInfo style = Styles[styleIndex];
            
            userChange = false;
            
            // 加载字体设置
            Cmb_ChnFontName.Text = style.ChnFontName;
            Cmb_EngFontName.Text = style.EngFontName;
            Cmb_FontSize.Text = style.FontSize;
            
            // 加载字体样式
            Chk_Bold.Checked = style.Bold;
            Chk_Italic.Checked = style.Italic;
            Chk_Underline.Checked = style.Underline;
            
            // 加载段落设置
            Cmb_Alignment.Text = style.HAlignment;
            Cmb_LineSpacing.Text = style.LineSpace;
            Cmb_SpaceBefore.Text = style.SpaceBefore;
            Cmb_SpaceAfter.Text = style.SpaceAfter;
            
            // 加载缩进设置
            Txt_LeftIndent.Text = style.LeftIndent;
            Txt_RightIndent.Text = style.RightIndent;
            Txt_FirstLineIndent.Text = style.FirstLineIndent;
            
            userChange = true;
        }

        private void ApplyStyleToDocument(int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= Styles.Count)
                return;

            WordStyleInfo style = Styles[styleIndex];
            
            // 应用字体设置
            style.ChnFontName = Cmb_ChnFontName.Text;
            style.EngFontName = Cmb_EngFontName.Text;
            style.FontSize = Cmb_FontSize.Text;
            
            // 应用字体样式
            style.Bold = Chk_Bold.Checked;
            style.Italic = Chk_Italic.Checked;
            style.Underline = Chk_Underline.Checked;
            
            // 应用段落设置
            style.HAlignment = Cmb_Alignment.Text;
            style.LineSpace = Cmb_LineSpacing.Text;
            style.SpaceBefore = Cmb_SpaceBefore.Text;
            style.SpaceAfter = Cmb_SpaceAfter.Text;
            
            // 应用缩进设置
            style.LeftIndent = Txt_LeftIndent.Text;
            style.RightIndent = Txt_RightIndent.Text;
            style.FirstLineIndent = Txt_FirstLineIndent.Text;
        }

        private void Lst_Styles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (userChange && Lst_Styles.SelectedIndex >= 0)
            {
                LoadStyleForEditing(Lst_Styles.SelectedIndex);
            }
        }

        private void Btn_Apply_Click(object sender, EventArgs e)
        {
            if (Lst_Styles.SelectedIndex >= 0)
            {
                ApplyStyleToDocument(Lst_Styles.SelectedIndex);
                
                // 应用样式到Word文档
                try
                {
                    WordStyleInfo style = Styles[Lst_Styles.SelectedIndex];
                    style.SetStyle(Globals.ThisAddIn.Application.ActiveDocument);
                    MessageBox.Show($"样式 '{style.StyleName}' 已应用", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"应用样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Btn_ApplyAll_Click(object sender, EventArgs e)
        {
            try
            {
                // 应用所有样式
                for (int i = 0; i < Styles.Count; i++)
                {
                    ApplyStyleToDocument(i);
                    Styles[i].SetStyle(Globals.ThisAddIn.Application.ActiveDocument);
                }
                MessageBox.Show($"所有 {Styles.Count} 个样式已应用", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用所有样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_Reset_Click(object sender, EventArgs e)
        {
            if (Lst_Styles.SelectedIndex >= 0)
            {
                LoadStyleForEditing(Lst_Styles.SelectedIndex);
            }
        }

        private void Btn_ResetAll_Click(object sender, EventArgs e)
        {
            InitializeStyles();
            if (Lst_Styles.SelectedIndex >= 0)
            {
                LoadStyleForEditing(Lst_Styles.SelectedIndex);
            }
        }

        private void Btn_Save_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "XML文件|*.xml";
                saveDialog.Title = "保存样式配置";
                
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    StyleSerializationHelper.SerializeListToXml(Styles, saveDialog.FileName);
                    MessageBox.Show("样式配置已保存", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存样式配置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btn_Load_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                openDialog.Filter = "XML文件|*.xml";
                openDialog.Title = "加载样式配置";
                
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    Styles.Clear();
                    Styles.AddRange(StyleSerializationHelper.DeserializeListFromXml<WordStyleInfo>(openDialog.FileName));
                    
                    // 刷新界面
                    InitializeControls();
                    MessageBox.Show("样式配置已加载", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载样式配置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
