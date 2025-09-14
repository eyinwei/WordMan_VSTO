using System;
using System.Windows.Forms;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式设置窗体
    /// 使用StyleSettingsUIDesigner来创建和管理样式设置界面
    /// </summary>
    public partial class StyleSettingsForm : Form
    {
        private StyleSettingsUIDesigner _uiDesigner;

        public StyleSettingsForm()
        {
            InitializeComponent();
            InitializeUI();
        }

        /// <summary>
        /// 初始化UI
        /// </summary>
        private void InitializeUI()
        {
            _uiDesigner = new StyleSettingsUIDesigner(this);
            _uiDesigner.InitializeAllControls();
            
            // 绑定事件
            BindEvents();
        }

        /// <summary>
        /// 绑定控件事件
        /// </summary>
        private void BindEvents()
        {
            // Word字体设置按钮
            var btnWordFont = _uiDesigner.GetControl<Button>("btnWordFont");
            if (btnWordFont != null)
            {
                btnWordFont.Click += (s, e) => _uiDesigner.ShowWordFontDialog();
            }


            // Word段落设置按钮
            var btnWordParagraph = _uiDesigner.GetControl<Button>("btnWordParagraph");
            if (btnWordParagraph != null)
            {
                btnWordParagraph.Click += (s, e) => _uiDesigner.ShowWordParagraphDialog();
            }

            // 字体颜色按钮
            var btnFontColor = _uiDesigner.GetControl<Button>("btnFontColor");
            if (btnFontColor != null)
            {
                btnFontColor.Click += (s, e) => 
                {
                    var currentColor = btnFontColor.BackColor;
                    var newColor = _uiDesigner.ShowWordColorDialog(currentColor);
                    btnFontColor.BackColor = newColor;
                };
            }

            // 样式列表选择事件
            var styleList = _uiDesigner.GetControl<ListBox>("lstStyleList");
            if (styleList != null)
            {
                styleList.SelectedIndexChanged += (s, e) => 
                {
                    if (styleList.SelectedItem != null)
                    {
                        // 更新当前样式显示
                        var currentStyleLabel = _uiDesigner.GetControl<Label>("lblCurrentStyleName");
                        if (currentStyleLabel != null)
                        {
                            currentStyleLabel.Text = styleList.SelectedItem.ToString();
                        }
                        // 更新样式预览
                        _uiDesigner.UpdateStylePreview();
                    }
                };
            }

            // 样式预览更新事件
            var controls = new[] { "cmbChnFontName", "cmbEngFontName", "cmbFontSize", "cmbAlignment", 
                                 "txtFirstIndent", "cmbLineSpace", "txtSpaceBefore", "txtSpaceAfter",
                                 "chkBold", "chkItalic", "chkUnderline", "chkPageBreakBefore" };
            
            foreach (var controlName in controls)
            {
                var control = _uiDesigner.GetControl<Control>(controlName);
                if (control != null)
                {
                    if (control is ComboBox comboBox)
                    {
                        comboBox.SelectedIndexChanged += (s, e) => _uiDesigner.UpdateStylePreview();
                    }
                    else if (control is TextBox textBox)
                    {
                        textBox.TextChanged += (s, e) => _uiDesigner.UpdateStylePreview();
                    }
                    else if (control is CheckBox checkBox)
                    {
                        checkBox.CheckedChanged += (s, e) => _uiDesigner.UpdateStylePreview();
                    }
                }
            }

            // 应用样式按钮
            var btnApplyStyle = _uiDesigner.GetControl<Button>("btnApplyStyle");
            if (btnApplyStyle != null)
            {
                btnApplyStyle.Click += (s, e) => ApplyStyles();
            }

            // 取消按钮
            var btnCancel = _uiDesigner.GetControl<Button>("btnCancel");
            if (btnCancel != null)
            {
                btnCancel.Click += (s, e) => this.Close();
            }

            // 保存样式按钮
            var btnSaveStyle = _uiDesigner.GetControl<Button>("btnSaveStyle");
            if (btnSaveStyle != null)
            {
                btnSaveStyle.Click += (s, e) => SaveStyles();
            }

            // 读取样式按钮
            var btnLoadStyle = _uiDesigner.GetControl<Button>("btnLoadStyle");
            if (btnLoadStyle != null)
            {
                btnLoadStyle.Click += (s, e) => LoadStyles();
            }

            // 读取当前文档样式按钮
            var btnReadCurrent = _uiDesigner.GetControl<Button>("btnReadCurrent");
            if (btnReadCurrent != null)
            {
                btnReadCurrent.Click += (s, e) => ReadCurrentDocumentStyles();
            }

            // 添加样式按钮
            var btnAddStyle = _uiDesigner.GetControl<Button>("btnAddStyle");
            if (btnAddStyle != null)
            {
                btnAddStyle.Click += (s, e) => AddStyle();
            }

            // 删除样式按钮
            var btnDeleteStyle = _uiDesigner.GetControl<Button>("btnDeleteStyle");
            if (btnDeleteStyle != null)
            {
                btnDeleteStyle.Click += (s, e) => DeleteStyle();
            }
        }

        /// <summary>
        /// 应用样式
        /// </summary>
        private void ApplyStyles()
        {
            try
            {
                // 这里实现应用样式的逻辑
                MessageBox.Show("样式应用成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 保存样式
        /// </summary>
        private void SaveStyles()
        {
            try
            {
                // 这里实现保存样式的逻辑
                MessageBox.Show("样式保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 读取样式
        /// </summary>
        private void LoadStyles()
        {
            try
            {
                // 这里实现读取样式的逻辑
                MessageBox.Show("样式读取成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 读取当前文档样式
        /// </summary>
        private void ReadCurrentDocumentStyles()
        {
            try
            {
                // 这里实现读取当前文档样式的逻辑
                MessageBox.Show("当前文档样式读取成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取当前文档样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 添加样式
        /// </summary>
        private void AddStyle()
        {
            try
            {
                var txtNewStyleName = _uiDesigner.GetControl<TextBox>("txtNewStyleName");
                if (txtNewStyleName != null)
                {
                    var styleName = txtNewStyleName.Text.Trim();
                    if (string.IsNullOrEmpty(styleName) || styleName == "输入添加样式的名称")
                    {
                        MessageBox.Show("请输入样式名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    _uiDesigner.AddStyleToList(styleName);
                    txtNewStyleName.Text = "输入添加样式的名称";
                    MessageBox.Show($"样式 '{styleName}' 添加成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 删除样式
        /// </summary>
        private void DeleteStyle()
        {
            try
            {
                var selectedStyleName = _uiDesigner.GetSelectedStyleName();
                if (string.IsNullOrEmpty(selectedStyleName))
                {
                    MessageBox.Show("请先选择要删除的样式！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var result = MessageBox.Show($"确定要删除样式 '{selectedStyleName}' 吗？", "确认删除", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    _uiDesigner.RemoveStyleFromList(selectedStyleName);
                    MessageBox.Show($"样式 '{selectedStyleName}' 删除成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
