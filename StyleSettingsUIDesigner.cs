using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式设置窗口UI设计器
    /// 负责创建和管理样式设置窗口的所有UI控件
    /// 整体布局分为三块：左上角样式列表、右上角样式编辑、下方按钮区域
    /// </summary>
    public class StyleSettingsUIDesigner
    {
        private Form _form;
        private Dictionary<string, Control> _controls;

        public StyleSettingsUIDesigner(Form form)
        {
            _form = form;
            _controls = new Dictionary<string, Control>();
        }

        /// <summary>
        /// 初始化所有UI控件
        /// </summary>
        public void InitializeAllControls()
        {
            // 设置窗体基本属性
            _form.Size = new Size(920, 630); 
            _form.Text = "样式设置";
            _form.StartPosition = FormStartPosition.CenterScreen;
            _form.BackColor = Color.FromArgb(248, 249, 250);
            _form.FormBorderStyle = FormBorderStyle.FixedDialog;
            _form.MaximizeBox = false;
            _form.MinimizeBox = false;

            // 1. 第一块：左上角样式列表区域
            var leftPanel = CreateLeftStyleListPanel();
            _form.Controls.Add(leftPanel);

            // 2. 第二块：右上角样式编辑区域
            var rightPanel = CreateRightStyleEditPanel();
            _form.Controls.Add(rightPanel);

            // 3. 第三块：下方按钮区域
            var bottomPanel = CreateBottomButtonPanel();
            _form.Controls.Add(bottomPanel);
        }

        /// <summary>
        /// 创建第一块：左上角样式列表区域
        /// 包含样式列表、内置样式选择、添加/删除样式功能
        /// </summary>
        private Panel CreateLeftStyleListPanel()
        {
            var panel = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(200, 500),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 样式列表
            var styleList = CreateStyleList();

            // 内置样式选择按钮
            var btnSelectBuiltIn = new Button
            {
                Name = "btnSelectBuiltIn",
                Text = "选择内置样式",
                Location = new Point(10, 370),
                Size = new Size(180, 30),
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.FromArgb(240, 240, 240),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnSelectBuiltIn"] = btnSelectBuiltIn;

            // 输入添加样式的名称文本框
            var txtNewStyleName = new TextBox
            {
                Name = "txtNewStyleName",
                Location = new Point(10, 410),
                Size = new Size(180, 25),
                Font = new Font("微软雅黑", 9F),
                Text = "输入添加样式的名称"
            };
            _controls["txtNewStyleName"] = txtNewStyleName;

            // 添加样式按钮
            var btnAddStyle = new Button
            {
                Name = "btnAddStyle",
                Text = "添加样式",
                Location = new Point(10, 445),
                Size = new Size(85, 30),
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnAddStyle"] = btnAddStyle;

            // 删除样式按钮
            var btnDeleteStyle = new Button
            {
                Name = "btnDeleteStyle",
                Text = "删除样式",
                Location = new Point(105, 445),
                Size = new Size(85, 30),
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.FromArgb(196, 43, 28),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnDeleteStyle"] = btnDeleteStyle;

            panel.Controls.Add(styleList);
            panel.Controls.Add(btnSelectBuiltIn);
            panel.Controls.Add(txtNewStyleName);
            panel.Controls.Add(btnAddStyle);
            panel.Controls.Add(btnDeleteStyle);

            return panel;
        }

        /// <summary>
        /// 创建样式列表（ListBox）
        /// </summary>
        private ListBox CreateStyleList()
        {
            var listBox = new ListBox
            {
                Name = "lstStyleList",
                Location = new Point(10, 10),
                Size = new Size(180, 350),
                Font = new Font("微软雅黑", 10F),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                SelectionMode = SelectionMode.MultiExtended,
                ItemHeight = 25
            };

            // 设置样式
            listBox.DrawMode = DrawMode.OwnerDrawFixed;
            listBox.DrawItem += (sender, e) =>
            {
                if (e.Index < 0) return;

                e.DrawBackground();
                
                // 设置选中状态的背景色
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), e.Bounds);
                }
                else
                {
                    e.Graphics.FillRectangle(new SolidBrush(e.Index % 2 == 0 ? Color.White : Color.FromArgb(248, 249, 250)), e.Bounds);
                }

                // 绘制文本
                var text = listBox.Items[e.Index].ToString();
                var textColor = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? Color.White : Color.FromArgb(64, 64, 64);
                var textRect = new Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 10, e.Bounds.Height);
                
                using (var brush = new SolidBrush(textColor))
                {
                    e.Graphics.DrawString(text, listBox.Font, brush, textRect, StringFormat.GenericDefault);
                }

                e.DrawFocusRectangle();
            };

            // 添加示例样式
            listBox.Items.AddRange(new string[] {
                "正文",
                "标题 1",
                "标题 2", 
                "标题 3",
                "标题 4",
                "标题",
                "副标题",
                "附录标题",
                "表格标题",
                "插图标题",
                "表内文字"
            });

            _controls["lstStyleList"] = listBox;
            return listBox;
        }

        /// <summary>
        /// 创建第二块：右上角样式编辑区域
        /// 包含字体设置、段落设置等详细配置
        /// </summary>
        private Panel CreateRightStyleEditPanel()
        {
            var panel = new Panel
            {
                Location = new Point(220, 10),
                Size = new Size(670, 500),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // 样式设置组
            var styleSetupGroup = CreateStyleSetupGroup();
            styleSetupGroup.Location = new Point(10, 10);
            styleSetupGroup.Size = new Size(650, 480);

            panel.Controls.Add(styleSetupGroup);

            return panel;
        }


        /// <summary>
        /// 创建样式设置组
        /// </summary>
        private GroupBox CreateStyleSetupGroup()
        {
            var groupBox = new GroupBox
            {
                Text = "样式设置",
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            // 字体设置区域
            CreateFontControls(groupBox);

            // 段落设置区域
            CreateParagraphControls(groupBox);

            // 样式预览
            CreateStylePreview(groupBox);

            return groupBox;
        }

        /// <summary>
        /// 创建字体控件
        /// </summary>
        private void CreateFontControls(Control parentPanel)
        {
            // 第一行：中文字体和西文字体
            var lblChnFont = CreateLabel("中文字体", 20, 30);
            var cmbChnFont = CreateFontComboBox("cmbChnFontName", 100, 27);

            var lblEngFont = CreateLabel("西文字体", 350, 30);
            var cmbEngFont = CreateFontComboBox("cmbEngFontName", 430, 27);

            // 第二行：字体大小、颜色选择器和格式复选框
            var lblFontSize = CreateLabel("字体大小", 20, 65);
            var cmbFontSize = CreateSizeComboBox("cmbFontSize", 100, 62);

            // 字体颜色按钮（小方块，无标签）
            var btnFontColor = new Button
            {
                Name = "btnFontColor",
                Location = new Point(200, 62),
                Size = new Size(25, 25),
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(200, 200, 200) },
                Cursor = Cursors.Hand
            };
            _controls["btnFontColor"] = btnFontColor;

            // 格式复选框
            var chkBold = new CheckBox
            {
                Name = "chkBold",
                Text = "粗体",
                Location = new Point(350, 65),
                Size = new Size(60, 20),
                Font = new Font("微软雅黑", 9F)
            };
            _controls["chkBold"] = chkBold;

            var chkItalic = new CheckBox
            {
                Name = "chkItalic",
                Text = "斜体",
                Location = new Point(420, 65),
                Size = new Size(60, 20),
                Font = new Font("微软雅黑", 9F)
            };
            _controls["chkItalic"] = chkItalic;

            var chkUnderline = new CheckBox
            {
                Name = "chkUnderline",
                Text = "下划线",
                Location = new Point(490, 65),
                Size = new Size(70, 20),
                Font = new Font("微软雅黑", 9F)
            };
            _controls["chkUnderline"] = chkUnderline;

            parentPanel.Controls.AddRange(new Control[] {
                lblChnFont, cmbChnFont, lblEngFont, cmbEngFont,
                lblFontSize, cmbFontSize, btnFontColor,
                chkBold, chkItalic, chkUnderline
            });
        }

        /// <summary>
        /// 创建段落控件
        /// </summary>
        private void CreateParagraphControls(Control parentPanel)
        {
            // 第一行：段落对齐和段前分页
            var lblAlignment = CreateLabel("段落对齐", 20, 100);
            var cmbAlignment = CreateAlignmentComboBox("cmbAlignment", 100, 97);

            var chkPageBreakBefore = new CheckBox
            {
                Name = "chkPageBreakBefore",
                Text = "段前分页",
                Location = new Point(350, 100),
                Size = new Size(80, 20),
                Font = new Font("微软雅黑", 9F)
            };
            _controls["chkPageBreakBefore"] = chkPageBreakBefore;

            // 第二行：首行缩进和段落行距
            var lblFirstIndent = CreateLabel("首行缩进", 20, 135);
            var txtFirstIndent = new TextBox
            {
                Name = "txtFirstIndent",
                Location = new Point(100, 132),
                Size = new Size(120, 25),
                Font = new Font("微软雅黑", 9F),
                Text = "2字符"
            };
            _controls["txtFirstIndent"] = txtFirstIndent;

            var lblLineSpace = CreateLabel("段落行距", 350, 135);
            var cmbLineSpace = CreateLineSpaceComboBox("cmbLineSpace", 430, 132);

            // 第三行：段前间距和段后间距
            var lblSpaceBefore = CreateLabel("段前间距", 20, 170);
            var txtSpaceBefore = new TextBox
            {
                Name = "txtSpaceBefore",
                Location = new Point(100, 167),
                Size = new Size(120, 25),
                Font = new Font("微软雅黑", 9F),
                Text = "0.00行"
            };
            _controls["txtSpaceBefore"] = txtSpaceBefore;

            var lblSpaceAfter = CreateLabel("段后间距", 350, 170);
            var txtSpaceAfter = new TextBox
            {
                Name = "txtSpaceAfter",
                Location = new Point(430, 167),
                Size = new Size(120, 25),
                Font = new Font("微软雅黑", 9F),
                Text = "0.00行"
            };
            _controls["txtSpaceAfter"] = txtSpaceAfter;

            parentPanel.Controls.AddRange(new Control[] {
                lblAlignment, cmbAlignment, chkPageBreakBefore,
                lblFirstIndent, txtFirstIndent, lblLineSpace, cmbLineSpace,
                lblSpaceBefore, txtSpaceBefore, lblSpaceAfter, txtSpaceAfter
            });
        }

        /// <summary>
        /// 创建样式预览
        /// </summary>
        private void CreateStylePreview(Control parentPanel)
        {
            // 样式预览标签
            var lblPreview = CreateLabel("样式预览", 20, 210);
            lblPreview.Font = new Font("微软雅黑", 9F, FontStyle.Bold);

            // 样式预览文本框
            var txtPreview = new TextBox
            {
                Name = "txtStylePreview",
                Location = new Point(20, 235),
                Size = new Size(610, 140),
                Multiline = true,
                ReadOnly = true,
                Font = new Font("微软雅黑", 12F),
                Text = "这是样式预览文本，将显示当前设置的字体、段落等效果。",
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            _controls["txtStylePreview"] = txtPreview;

            parentPanel.Controls.Add(lblPreview);
            parentPanel.Controls.Add(txtPreview);
        }

        /// <summary>
        /// 创建标签
        /// </summary>
        private Label CreateLabel(string text, int x, int y)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(80, 20),
                Font = new Font("微软雅黑", 9F),
                ForeColor = Color.FromArgb(64, 64, 64),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        /// <summary>
        /// 创建文本框
        /// </summary>
        private TextBox CreateTextBox(string name, int x, int y)
        {
            var textBox = new TextBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(120, 25),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.White
            };
            _controls[name] = textBox;
            return textBox;
        }

        /// <summary>
        /// 创建字体下拉框
        /// </summary>
        private ComboBox CreateFontComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            // 添加常用字体
            string[] fonts = { "宋体", "黑体", "楷体", "仿宋", "微软雅黑", "Arial", "Times New Roman", "Calibri" };
            comboBox.Items.AddRange(fonts);
            comboBox.SelectedIndex = 0;
            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建字号下拉框
        /// </summary>
        private ComboBox CreateSizeComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(80, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            string[] sizes = { "8", "9", "10", "10.5", "11", "12", "14", "16", "18", "20", "22", "24" };
            comboBox.Items.AddRange(sizes);
            comboBox.SelectedIndex = 5; // 默认选择12号
            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建行距下拉框
        /// </summary>
        private ComboBox CreateLineSpaceComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            string[] lineSpaces = { "单倍行距", "1.5倍行距", "2倍行距", "最小值", "固定值", "多倍行距" };
            comboBox.Items.AddRange(lineSpaces);
            comboBox.SelectedIndex = 0;
            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建间距下拉框
        /// </summary>
        private ComboBox CreateSpaceComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(80, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            string[] spaces = { "0磅", "6磅", "12磅", "18磅", "24磅", "30磅" };
            comboBox.Items.AddRange(spaces);
            comboBox.SelectedIndex = 0;
            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建对齐方式下拉框
        /// </summary>
        private ComboBox CreateAlignmentComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            comboBox.Items.AddRange(new string[] { "左对齐", "居中", "右对齐", "两端对齐" });
            comboBox.SelectedIndex = 0; // 默认左对齐
            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建切换按钮
        /// </summary>
        private Button CreateToggleButton(string name, int x, int y, string text)
        {
            var button = new Button
            {
                Name = name,
                Text = text,
                Location = new Point(x, y),
                Size = new Size(50, 28),
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(64, 64, 64),
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(200, 200, 200) },
                Cursor = Cursors.Hand
            };
            _controls[name] = button;
            return button;
        }

        /// <summary>
        /// 创建颜色按钮
        /// </summary>
        private Button CreateColorButton(string name, int x, int y)
        {
            var button = new Button
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(50, 28),
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 1, BorderColor = Color.FromArgb(200, 200, 200) },
                Cursor = Cursors.Hand,
                Text = "颜色"
            };
            _controls[name] = button;
            return button;
        }

        /// <summary>
        /// 创建第三块：下方按钮区域
        /// 包含预设样式选择、读取/保存样式、应用样式等功能
        /// </summary>
        private Panel CreateBottomButtonPanel()
        {
            var panel = new Panel
            {
                Location = new Point(10, 520),
                Size = new Size(880, 60),
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            // 预设样式标签
            var lblPreset = new Label
            {
                Text = "预设样式：",
                Location = new Point(20, 20),
                Size = new Size(80, 20),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            // 预设样式下拉框
            var cmbPresetStyle = new ComboBox
            {
                Name = "cmbPresetStyle",
                Location = new Point(110, 18),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("微软雅黑", 9F)
            };
            cmbPresetStyle.Items.AddRange(new string[] { "公文风格", "默认", "学术风格", "现代风格" });
            cmbPresetStyle.SelectedIndex = 0;
            _controls["cmbPresetStyle"] = cmbPresetStyle;

            // 读取样式按钮
            var btnLoadStyle = new Button
            {
                Name = "btnLoadStyle",
                Text = "读取样式",
                Location = new Point(250, 18),
                Size = new Size(80, 25),
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.FromArgb(52, 144, 220),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnLoadStyle"] = btnLoadStyle;

            // 保存样式按钮
            var btnSaveStyle = new Button
            {
                Name = "btnSaveStyle",
                Text = "保存样式",
                Location = new Point(350, 18),
                Size = new Size(80, 25),
                Font = new Font("微软雅黑", 9F),
                BackColor = Color.FromArgb(52, 144, 220),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnSaveStyle"] = btnSaveStyle;

            // 应用样式按钮（最右下角）
            var btnApplyStyle = new Button
            {
                Name = "btnApplyStyle",
                Text = "应用样式",
                Location = new Point(760, 15),
                Size = new Size(100, 30),
                Font = new Font("微软雅黑", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnApplyStyle"] = btnApplyStyle;

            panel.Controls.Add(lblPreset);
            panel.Controls.Add(cmbPresetStyle);
            panel.Controls.Add(btnLoadStyle);
            panel.Controls.Add(btnSaveStyle);
            panel.Controls.Add(btnApplyStyle);

            return panel;
        }

        /// <summary>
        /// 获取指定名称的控件
        /// </summary>
        public T GetControl<T>(string name) where T : Control
        {
            if (_controls.ContainsKey(name))
                return _controls[name] as T;
            return null;
        }

        /// <summary>
        /// 获取所有控件
        /// </summary>
        public Dictionary<string, Control> GetAllControls()
        {
            return _controls;
        }

        /// <summary>
        /// 调用Word字体对话框
        /// </summary>
        public void ShowWordFontDialog()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app != null)
                {
                    // 使用Word的字体对话框
                    app.Dialogs[Word.WdWordDialog.wdDialogFormatFont].Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调用Word字体对话框失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// 调用Word段落对话框
        /// </summary>
        public void ShowWordParagraphDialog()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app != null)
                {
                    // 使用Word的段落对话框
                    app.Dialogs[Word.WdWordDialog.wdDialogFormatParagraph].Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调用Word段落对话框失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 调用Word颜色对话框
        /// </summary>
        public Color ShowWordColorDialog(Color currentColor)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app != null)
                {
                    // 使用Word的颜色对话框
                    var colorDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatFont];
                    colorDialog.Show();
                    // 这里需要根据实际Word API来获取选择的颜色
                    return currentColor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调用Word颜色对话框失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return currentColor;
        }

        /// <summary>
        /// 更新样式预览
        /// </summary>
        public void UpdateStylePreview()
        {
            try
            {
                var previewTextBox = GetControl<TextBox>("txtStylePreview");
                if (previewTextBox == null) return;

                // 获取当前字体设置
                var chnFont = GetControl<ComboBox>("cmbChnFontName")?.Text ?? "仿宋";
                var engFont = GetControl<ComboBox>("cmbEngFontName")?.Text ?? "Arial";
                var fontSize = GetControl<ComboBox>("cmbFontSize")?.Text ?? "12";
                var isBold = GetControl<CheckBox>("chkBold")?.Checked ?? false;
                var isItalic = GetControl<CheckBox>("chkItalic")?.Checked ?? false;
                var isUnderline = GetControl<CheckBox>("chkUnderline")?.Checked ?? false;

                // 创建字体
                var fontStyle = FontStyle.Regular;
                if (isBold) fontStyle |= FontStyle.Bold;
                if (isItalic) fontStyle |= FontStyle.Italic;
                if (isUnderline) fontStyle |= FontStyle.Underline;

                var font = new Font(chnFont, float.Parse(fontSize), fontStyle);
                previewTextBox.Font = font;

                // 更新预览文本
                var alignment = GetControl<ComboBox>("cmbAlignment")?.Text ?? "左对齐";
                var lineSpace = GetControl<ComboBox>("cmbLineSpace")?.Text ?? "单倍行距";
                var firstIndent = GetControl<TextBox>("txtFirstIndent")?.Text ?? "2字符";
                var spaceBefore = GetControl<TextBox>("txtSpaceBefore")?.Text ?? "0.00行";
                var spaceAfter = GetControl<TextBox>("txtSpaceAfter")?.Text ?? "0.00行";

                var previewText = $"字体：{chnFont} {fontSize}号\n" +
                                $"格式：{(isBold ? "粗体 " : "")}{(isItalic ? "斜体 " : "")}{(isUnderline ? "下划线 " : "")}\n" +
                                $"对齐：{alignment}\n" +
                                $"行距：{lineSpace}\n" +
                                $"缩进：{firstIndent}\n" +
                                $"间距：段前{spaceBefore} 段后{spaceAfter}";

                previewTextBox.Text = previewText;
            }
            catch (Exception ex)
            {
                // 预览更新失败时静默处理
                System.Diagnostics.Debug.WriteLine($"更新样式预览失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 添加样式到列表
        /// </summary>
        public void AddStyleToList(string styleName)
        {
            try
            {
                var styleList = GetControl<ListBox>("lstStyleList");
                if (styleList != null && !string.IsNullOrEmpty(styleName))
                {
                    if (!styleList.Items.Contains(styleName))
                    {
                        styleList.Items.Add(styleName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从列表中删除样式
        /// </summary>
        public void RemoveStyleFromList(string styleName)
        {
            try
            {
                var styleList = GetControl<ListBox>("lstStyleList");
                if (styleList != null && !string.IsNullOrEmpty(styleName))
                {
                    styleList.Items.Remove(styleName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取当前选中的样式名称
        /// </summary>
        public string GetSelectedStyleName()
        {
            try
            {
                var styleList = GetControl<ListBox>("lstStyleList");
                if (styleList != null && styleList.SelectedItem != null)
                {
                    return styleList.SelectedItem.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取选中样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return string.Empty;
        }

        /// <summary>
        /// 设置当前选中的样式
        /// </summary>
        public void SetSelectedStyle(string styleName)
        {
            try
            {
                var styleList = GetControl<ListBox>("lstStyleList");
                if (styleList != null && !string.IsNullOrEmpty(styleName))
                {
                    var index = styleList.Items.IndexOf(styleName);
                    if (index >= 0)
                    {
                        styleList.SelectedIndex = index;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置选中样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
