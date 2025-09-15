using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
            _form.BackColor = Color.FromArgb(240, 248, 255); // 浅蓝色背景
            _form.FormBorderStyle = FormBorderStyle.FixedDialog;
            _form.MaximizeBox = false;
            _form.MinimizeBox = false;
            _form.Icon = SystemIcons.Information; // 使用信息图标，与其他窗体保持一致

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
                BackColor = Color.FromArgb(248, 250, 252), // 浅灰蓝色
                BorderStyle = BorderStyle.FixedSingle
            };

            // 样式列表
            var styleList = CreateStyleList();

            // 内置样式选择按钮
            var btnSelectBuiltIn = new Button
            {
                Name = "btnSelectBuiltIn",
                Text = "选择内置样式",
                Location = new Point(10, 390), // 向上平移10像素
                Size = new Size(180, 30),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                BackColor = Color.FromArgb(240, 240, 240),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _controls["btnSelectBuiltIn"] = btnSelectBuiltIn;

            // 输入添加样式的名称文本框
            var txtNewStyleName = new TextBox
            {
                Name = "txtNewStyleName",
                Location = new Point(10, 430), // 向上平移10像素，保持10像素间隔
                Size = new Size(180, 25),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                Text = "输入添加样式的名称",
                ForeColor = Color.Gray
            };
            
            // 添加输入事件处理
            txtNewStyleName.Enter += (sender, e) =>
            {
                if (txtNewStyleName.Text == "输入添加样式的名称")
                {
                    txtNewStyleName.Text = "";
                    txtNewStyleName.ForeColor = Color.Black;
                }
            };
            
            txtNewStyleName.Leave += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtNewStyleName.Text))
                {
                    txtNewStyleName.Text = "输入添加样式的名称";
                    txtNewStyleName.ForeColor = Color.Gray;
                }
            };
            
            _controls["txtNewStyleName"] = txtNewStyleName;

            // 添加样式按钮
            var btnAddStyle = new Button
            {
                Name = "btnAddStyle",
                Text = "添加样式",
                Location = new Point(10, 460), // 再向上平移5像素
                Size = new Size(85, 30),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                Location = new Point(105, 460), // 再向上平移5像素
                Size = new Size(85, 30),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                Size = new Size(180, 380), // 增大高度
                Font = new Font("微软雅黑", 10.5F), // 字体稍微增大
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

            // 添加样式列表，与任务窗格内容一致
            listBox.Items.AddRange(new string[] {
                "标题 1",
                "标题 2", 
                "标题 3",
                "标题 4",
                "标题 5",
                "标题 6",
                "正文",
                "题注",
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
                BackColor = Color.FromArgb(252, 254, 255), // 更浅的蓝色
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
            var lblChnFont = CreateLabel("中文字体", 20, 20);
            var cmbChnFont = CreateFontComboBox("cmbChnFontName", 100, 17);

            var lblEngFont = CreateLabel("西文字体", 350, 20);
            var cmbEngFont = CreateFontComboBox("cmbEngFontName", 430, 17);

            // 第二行：字体大小、颜色选择器和格式复选框
            var lblFontSize = CreateLabel("字体大小", 20, 55);
            var cmbFontSize = CreateSizeComboBox("cmbFontSize", 100, 52);

            // 字体颜色按钮（小方块，无标签）
            var btnFontColor = new Button
            {
                Name = "btnFontColor",
                Location = new Point(200, 52),
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
                Location = new Point(350, 55),
                Size = new Size(60, 20),
                Font = new Font("微软雅黑", 9.5F) // 字体稍微增大
            };
            _controls["chkBold"] = chkBold;

            var chkItalic = new CheckBox
            {
                Name = "chkItalic",
                Text = "斜体",
                Location = new Point(420, 55),
                Size = new Size(60, 20),
                Font = new Font("微软雅黑", 9.5F) // 字体稍微增大
            };
            _controls["chkItalic"] = chkItalic;

            var chkUnderline = new CheckBox
            {
                Name = "chkUnderline",
                Text = "下划线",
                Location = new Point(490, 55),
                Size = new Size(70, 20),
                Font = new Font("微软雅黑", 9.5F) // 字体稍微增大
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
                Font = new Font("微软雅黑", 9.5F) // 字体稍微增大
            };
            _controls["chkPageBreakBefore"] = chkPageBreakBefore;

            // 第二行：大纲级别和段落行距
            var lblOutlineLevel = CreateLabel("大纲级别", 20, 135);
            var cmbOutlineLevel = CreateOutlineLevelComboBox("cmbOutlineLevel", 100, 132);

            var lblLineSpace = CreateLabel("段落行距", 350, 135);
            var cmbLineSpace = CreateLineSpaceComboBox("cmbLineSpace", 430, 132);
            
            // 行距输入框（当选择固定值或最小值时显示）
            var txtLineSpaceValue = new TextBox
            {
                Name = "txtLineSpaceValue",
                Location = new Point(560, 132),
                Size = new Size(80, 25),
                Font = new Font("微软雅黑", 9.5F),
                Text = "12磅",
                Visible = false // 默认隐藏
            };
            _controls["txtLineSpaceValue"] = txtLineSpaceValue;

            // 第三行：缩进方式和缩进距离
            var lblIndentType = CreateLabel("缩进方式", 20, 170);
            var cmbIndentType = CreateIndentTypeComboBox("cmbIndentType", 100, 167);

            var lblIndentDistance = CreateLabel("缩进距离", 350, 170);
            var txtIndentDistance = new TextBox
            {
                Name = "txtIndentDistance",
                Location = new Point(430, 167),
                Size = new Size(100, 25),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                Text = "2字符"
            };
            _controls["txtIndentDistance"] = txtIndentDistance;

            // 为缩进距离输入框添加单位自动识别功能（不显示标签）
            txtIndentDistance.TextChanged += (sender, e) =>
            {
                // 单位自动识别和补充，但不显示标签
                AutoCompleteUnit(txtIndentDistance, new string[] { "字符", "厘米", "磅" });
            };

            // 第四行：段前间距和段后间距
            var lblSpaceBefore = CreateLabel("段前间距", 20, 205);
            var txtSpaceBefore = new TextBox
            {
                Name = "txtSpaceBefore",
                Location = new Point(100, 202),
                Size = new Size(100, 25),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                Text = "0.00行"
            };
            _controls["txtSpaceBefore"] = txtSpaceBefore;

            // 为段前间距输入框添加单位自动识别功能（不显示标签）
            txtSpaceBefore.TextChanged += (sender, e) =>
            {
                // 单位自动识别和补充，但不显示标签
                AutoCompleteUnit(txtSpaceBefore, new string[] { "行", "磅" });
            };

            var lblSpaceAfter = CreateLabel("段后间距", 350, 205);
            var txtSpaceAfter = new TextBox
            {
                Name = "txtSpaceAfter",
                Location = new Point(430, 202),
                Size = new Size(100, 25),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                Text = "0.00行"
            };
            _controls["txtSpaceAfter"] = txtSpaceAfter;

            // 为段后间距输入框添加单位自动识别功能（不显示标签）
            txtSpaceAfter.TextChanged += (sender, e) =>
            {
                // 单位自动识别和补充，但不显示标签
                AutoCompleteUnit(txtSpaceAfter, new string[] { "行", "磅" });
            };

            parentPanel.Controls.AddRange(new Control[] {
                lblAlignment, cmbAlignment, chkPageBreakBefore,
                lblOutlineLevel, cmbOutlineLevel, lblLineSpace, cmbLineSpace, txtLineSpaceValue,
                lblIndentType, cmbIndentType, lblIndentDistance, txtIndentDistance,
                lblSpaceBefore, txtSpaceBefore, lblSpaceAfter, txtSpaceAfter
            });
        }

        /// <summary>
        /// 创建样式预览
        /// </summary>
        private void CreateStylePreview(Control parentPanel)
        {
            // 样式预览标签
            var lblPreview = CreateLabel("样式预览", 20, 250);
            lblPreview.Font = new Font("微软雅黑", 9.5F, FontStyle.Bold); // 字体稍微增大

            // 样式预览文本框
            var txtPreview = new TextBox
            {
                Name = "txtStylePreview",
                Location = new Point(20, 275),
                Size = new Size(610, 140),
                Multiline = true,
                ReadOnly = true,
                Font = new Font("微软雅黑", 12.5F), // 字体稍微增大
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
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                DropDownStyle = ComboBoxStyle.DropDown, // 改为DropDown以支持键入匹配
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 20,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend, // 启用自动完成
                AutoCompleteSource = AutoCompleteSource.ListItems // 使用列表项作为自动完成源
            };

            // 通过Word API获取系统字体
            List<string> fontNames;
            try
            {
                fontNames = WordAPIHelper.GetSystemFonts();
                comboBox.Items.AddRange(fontNames.ToArray());
            }
            catch (Exception ex)
            {
                // 如果Word API失败，使用备用方法
                var installedFonts = new System.Drawing.Text.InstalledFontCollection();
                fontNames = new List<string>();
                
                foreach (FontFamily fontFamily in installedFonts.Families)
                {
                    fontNames.Add(fontFamily.Name);
                }
                
                // 按字母顺序排序
                fontNames.Sort();
                comboBox.Items.AddRange(fontNames.ToArray());
                System.Diagnostics.Debug.WriteLine($"使用备用字体获取方法：{ex.Message}");
            }
            
            // 设置默认选中项
            if (name == "cmbChnFontName")
            {
                var defaultIndex = fontNames.IndexOf("仿宋");
                comboBox.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
            }
            else
            {
                var defaultIndex = fontNames.IndexOf("Arial");
                comboBox.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 0;
            }

            // 设置字体绘制事件
            comboBox.DrawItem += (sender, e) =>
            {
                if (e.Index < 0) return;

                e.DrawBackground();
                
                // 设置选中状态的背景色
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), e.Bounds);
                }

                // 获取字体名称
                string fontName = comboBox.Items[e.Index].ToString();
                
                // 创建字体对象用于预览
                Font previewFont;
                try
                {
                    previewFont = new Font(fontName, 9F);
                }
                catch
                {
                    previewFont = new Font("微软雅黑", 9F);
                }

                // 绘制文本
                var textColor = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? Color.White : Color.FromArgb(64, 64, 64);
                var textRect = new Rectangle(e.Bounds.X + 5, e.Bounds.Y, e.Bounds.Width - 5, e.Bounds.Height);
                
                using (var brush = new SolidBrush(textColor))
                {
                    e.Graphics.DrawString(fontName, previewFont, brush, textRect, StringFormat.GenericDefault);
                }

                e.DrawFocusRectangle();
            };

            // 添加智能键入匹配功能
            comboBox.TextChanged += (sender, e) =>
            {
                var currentText = comboBox.Text;
                if (string.IsNullOrEmpty(currentText)) return;

                // 查找匹配的字体
                var matchedFont = FindMatchingFont(fontNames, currentText);
                if (matchedFont != null && matchedFont != currentText)
                {
                    // 如果找到匹配的字体，自动完成
                    comboBox.Text = matchedFont;
                    comboBox.SelectionStart = currentText.Length;
                    comboBox.SelectionLength = matchedFont.Length - currentText.Length;
                }
            };

            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建字号下拉框（类似Word的字体大小选择）
        /// </summary>
        private ComboBox CreateSizeComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(100, 25), // 增大宽度以容纳中文大小
                DropDownStyle = ComboBoxStyle.DropDown, // 支持键入匹配
                Font = new Font("微软雅黑", 9.5F),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            
            // 通过Word API获取字体大小选项
            List<string> sizes;
            try
            {
                sizes = WordAPIHelper.GetFontSizes();
                comboBox.Items.AddRange(sizes.ToArray());
            }
            catch (Exception ex)
            {
                // 如果Word API失败，使用备用选项
                string[] fallbackSizes = { 
                    "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", 
                    "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
                    "8", "9", "10", "10.5", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"
                };
                sizes = fallbackSizes.ToList();
                comboBox.Items.AddRange(fallbackSizes);
                System.Diagnostics.Debug.WriteLine($"使用备用字体大小选项：{ex.Message}");
            }
            
            // 设置默认选择（小四号，对应12磅）
            var defaultIndex = sizes.IndexOf("小四");
            comboBox.SelectedIndex = defaultIndex >= 0 ? defaultIndex : 5; // 如果找不到小四，则选择12
            
            // 添加智能大小匹配功能
            comboBox.TextChanged += (sender, e) =>
            {
                var currentText = comboBox.Text;
                if (string.IsNullOrEmpty(currentText)) return;

                var matchedSize = FindMatchingFontSize(sizes.ToArray(), currentText);
                if (matchedSize != null && matchedSize != currentText)
                {
                    comboBox.Text = matchedSize;
                    comboBox.SelectionStart = currentText.Length;
                    comboBox.SelectionLength = matchedSize.Length - currentText.Length;
                }
            };
            
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
                DropDownStyle = ComboBoxStyle.DropDownList, // 改回DropDownList，因为选项固定
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                DropDownStyle = ComboBoxStyle.DropDown, // 改为DropDown以支持键入匹配
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend, // 启用自动完成
                AutoCompleteSource = AutoCompleteSource.ListItems // 使用列表项作为自动完成源
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
                DropDownStyle = ComboBoxStyle.DropDownList, // 改回DropDownList，因为选项固定
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
        /// 创建大纲级别下拉框
        /// </summary>
        private ComboBox CreateOutlineLevelComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList, // 改回DropDownList，因为选项固定
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            comboBox.Items.AddRange(new string[] { "正文文本", "级别 1", "级别 2", "级别 3", "级别 4", "级别 5", "级别 6", "级别 7", "级别 8", "级别 9" });
            comboBox.SelectedIndex = 0; // 默认正文文本
            _controls[name] = comboBox;
            return comboBox;
        }

        /// <summary>
        /// 创建缩进方式下拉框
        /// </summary>
        private ComboBox CreateIndentTypeComboBox(string name, int x, int y)
        {
            var comboBox = new ComboBox
            {
                Name = name,
                Location = new Point(x, y),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList, // 改回DropDownList，因为选项固定
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Cursor = Cursors.Hand
            };
            comboBox.Items.AddRange(new string[] { "无", "首行缩进", "悬挂缩进" });
            comboBox.SelectedIndex = 1; // 默认首行缩进
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
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                Font = new Font("微软雅黑", 9.5F, FontStyle.Bold), // 字体稍微增大
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            // 预设样式下拉框
            var cmbPresetStyle = new ComboBox
            {
                Name = "cmbPresetStyle",
                Location = new Point(110, 18),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDown, // 改为DropDown以支持键入匹配
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
                AutoCompleteMode = AutoCompleteMode.SuggestAppend, // 启用自动完成
                AutoCompleteSource = AutoCompleteSource.ListItems // 使用列表项作为自动完成源
            };
            cmbPresetStyle.Items.AddRange(new string[] { "公文风格", "学术风格" });
            cmbPresetStyle.SelectedIndex = 0;
            _controls["cmbPresetStyle"] = cmbPresetStyle;

            // 读取样式按钮
            var btnLoadStyle = new Button
            {
                Name = "btnLoadStyle",
                Text = "读取样式",
                Location = new Point(250, 18),
                Size = new Size(80, 25),
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                Font = new Font("微软雅黑", 9.5F), // 字体稍微增大
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
                Font = new Font("微软雅黑", 10.5F, FontStyle.Bold), // 字体稍微增大
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
        /// 调用Word字体对话框（使用Word API）
        /// </summary>
        public void ShowWordFontDialog()
        {
            try
            {
                WordAPIHelper.ShowWordFontDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调用Word字体对话框失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 调用Word段落对话框（使用Word API）
        /// </summary>
        public void ShowWordParagraphDialog()
        {
            try
            {
                WordAPIHelper.ShowWordParagraphDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调用Word段落对话框失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 调用Word颜色对话框（使用Word API）
        /// </summary>
        public Color ShowWordColorDialog(Color currentColor)
        {
            try
            {
                return WordAPIHelper.ShowWordColorDialog(currentColor);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调用Word颜色对话框失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return currentColor;
            }
        }

        /// <summary>
        /// 应用段落格式到预览文本框
        /// </summary>
        private void ApplyParagraphFormatToPreview(TextBox previewTextBox, string alignment, string lineSpace, 
            string lineSpaceValue, string indentType, string indentDistance, string spaceBefore, 
            string spaceAfter, bool pageBreakBefore)
        {
            try
            {
                // 设置文本对齐方式
                switch (alignment)
                {
                    case "左对齐":
                        previewTextBox.TextAlign = HorizontalAlignment.Left;
                        break;
                    case "居中":
                        previewTextBox.TextAlign = HorizontalAlignment.Center;
                        break;
                    case "右对齐":
                        previewTextBox.TextAlign = HorizontalAlignment.Right;
                        break;
                    case "两端对齐":
                        previewTextBox.TextAlign = HorizontalAlignment.Left; // TextBox不支持两端对齐，使用左对齐
                        break;
                    default:
                        previewTextBox.TextAlign = HorizontalAlignment.Left;
                        break;
                }

                // 设置行距（通过调整字体大小来模拟）
                if (lineSpace == "单倍行距")
                {
                    // 保持当前字体大小
                }
                else if (lineSpace == "1.5倍行距")
                {
                    // 稍微增加字体大小来模拟1.5倍行距
                    var currentFont = previewTextBox.Font;
                    previewTextBox.Font = new Font(currentFont.FontFamily, currentFont.Size * 1.1f, currentFont.Style);
                }
                else if (lineSpace == "2倍行距")
                {
                    // 增加字体大小来模拟2倍行距
                    var currentFont = previewTextBox.Font;
                    previewTextBox.Font = new Font(currentFont.FontFamily, currentFont.Size * 1.2f, currentFont.Style);
                }
                else if (lineSpace == "固定值" || lineSpace == "最小值")
                {
                    // 根据固定值调整
                    if (!string.IsNullOrEmpty(lineSpaceValue))
                    {
                        try
                        {
                            var value = float.Parse(lineSpaceValue.Replace("磅", "").Replace("行", "").Trim());
                            var currentFont = previewTextBox.Font;
                            if (lineSpaceValue.Contains("磅"))
                            {
                                // 根据磅值调整字体大小
                                var newSize = Math.Max(8, Math.Min(24, value / 2)); // 简单的转换
                                previewTextBox.Font = new Font(currentFont.FontFamily, newSize, currentFont.Style);
                            }
                            else
                            {
                                // 根据行值调整
                                var newSize = currentFont.Size * value;
                                previewTextBox.Font = new Font(currentFont.FontFamily, Math.Max(8, Math.Min(24, newSize)), currentFont.Style);
                            }
                        }
                        catch
                        {
                            // 解析失败时保持当前字体
                        }
                    }
                }

                // 设置缩进（通过调整文本框的边距来模拟）
                var indentValue = 0;
                if (indentType == "首行缩进" || indentType == "悬挂缩进")
                {
                    try
                    {
                        var value = float.Parse(indentDistance.Replace("字符", "").Replace("厘米", "").Replace("磅", "").Trim());
                        if (indentDistance.Contains("字符"))
                        {
                            indentValue = (int)(value * 12); // 每个字符约12像素
                        }
                        else if (indentDistance.Contains("厘米"))
                        {
                            indentValue = (int)(value * 37.8); // 1厘米约37.8像素
                        }
                        else if (indentDistance.Contains("磅"))
                        {
                            indentValue = (int)(value * 1.33); // 1磅约1.33像素
                        }
                    }
                    catch
                    {
                        indentValue = 24; // 默认2字符缩进
                    }
                }

                // 应用缩进
                previewTextBox.Margin = new Padding(indentValue, 0, 0, 0);

                // 设置段前段后间距（通过调整文本框的上下边距来模拟）
                var spaceBeforeValue = 0;
                var spaceAfterValue = 0;

                try
                {
                    var beforeValue = float.Parse(spaceBefore.Replace("行", "").Replace("磅", "").Trim());
                    if (spaceBefore.Contains("行"))
                    {
                        spaceBeforeValue = (int)(beforeValue * 20); // 每行约20像素
                    }
                    else if (spaceBefore.Contains("磅"))
                    {
                        spaceBeforeValue = (int)(beforeValue * 1.33); // 1磅约1.33像素
                    }
                }
                catch
                {
                    spaceBeforeValue = 0;
                }

                try
                {
                    var afterValue = float.Parse(spaceAfter.Replace("行", "").Replace("磅", "").Trim());
                    if (spaceAfter.Contains("行"))
                    {
                        spaceAfterValue = (int)(afterValue * 20); // 每行约20像素
                    }
                    else if (spaceAfter.Contains("磅"))
                    {
                        spaceAfterValue = (int)(afterValue * 1.33); // 1磅约1.33像素
                    }
                }
                catch
                {
                    spaceAfterValue = 0;
                }

                // 应用间距
                var currentMargin = previewTextBox.Margin;
                previewTextBox.Margin = new Padding(currentMargin.Left, spaceBeforeValue, currentMargin.Right, spaceAfterValue);
            }
            catch (Exception ex)
            {
                // 段落格式应用失败时静默处理
                System.Diagnostics.Debug.WriteLine($"应用段落格式失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 更新样式预览（使用Word API）
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

                // 获取段落设置
                var alignment = GetControl<ComboBox>("cmbAlignment")?.Text ?? "左对齐";
                var lineSpace = GetControl<ComboBox>("cmbLineSpace")?.Text ?? "单倍行距";
                var lineSpaceValue = GetControl<TextBox>("txtLineSpaceValue")?.Text ?? "";
                var outlineLevel = GetControl<ComboBox>("cmbOutlineLevel")?.Text ?? "正文文本";
                var indentType = GetControl<ComboBox>("cmbIndentType")?.Text ?? "首行缩进";
                var indentDistance = GetControl<TextBox>("txtIndentDistance")?.Text ?? "2字符";
                var spaceBefore = GetControl<TextBox>("txtSpaceBefore")?.Text ?? "0.00行";
                var spaceAfter = GetControl<TextBox>("txtSpaceAfter")?.Text ?? "0.00行";
                var pageBreakBefore = GetControl<CheckBox>("chkPageBreakBefore")?.Checked ?? false;

                // 使用Word API创建样式预览
                try
                {
                    WordAPIHelper.CreateStylePreview(previewTextBox, chnFont, engFont, fontSize, 
                        isBold, isItalic, isUnderline, alignment, lineSpace, lineSpaceValue, 
                        outlineLevel, indentType, indentDistance, spaceBefore, spaceAfter, pageBreakBefore);
                }
                catch (Exception ex)
                {
                    // 如果Word API失败，使用备用方法
                    CreateFallbackPreview(previewTextBox, chnFont, engFont, fontSize, 
                        isBold, isItalic, isUnderline, alignment, lineSpace, lineSpaceValue, 
                        outlineLevel, indentType, indentDistance, spaceBefore, spaceAfter, pageBreakBefore);
                    System.Diagnostics.Debug.WriteLine($"使用备用预览方法：{ex.Message}");
                }

                // 构建样式信息文本
                var formatInfo = "";
                if (isBold) formatInfo += "粗体";
                if (isItalic) formatInfo += (formatInfo.Length > 0 ? "、" : "") + "斜体";
                if (isUnderline) formatInfo += (formatInfo.Length > 0 ? "、" : "") + "下划线";
                if (formatInfo.Length == 0) formatInfo = "常规";
                
                var lineSpaceInfo = lineSpace;
                if (!string.IsNullOrEmpty(lineSpaceValue) && (lineSpace == "固定值" || lineSpace == "最小值"))
                {
                    lineSpaceInfo += $" {lineSpaceValue}";
                }
                
                var spaceInfo = $"段前间距：{spaceBefore}，段后间距：{spaceAfter}";
                if (pageBreakBefore)
                {
                    spaceInfo += "，段前分页";
                }
                
                var styleInfo = $"字体：{chnFont}（中文）/{engFont}（西文），大小：{fontSize}号，格式：{formatInfo}，对齐：{alignment}，行距：{lineSpaceInfo}，大纲级别：{outlineLevel}，缩进：{indentType} {indentDistance}，{spaceInfo}。\r\n\r\n";
                
                // 将样式信息添加到预览文本前面
                previewTextBox.Text = styleInfo + previewTextBox.Text;
            }
            catch (Exception ex)
            {
                // 预览更新失败时静默处理
                System.Diagnostics.Debug.WriteLine($"更新样式预览失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 备用预览方法（当Word API不可用时）
        /// </summary>
        private void CreateFallbackPreview(TextBox previewTextBox, string chnFont, string engFont, 
            string fontSize, bool isBold, bool isItalic, bool isUnderline, string alignment, 
            string lineSpace, string lineSpaceValue, string outlineLevel, string indentType, 
            string indentDistance, string spaceBefore, string spaceAfter, bool pageBreakBefore)
        {
            try
            {
                // 创建字体
                var fontStyle = FontStyle.Regular;
                if (isBold) fontStyle |= FontStyle.Bold;
                if (isItalic) fontStyle |= FontStyle.Italic;
                if (isUnderline) fontStyle |= FontStyle.Underline;

                var font = new Font(new FontFamily(chnFont), WordAPIHelper.ConvertFontSize(fontSize), fontStyle);
                previewTextBox.Font = font;

                // 应用段落格式到预览文本框
                ApplyParagraphFormatToPreview(previewTextBox, alignment, lineSpace, lineSpaceValue, 
                    indentType, indentDistance, spaceBefore, spaceAfter, pageBreakBefore);

                // 设置预览文本
                previewTextBox.Text = "这是样式预览文本，将显示当前设置的字体、段落等效果。\r\n示例文字 示例文字 示例文字 示例文字 示例文字\r\n示例文字 示例文字 示例文字 示例文字 示例文字";
            }
            catch (Exception ex)
            {
                previewTextBox.Text = "样式预览不可用";
                System.Diagnostics.Debug.WriteLine($"备用预览方法失败：{ex.Message}");
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

        /// <summary>
        /// 智能查找匹配的字体名称
        /// 支持中文拼音匹配和英文模糊匹配
        /// </summary>
        private string FindMatchingFont(List<string> fontNames, string inputText)
        {
            if (string.IsNullOrEmpty(inputText) || fontNames == null || fontNames.Count == 0)
                return null;

            inputText = inputText.ToLower().Trim();

            // 1. 精确匹配
            var exactMatch = fontNames.FirstOrDefault(f => f.ToLower() == inputText);
            if (exactMatch != null) return exactMatch;

            // 2. 开头匹配
            var startMatch = fontNames.FirstOrDefault(f => f.ToLower().StartsWith(inputText));
            if (startMatch != null) return startMatch;

            // 3. 包含匹配
            var containsMatch = fontNames.FirstOrDefault(f => f.ToLower().Contains(inputText));
            if (containsMatch != null) return containsMatch;

            // 4. 中文拼音匹配（简单实现）
            var pinyinMatch = FindByPinyin(fontNames, inputText);
            if (pinyinMatch != null) return pinyinMatch;

            // 5. 模糊匹配（编辑距离）
            var fuzzyMatch = FindByFuzzyMatch(fontNames, inputText);
            if (fuzzyMatch != null) return fuzzyMatch;

            return null;
        }

        /// <summary>
        /// 通过拼音查找字体（简单实现）
        /// </summary>
        private string FindByPinyin(List<string> fontNames, string inputText)
        {
            // 常见中文字体的拼音映射
            var pinyinMap = new Dictionary<string, string>
            {
                { "fs", "仿宋" },
                { "fangsong", "仿宋" },
                { "kt", "楷体" },
                { "kaiti", "楷体" },
                { "st", "宋体" },
                { "songti", "宋体" },
                { "ht", "黑体" },
                { "heiti", "黑体" },
                { "msyh", "微软雅黑" },
                { "weiruanyahei", "微软雅黑" },
                { "yh", "雅黑" },
                { "yahei", "雅黑" }
            };

            foreach (var kvp in pinyinMap)
            {
                if (inputText.Contains(kvp.Key))
                {
                    var match = fontNames.FirstOrDefault(f => f.Contains(kvp.Value));
                    if (match != null) return match;
                }
            }

            return null;
        }

        /// <summary>
        /// 通过模糊匹配查找字体
        /// </summary>
        private string FindByFuzzyMatch(List<string> fontNames, string inputText)
        {
            string bestMatch = null;
            int minDistance = int.MaxValue;

            foreach (var fontName in fontNames)
            {
                var distance = CalculateLevenshteinDistance(inputText, fontName.ToLower());
                if (distance < minDistance && distance <= 3) // 允许最多3个字符的差异
                {
                    minDistance = distance;
                    bestMatch = fontName;
                }
            }

            return bestMatch;
        }

        /// <summary>
        /// 计算两个字符串的编辑距离
        /// </summary>
        private int CalculateLevenshteinDistance(string s1, string s2)
        {
            int[,] d = new int[s1.Length + 1, s2.Length + 1];

            for (int i = 0; i <= s1.Length; i++)
                d[i, 0] = i;

            for (int j = 0; j <= s2.Length; j++)
                d[0, j] = j;

            for (int i = 1; i <= s1.Length; i++)
            {
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }

            return d[s1.Length, s2.Length];
        }

        /// <summary>
        /// 智能查找匹配的字体大小
        /// 支持中文大小和数字大小匹配
        /// </summary>
        private string FindMatchingFontSize(string[] sizes, string inputText)
        {
            if (string.IsNullOrEmpty(inputText) || sizes == null || sizes.Length == 0)
                return null;

            inputText = inputText.ToLower().Trim();

            // 1. 精确匹配
            var exactMatch = sizes.FirstOrDefault(s => s.ToLower() == inputText);
            if (exactMatch != null) return exactMatch;

            // 2. 开头匹配
            var startMatch = sizes.FirstOrDefault(s => s.ToLower().StartsWith(inputText));
            if (startMatch != null) return startMatch;

            // 3. 包含匹配
            var containsMatch = sizes.FirstOrDefault(s => s.ToLower().Contains(inputText));
            if (containsMatch != null) return containsMatch;

            // 4. 中文大小拼音匹配
            var pinyinMatch = FindSizeByPinyin(sizes, inputText);
            if (pinyinMatch != null) return pinyinMatch;

            // 5. 数字大小匹配
            var numberMatch = FindSizeByNumber(sizes, inputText);
            if (numberMatch != null) return numberMatch;

            return null;
        }

        /// <summary>
        /// 通过拼音查找字体大小
        /// </summary>
        private string FindSizeByPinyin(string[] sizes, string inputText)
        {
            var pinyinMap = new Dictionary<string, string>
            {
                { "ch", "初号" },
                { "chuhao", "初号" },
                { "xc", "小初" },
                { "xiaochu", "小初" },
                { "yh", "一号" },
                { "yihao", "一号" },
                { "xy", "小一" },
                { "xiaoyi", "小一" },
                { "eh", "二号" },
                { "erhao", "二号" },
                { "xe", "小二" },
                { "xiaoer", "小二" },
                { "sh", "三号" },
                { "sanhao", "三号" },
                { "xs", "小三" },
                { "xiaosan", "小三" },
                { "sih", "四号" },
                { "sihao", "四号" },
                { "xs", "小四" },
                { "xiaosi", "小四" },
                { "wh", "五号" },
                { "wuhao", "五号" },
                { "xw", "小五" },
                { "xiaowu", "小五" },
                { "lh", "六号" },
                { "liuhao", "六号" },
                { "xl", "小六" },
                { "xiaoliu", "小六" },
                { "qh", "七号" },
                { "qihao", "七号" },
                { "bh", "八号" },
                { "bahao", "八号" }
            };

            foreach (var kvp in pinyinMap)
            {
                if (inputText.Contains(kvp.Key))
                {
                    var match = sizes.FirstOrDefault(s => s.Contains(kvp.Value));
                    if (match != null) return match;
                }
            }

            return null;
        }

        /// <summary>
        /// 通过数字查找字体大小
        /// </summary>
        private string FindSizeByNumber(string[] sizes, string inputText)
        {
            // 提取输入文本中的数字
            var numbers = System.Text.RegularExpressions.Regex.Matches(inputText, @"\d+(\.\d+)?");
            if (numbers.Count == 0) return null;

            var inputNumber = double.Parse(numbers[0].Value);
            
            // 查找最接近的数字大小
            string bestMatch = null;
            double minDifference = double.MaxValue;

            foreach (var size in sizes)
            {
                var sizeNumbers = System.Text.RegularExpressions.Regex.Matches(size, @"\d+(\.\d+)?");
                if (sizeNumbers.Count > 0)
                {
                    var sizeNumber = double.Parse(sizeNumbers[0].Value);
                    var difference = Math.Abs(inputNumber - sizeNumber);
                    if (difference < minDifference)
                    {
                        minDifference = difference;
                        bestMatch = size;
                    }
                }
            }

            return bestMatch;
        }

        /// <summary>
        /// 自动补充单位到输入框（不显示标签）
        /// </summary>
        private void AutoCompleteUnit(TextBox textBox, string[] validUnits)
        {
            try
            {
                var text = textBox.Text.Trim();
                if (string.IsNullOrEmpty(text)) return;

                // 检查是否已经包含单位
                var hasUnit = false;
                foreach (var unit in validUnits)
                {
                    if (text.EndsWith(unit))
                    {
                        hasUnit = true;
                        break;
                    }
                }

                // 如果没有单位，根据输入内容智能判断单位
                if (!hasUnit)
                {
                    var unit = DetectUnitFromText(text, validUnits);
                    if (!string.IsNullOrEmpty(unit))
                    {
                        // 自动补充单位到输入框
                        var numberPart = ExtractNumberFromText(text);
                        if (!string.IsNullOrEmpty(numberPart))
                        {
                            textBox.Text = numberPart + unit;
                            textBox.SelectionStart = numberPart.Length;
                            textBox.SelectionLength = unit.Length;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 静默处理错误，避免影响用户体验
                System.Diagnostics.Debug.WriteLine($"自动补充单位失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 更新单位标签，根据输入内容自动识别和补充单位
        /// </summary>
        private void UpdateUnitLabel(TextBox textBox, Label unitLabel, string[] validUnits)
        {
            try
            {
                var text = textBox.Text.Trim();
                if (string.IsNullOrEmpty(text)) return;

                // 检查是否已经包含单位
                var hasUnit = false;
                foreach (var unit in validUnits)
                {
                    if (text.EndsWith(unit))
                    {
                        hasUnit = true;
                        unitLabel.Text = unit;
                        break;
                    }
                }

                // 如果没有单位，根据输入内容智能判断单位
                if (!hasUnit)
                {
                    var unit = DetectUnitFromText(text, validUnits);
                    if (!string.IsNullOrEmpty(unit))
                    {
                        // 自动补充单位到输入框
                        var numberPart = ExtractNumberFromText(text);
                        if (!string.IsNullOrEmpty(numberPart))
                        {
                            textBox.Text = numberPart + unit;
                            textBox.SelectionStart = numberPart.Length;
                            textBox.SelectionLength = unit.Length;
                        }
                        unitLabel.Text = unit;
                    }
                }
            }
            catch (Exception ex)
            {
                // 静默处理错误，避免影响用户体验
                System.Diagnostics.Debug.WriteLine($"更新单位标签失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 从文本中检测单位（使用Word API）
        /// </summary>
        private string DetectUnitFromText(string text, string[] validUnits)
        {
            try
            {
                // 提取数字部分
                var numberMatch = System.Text.RegularExpressions.Regex.Match(text, @"\d+(\.\d+)?");
                if (!numberMatch.Success) return null;

                var number = double.Parse(numberMatch.Value);

                // 使用Word API进行单位检测
                return WordAPIHelper.DetectUnitFromNumber(number, validUnits);
            }
            catch (Exception ex)
            {
                // 如果Word API失败，使用备用方法
                System.Diagnostics.Debug.WriteLine($"使用备用单位检测方法：{ex.Message}");
                return DetectUnitFromTextFallback(text, validUnits);
            }
        }

        /// <summary>
        /// 备用单位检测方法
        /// </summary>
        private string DetectUnitFromTextFallback(string text, string[] validUnits)
        {
            // 提取数字部分
            var numberMatch = System.Text.RegularExpressions.Regex.Match(text, @"\d+(\.\d+)?");
            if (!numberMatch.Success) return null;

            var number = double.Parse(numberMatch.Value);

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

        /// <summary>
        /// 从文本中提取数字部分
        /// </summary>
        private string ExtractNumberFromText(string text)
        {
            var numberMatch = System.Text.RegularExpressions.Regex.Match(text, @"\d+(\.\d+)?");
            return numberMatch.Success ? numberMatch.Value : null;
        }
    }
}
