using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式设置窗口UI设计器
    /// 负责创建和管理样式设置窗口的所有UI控件
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
            // 1. 样式列表区域（包含当前样式显示）
            var styleListPanel = CreateStyleListPanel();
            _form.Controls.Add(styleListPanel);

            // 2. 样式编辑区域
            var styleEditGroup = CreateStyleEditGroup();
            _form.Controls.Add(styleEditGroup);

            // 3. 按钮面板
            InitializeButtonPanel();
        }

        /// <summary>
        /// 创建样式列表面板（包含当前样式显示和样式列表）
        /// </summary>
        private Panel CreateStyleListPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 300,
                BackColor = Color.White,
                Padding = new Padding(15)
            };

            // 当前样式显示区域
            var currentStylePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(240, 248, 255),
                BorderStyle = BorderStyle.FixedSingle
            };

            var currentStyleLabel = new Label
            {
                Text = "当前样式：",
                Location = new Point(15, 15),
                Size = new Size(80, 20),
                Font = new Font("微软雅黑", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            var currentStyleName = new Label
            {
                Name = "lblCurrentStyleName",
                Text = "公文风格",
                Location = new Point(100, 15),
                Size = new Size(200, 20),
                Font = new Font("微软雅黑", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215)
            };
            _controls["lblCurrentStyleName"] = currentStyleName;

            currentStylePanel.Controls.Add(currentStyleLabel);
            currentStylePanel.Controls.Add(currentStyleName);

            // 样式列表
            var styleList = CreateStyleList();
            styleList.Dock = DockStyle.Fill;
            styleList.Margin = new Padding(0, 10, 0, 0);

            panel.Controls.Add(styleList);
            panel.Controls.Add(currentStylePanel);

            return panel;
        }

        /// <summary>
        /// 创建样式列表（DataGridView）
        /// </summary>
        private DataGridView CreateStyleList()
        {
            var dataGridView = new DataGridView
            {
                Name = "dgvStyleList",
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing,
                ColumnHeadersHeight = 35,
                RowTemplate = { Height = 28 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = Color.FromArgb(230, 230, 230),
                EnableHeadersVisualStyles = false,
                Font = new Font("微软雅黑", 9F)
            };

            // 设置表头样式
            dataGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.White,
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleCenter
            };

            // 设置行样式
            dataGridView.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.White,
                ForeColor = Color.FromArgb(64, 64, 64),
                Font = new Font("微软雅黑", 9F),
                SelectionBackColor = Color.FromArgb(0, 120, 215),
                SelectionForeColor = Color.White
            };

            // 设置交替行颜色
            dataGridView.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(248, 249, 250)
            };

            // 添加列
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_StyleName",
                HeaderText = "样式名",
                Width = 120,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                Frozen = true
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_ChnFontName",
                HeaderText = "中文字体",
                Width = 100,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_EngFontName",
                HeaderText = "西文字体",
                Width = 100,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_FontSize",
                HeaderText = "字号",
                Width = 60,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Col_Bold",
                HeaderText = "粗体",
                Width = 60,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Col_Italic",
                HeaderText = "斜体",
                Width = 60,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Col_Underline",
                HeaderText = "下划线",
                Width = 60,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_LineSpace",
                HeaderText = "行距",
                Width = 80,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_SpaceBefore",
                HeaderText = "段前间距",
                Width = 80,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_SpaceAfter",
                HeaderText = "段后间距",
                Width = 80,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Col_HAlignment",
                HeaderText = "对齐方式",
                Width = 80,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            _controls["dgvStyleList"] = dataGridView;
            return dataGridView;
        }

        /// <summary>
        /// 创建样式编辑分组
        /// </summary>
        private GroupBox CreateStyleEditGroup()
        {
            var group = new GroupBox
            {
                Text = "样式编辑设置",
                Dock = DockStyle.Top,
                Height = 220,
                Font = new Font("微软雅黑", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64),
                BackColor = Color.White,
                Padding = new Padding(20, 25, 20, 20)
            };

            // 创建主容器面板
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            // 字体设置区域
            var fontPanel = CreateFontSettingsPanel();
            fontPanel.Location = new Point(0, 0);
            fontPanel.Size = new Size(480, 80);

            // 段落设置区域
            var paragraphPanel = CreateParagraphSettingsPanel();
            paragraphPanel.Location = new Point(0, 90);
            paragraphPanel.Size = new Size(480, 80);

            // 应用按钮
            var btnSetStyles = new Button
            {
                Text = "应用设置",
                Name = "btnSetStyles",
                Location = new Point(500, 20),
                Size = new Size(100, 35),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand
            };
            _controls["btnSetStyles"] = btnSetStyles;

            mainPanel.Controls.Add(fontPanel);
            mainPanel.Controls.Add(paragraphPanel);
            mainPanel.Controls.Add(btnSetStyles);

            group.Controls.Add(mainPanel);

            return group;
        }

        /// <summary>
        /// 创建字体设置面板
        /// </summary>
        private Panel CreateFontSettingsPanel()
        {
            var panel = new Panel
            {
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            // 标题
            var titleLabel = new Label
            {
                Text = "字体设置",
                Location = new Point(10, 8),
                Size = new Size(80, 20),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            // 中文字体
            var lblChnFont = CreateLabel("中文字体", 20, 35);
            var cmbChnFont = CreateFontComboBox("cmbChnFontName", 90, 32);

            // 西文字体
            var lblEngFont = CreateLabel("西文字体", 250, 35);
            var cmbEngFont = CreateFontComboBox("cmbEngFontName", 320, 32);

            // 字体大小
            var lblFontSize = CreateLabel("字体大小", 20, 55);
            var cmbFontSize = CreateSizeComboBox("cmbFontSize", 90, 32);

            // 字体颜色
            var lblFontColor = CreateLabel("字体颜色", 250, 55);
            var btnFontColor = CreateColorButton("btnFontColor", 320, 52);

            panel.Controls.AddRange(new Control[] {
                titleLabel, lblChnFont, cmbChnFont, lblEngFont, cmbEngFont,
                lblFontSize, cmbFontSize, lblFontColor, btnFontColor
            });

            return panel;
        }

        /// <summary>
        /// 创建段落设置面板
        /// </summary>
        private Panel CreateParagraphSettingsPanel()
        {
            var panel = new Panel
            {
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            // 标题
            var titleLabel = new Label
            {
                Text = "段落设置",
                Location = new Point(10, 8),
                Size = new Size(80, 20),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            // 行距
            var lblLineSpace = CreateLabel("行距", 20, 35);
            var cmbLineSpace = CreateLineSpaceComboBox("cmbLineSpace", 90, 32);

            // 段前间距
            var lblSpaceBefore = CreateLabel("段前间距", 200, 35);
            var cmbSpaceBefore = CreateSpaceComboBox("cmbSpaceBefore", 280, 32);

            // 段后间距
            var lblSpaceAfter = CreateLabel("段后间距", 20, 55);
            var cmbSpaceAfter = CreateSpaceComboBox("cmbSpaceAfter", 90, 32);

            // 对齐方式
            var lblHAlignment = CreateLabel("对齐方式", 200, 55);
            var cmbHAlignment = CreateAlignmentComboBox("cmbHAlignment", 280, 32);

            panel.Controls.AddRange(new Control[] {
                titleLabel, lblLineSpace, cmbLineSpace, lblSpaceBefore, cmbSpaceBefore,
                lblSpaceAfter, cmbSpaceAfter, lblHAlignment, cmbHAlignment
            });

            return panel;
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
                Size = new Size(70, 20),
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
                Size = new Size(100, 25),
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
                Size = new Size(120, 28),
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
                Size = new Size(80, 28),
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
                Size = new Size(100, 28),
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
                Size = new Size(100, 28),
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
        /// 初始化按钮面板
        /// </summary>
        private void InitializeButtonPanel()
        {
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                Padding = new Padding(20, 15, 20, 15),
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            var saveButton = new Button 
            { 
                Text = "另存样式", 
                Name = "btnSaveStyle", 
                Width = 90, 
                Height = 35,
                Margin = new Padding(5),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                BackColor = Color.FromArgb(52, 144, 220),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand
            };
            _controls["btnSaveStyle"] = saveButton;

            var loadButton = new Button 
            { 
                Text = "加载样式", 
                Name = "btnLoadStyle", 
                Width = 90, 
                Height = 35,
                Margin = new Padding(5),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                BackColor = Color.FromArgb(52, 144, 220),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand
            };
            _controls["btnLoadStyle"] = loadButton;

            var currentButton = new Button 
            { 
                Text = "读取当前文档样式", 
                Name = "btnCurrentStyle", 
                Width = 130, 
                Height = 35,
                Margin = new Padding(5),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                BackColor = Color.FromArgb(52, 144, 220),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand
            };
            _controls["btnCurrentStyle"] = currentButton;

            var applyButton = new Button 
            { 
                Text = "应用", 
                Name = "btnApply", 
                Width = 80, 
                Height = 35,
                Margin = new Padding(5),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand
            };
            _controls["btnApply"] = applyButton;

            var cancelButton = new Button 
            { 
                Text = "取消", 
                Name = "btnCancel", 
                Width = 80, 
                Height = 35,
                Margin = new Padding(5),
                Font = new Font("微软雅黑", 9F, FontStyle.Bold),
                BackColor = Color.FromArgb(196, 43, 28),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                Cursor = Cursors.Hand
            };
            _controls["btnCancel"] = cancelButton;

            var flowLayoutPanel = new FlowLayoutPanel 
            { 
                Dock = DockStyle.Fill, 
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 10, 0, 0)
            };
            
            flowLayoutPanel.Controls.Add(cancelButton);
            flowLayoutPanel.Controls.Add(applyButton);
            flowLayoutPanel.Controls.Add(currentButton);
            flowLayoutPanel.Controls.Add(loadButton);
            flowLayoutPanel.Controls.Add(saveButton);

            buttonPanel.Controls.Add(flowLayoutPanel);
            _form.Controls.Add(buttonPanel);
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
    }
}
