using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;

namespace WordMan_VSTO
{
    public class RoundedButton : Button
    {
        private int borderRadius = 8;
        private bool isHovered = false;

        public int BorderRadius
        {
            get { return borderRadius; }
            set { borderRadius = value; this.Invalidate(); }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            GraphicsPath path = new GraphicsPath();
            System.Drawing.Rectangle rect = new System.Drawing.Rectangle(0, 0, this.Width - 1, this.Height - 1);
            path = AddRoundedRectangle(rect, borderRadius);
            this.Region = new Region(path);
            
            // 绘制轮廓线
            Color borderColor = isHovered ? Color.FromArgb(100, 100, 100) : Color.FromArgb(200, 200, 200);
            using (Pen borderPen = new Pen(borderColor, 1))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(borderPen, path);
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            isHovered = true;
            this.Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            isHovered = false;
            this.Invalidate();
        }

        private GraphicsPath AddRoundedRectangle(System.Drawing.Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = radius * 2;
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public class NumericUpDownWithUnit : NumericUpDown
    {
        private string _label = "厘米";
        private Label _unitLabel;

        public string Label
        {
            get { return _label; }
            set
            {
                _label = value;
                if (_unitLabel != null)
                    _unitLabel.Text = value;
            }
        }

        public NumericUpDownWithUnit()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this._unitLabel = new Label();
            this._unitLabel.Text = _label;
            this._unitLabel.AutoSize = true;
            this._unitLabel.Font = new Font("Microsoft YaHei", 8F);
            this._unitLabel.ForeColor = Color.Gray;
            this._unitLabel.BackColor = Color.Transparent;
            this.Controls.Add(_unitLabel);
            this.TextAlign = HorizontalAlignment.Left;
            this.Width = 100;
            this.Height = 25;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.DecimalPlaces = 2;
            this.Increment = 0.01m;
            // 不在这里设置Minimum和Maximum，让调用者设置
            UpdateUnitLabelPosition();
        }

        private void UpdateUnitLabelPosition()
        {
            if (_unitLabel != null)
            {
                // 计算单位标签的位置，使其显示在输入框右端
                int labelWidth = (int)_unitLabel.CreateGraphics().MeasureString(_unitLabel.Text, _unitLabel.Font).Width;
                int rightMargin = 5; // 右边距
                _unitLabel.Location = new Point(this.Width - labelWidth - rightMargin, (this.Height - _unitLabel.Height) / 2);
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateUnitLabelPosition();
        }
    }

    public class LevelData
    {
        public int Level { get; set; }
        public string NumberStyle { get; set; }
        public string NumberFormat { get; set; }
        public decimal NumberIndent { get; set; }
        public decimal TextIndent { get; set; }
        public string AfterNumberType { get; set; } // 编号之后类型：无、空格、制表位
        public decimal TabPosition { get; set; } // 制表位位置
        public string LinkedStyle { get; set; }
    }

    public class LevelDataEventArgs : EventArgs
    {
        public LevelData LevelData { get; set; }
        
        public LevelDataEventArgs(LevelData levelData)
        {
            LevelData = levelData;
        }
    }


    public partial class MultiLevelList : Form
    {
        private int currentLevels = 5;
        private List<LevelData> levelDataList = new List<LevelData>();
        private Microsoft.Office.Interop.Word.Application app;
        
        // 编号样式数组，与参考代码保持一致
        private readonly WdListNumberStyle[] LevelNumStyle = new WdListNumberStyle[]
        {
            WdListNumberStyle.wdListNumberStyleArabic,           // 0: 1,2,3...
            WdListNumberStyle.wdListNumberStyleLegalLZ,          // 1: 01,02,03...
            WdListNumberStyle.wdListNumberStyleUppercaseLetter,  // 2: A,B,C...
            WdListNumberStyle.wdListNumberStyleLowercaseLetter,  // 3: a,b,c...
            WdListNumberStyle.wdListNumberStyleUppercaseRoman,   // 4: I,II,III...
            WdListNumberStyle.wdListNumberStyleLowercaseRoman,   // 5: i,ii,iii...
            WdListNumberStyle.wdListNumberStyleSimpChinNum1,     // 6: 一,二,三...
            WdListNumberStyle.wdListNumberStyleSimpChinNum2,     // 7: 壹,贰,叁...
            WdListNumberStyle.wdListNumberStyleZodiac1,          // 8: 甲,乙,丙...
            WdListNumberStyle.wdListNumberStyleArabic            // 9: 正规编号
        };

        public MultiLevelList()
        {
            InitializeComponent();
            app = Globals.ThisAddIn.Application;
            InitializeData();
            SetupEventHandlers();
            CreateLevelControls();
        }

        private void InitializeData()
        {
            // 初始化级别数据
            for (int i = 1; i <= 9; i++)
            {
                // 根据级别设置不同的初始值
                string numberStyle = "1,2,3...";
                string numberFormat = GenerateNumberFormat(i);
                decimal textIndent = 0m; // 默认文本缩进为0
                string afterNumberType = "空格";
                decimal tabPosition = 0m;
                string linkedStyle = "无";
                
                if (i == 1)
                {
                    numberFormat = "第%1章";
                    linkedStyle = "标题 1";
                }
                else if (i == 2)
                {
                    linkedStyle = "标题 2";
                }
                else if (i == 3)
                {
                    linkedStyle = "标题 3";
                }
                else if (i == 4)
                {
                    numberStyle = "1,2,3...";
                    numberFormat = "(%4)";
                    linkedStyle = "标题 4";
                }
                else if (i == 5)
                {
                    linkedStyle = "标题 5";
                }
                else if (i == 6)
                {
                    linkedStyle = "标题 6";
                }
                else if (i == 7)
                {
                    linkedStyle = "标题 7";
                }
                else if (i == 8)
                {
                    linkedStyle = "标题 8";
                }
                else if (i == 9)
                {
                    linkedStyle = "标题 9";
                }
                
                levelDataList.Add(new LevelData
                {
                    Level = i,
                    NumberStyle = numberStyle,
                    NumberFormat = numberFormat,
                    NumberIndent = 0.0m,
                    TextIndent = textIndent,
                    AfterNumberType = afterNumberType,
                    TabPosition = tabPosition,
                    LinkedStyle = linkedStyle
                });
            }

            // 设置默认显示4级
            cmbLevelCount.SelectedItem = "4";
            currentLevels = 4;
        }

        private void CreateLevelControls()
        {
            levelsContainer.Controls.Clear();

            // 动态创建级别控件 - 按正确顺序（标题在上，1级在下）
            for (int i = currentLevels; i >= 1; i--)
            {
                CreateLevelRow(i);
            }

            // 添加列标题 - 放在最后，这样会显示在最上方
            var headerPanel = new Panel();
            headerPanel.Height = 30;
            headerPanel.Dock = DockStyle.Top;
            headerPanel.BackColor = Color.Transparent;
            
            var lblLevel = new Label { Text = "级别", Location = new Point(10, 8), Size = new Size(50, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblNumberStyle = new Label { Text = "编号样式", Location = new Point(70, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblNumberFormat = new Label { Text = "编号格式", Location = new Point(180, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblNumberIndent = new Label { Text = "编号缩进", Location = new Point(290, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblTextIndent = new Label { Text = "文本缩进", Location = new Point(400, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblAfterNumber = new Label { Text = "编号之后", Location = new Point(510, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblTabPosition = new Label { Text = "制表位位置", Location = new Point(620, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            var lblLinkedStyle = new Label { Text = "链接样式", Location = new Point(730, 8), Size = new Size(100, 20), Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold) };
            
            headerPanel.Controls.AddRange(new Control[] { lblLevel, lblNumberStyle, lblNumberFormat, lblNumberIndent, lblTextIndent, lblAfterNumber, lblTabPosition, lblLinkedStyle });
            levelsContainer.Controls.Add(headerPanel);
            
            // 设置所有级别的制表位位置启用状态
            for (int i = 1; i <= currentLevels; i++)
            {
                UpdateTabPositionEnabled(i);
            }
        }

        private void CreateLevelRow(int level)
        {
            var rowPanel = new Panel();
            rowPanel.Height = 35;
            rowPanel.Dock = DockStyle.Top;
            rowPanel.BackColor = Color.Transparent;
            rowPanel.BorderStyle = BorderStyle.None;

            // 级别标签
            var lblLevel = new Label
            {
                Text = "第" + level + "级",
                Location = new Point(10, 8),
                Size = new Size(50, 20),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 123, 255)
            };

            // 编号样式下拉框
            var cmbNumberStyle = new ComboBox
            {
                Name = "CmbNumStyle" + level,
                Location = new Point(70, 5),
                Size = new Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F),
                BackColor = Color.FromArgb(245, 245, 245),
                Items = { "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,贰,叁...", "甲,乙,丙...", "正规编号" }
            };
            // 根据级别设置默认选择
            if (level == 4)
            {
                cmbNumberStyle.SelectedIndex = 0; // 1,2,3...
            }
            else
            {
                cmbNumberStyle.SelectedIndex = 0; // 1,2,3...
            }

            // 编号格式文本框
            var txtNumberFormat = new TextBox
            {
                Name = "TextBoxNumFormat" + level,
                Location = new Point(180, 5),
                Size = new Size(100, 25),
                Font = new Font("Microsoft YaHei", 9F),
                BackColor = Color.FromArgb(245, 245, 245)
            };
            // 根据级别设置默认格式
            if (level == 1)
            {
                txtNumberFormat.Text = "第%1章";
            }
            else if (level == 4)
            {
                txtNumberFormat.Text = "(%4)";
            }
            else
            {
                txtNumberFormat.Text = GenerateNumberFormat(level);
            }

            // 编号缩进
            var nudNumberIndent = new NumericUpDownWithUnit
            {
                Name = "TxtBoxNumIndent" + level,
                Location = new Point(290, 5),
                Size = new Size(100, 25)
            };
            nudNumberIndent.Minimum = 0;
            nudNumberIndent.Maximum = 50;
            nudNumberIndent.Value = 0;

            // 文本缩进
            var nudTextIndent = new NumericUpDownWithUnit
            {
                Name = "TxtBoxTextIndent" + level,
                Location = new Point(400, 5),
                Size = new Size(100, 25)
            };
            nudTextIndent.Minimum = 0;
            nudTextIndent.Maximum = 50;
            // 设置默认文本缩进为0
            nudTextIndent.Value = 0m;

            // 编号之后下拉框
            var cmbAfterNumber = new ComboBox
            {
                Name = "CmbAfterNumber" + level,
                Location = new Point(510, 5),
                Size = new Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F),
                BackColor = Color.FromArgb(245, 245, 245),
                Items = { "无", "空格", "制表位" }
            };
            cmbAfterNumber.SelectedIndex = 1; // 默认选择"空格"

            // 制表位位置
            var nudTabPosition = new NumericUpDownWithUnit
            {
                Name = "TxtBoxTabPosition" + level,
                Location = new Point(620, 5),
                Size = new Size(100, 25)
            };
            nudTabPosition.Minimum = 0;
            nudTabPosition.Maximum = 50;
            nudTabPosition.Value = 0m; // 初始值为0
            nudTabPosition.Enabled = false; // 初始状态禁用，因为默认选择"空格"

            // 链接样式下拉框
            var cmbLinkedStyle = new ComboBox
            {
                Name = "CmbLinkedStyle" + level,
                Location = new Point(730, 5),
                Size = new Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei", 9F),
                BackColor = Color.FromArgb(245, 245, 245),
                Items = { "无", "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9" }
            };
            // 根据级别设置默认链接样式
            cmbLinkedStyle.SelectedIndex = level; // 0=无, 1=标题 1, 2=标题 2, ...

            // 添加控件到行面板
            rowPanel.Controls.AddRange(new Control[] { 
                lblLevel, cmbNumberStyle, txtNumberFormat, 
                nudNumberIndent, nudTextIndent, cmbAfterNumber, nudTabPosition, cmbLinkedStyle 
            });

            // 添加事件处理
            cmbNumberStyle.SelectedIndexChanged += (s, e) => UpdateLevelData(level);
            txtNumberFormat.TextChanged += (s, e) => UpdateLevelData(level);
            nudNumberIndent.ValueChanged += (s, e) => UpdateLevelData(level);
            nudTextIndent.ValueChanged += (s, e) => UpdateLevelData(level);
            cmbAfterNumber.SelectedIndexChanged += (s, e) => {
                UpdateLevelData(level);
                UpdateTabPositionEnabled(level);
            };
            nudTabPosition.ValueChanged += (s, e) => UpdateLevelData(level);
            cmbLinkedStyle.SelectedIndexChanged += (s, e) => UpdateLevelData(level);

            levelsContainer.Controls.Add(rowPanel);
        }

        private void UpdateLevelData(int level)
        {
            if (level < 1 || level > levelDataList.Count) return;

            var levelData = levelDataList[level - 1];
            var cmbNumberStyle = levelsContainer.Controls.Find("CmbNumStyle" + level, true).FirstOrDefault() as ComboBox;
            var txtNumberFormat = levelsContainer.Controls.Find("TextBoxNumFormat" + level, true).FirstOrDefault() as TextBox;
            var nudNumberIndent = levelsContainer.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var nudTextIndent = levelsContainer.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var cmbAfterNumber = levelsContainer.Controls.Find("CmbAfterNumber" + level, true).FirstOrDefault() as ComboBox;
            var nudTabPosition = levelsContainer.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var cmbLinkedStyle = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as ComboBox;

            if (cmbNumberStyle != null) levelData.NumberStyle = cmbNumberStyle.Text;
            if (txtNumberFormat != null) levelData.NumberFormat = txtNumberFormat.Text;
            if (nudNumberIndent != null) levelData.NumberIndent = nudNumberIndent.Value;
            if (nudTextIndent != null) levelData.TextIndent = nudTextIndent.Value;
            if (cmbAfterNumber != null) levelData.AfterNumberType = cmbAfterNumber.Text;
            if (nudTabPosition != null) levelData.TabPosition = nudTabPosition.Value;
            if (cmbLinkedStyle != null) levelData.LinkedStyle = cmbLinkedStyle.Text;
        }

        private void UpdateTabPositionEnabled(int level)
        {
            var cmbAfterNumber = levelsContainer.Controls.Find("CmbAfterNumber" + level, true).FirstOrDefault() as ComboBox;
            var nudTabPosition = levelsContainer.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            
            if (cmbAfterNumber != null && nudTabPosition != null)
            {
                // 只有当"编号之后"选择"制表位"时，制表位位置输入框才启用
                nudTabPosition.Enabled = (cmbAfterNumber.Text == "制表位");
            }
        }


        private string GenerateNumberFormat(int level)
        {
            StringBuilder format = new StringBuilder();
            for (int i = 1; i <= level; i++)
            {
                format.Append("%" + i);
                if (i < level)
                    format.Append(".");
            }
            return format.ToString();
        }

        private void SetupEventHandlers()
        {
            // 底部控制按钮事件
            cmbLevelCount.SelectedIndexChanged += CmbLevelCount_SelectedIndexChanged;
            btnSetLevelStyle.Click += BtnSetLevelStyle_Click;
            btnLoadCurrentList.Click += BtnLoadCurrentList_Click;
            btnSetMultiLevelList.Click += BtnApply_Click;
            btnClose.Click += btnClose_Click;
            btnApplySettings.Click += BtnApplySettings_Click;

            // 右侧快捷设置事件
            checkBox1.CheckedChanged += CheckBox_CheckedChanged; // 编号缩进
            checkBox2.CheckedChanged += CheckBox_CheckedChanged; // 文本缩进
            checkBox3.CheckedChanged += CheckBox_CheckedChanged; // 制表位位置
            checkBox4.CheckedChanged += CheckBox4_CheckedChanged; // 递进缩进设置
            checkBox5.CheckedChanged += CheckBox5_CheckedChanged; // 链接到标题样式
            checkBox6.CheckedChanged += CheckBox6_CheckedChanged; // 不链接标题样式
        }

        private void CmbLevelCount_SelectedIndexChanged(object sender, EventArgs e)
        {
            currentLevels = int.Parse(cmbLevelCount.SelectedItem.ToString());
            CreateLevelControls();
        }

        private void BtnSetLevelStyle_Click(object sender, EventArgs e)
        {
            // 调用样式设置窗格
            StyleSettings styleForm = new StyleSettings();
            styleForm.ShowDialog();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnLoadCurrentList_Click(object sender, EventArgs e)
        {
            try
            {
                var selection = app.Selection;
                var listTemplate = selection.Range.ListFormat.ListTemplate;

                if (listTemplate == null || listTemplate.ListLevels == null)
                {
                    MessageBox.Show("当前位置无多级列表！", "提醒");
                    return;
                }

                int maxLevel = 1;
                foreach (ListLevel listLevel in listTemplate.ListLevels)
                {
                    if (listLevel.NumberFormat == "")
                        break;
                    maxLevel = listLevel.Index;
                }

                cmbLevelCount.SelectedItem = maxLevel.ToString();
                currentLevels = maxLevel;
                CreateLevelControls();

                // 载入数据到控件
                for (int i = 1; i <= maxLevel; i++)
                {
                    ListLevel listLevel = listTemplate.ListLevels[i];
                    
                    var cmbNumberStyle = levelsContainer.Controls.Find("CmbNumStyle" + i, true).FirstOrDefault() as ComboBox;
                    var txtNumberFormat = levelsContainer.Controls.Find("TextBoxNumFormat" + i, true).FirstOrDefault() as TextBox;
                    var nudNumberIndent = levelsContainer.Controls.Find("TxtBoxNumIndent" + i, true).FirstOrDefault() as NumericUpDownWithUnit;
                    var nudTextIndent = levelsContainer.Controls.Find("TxtBoxTextIndent" + i, true).FirstOrDefault() as NumericUpDownWithUnit;
                    var nudAfterIndent = levelsContainer.Controls.Find("TxtBoxAfterNumIndent" + i, true).FirstOrDefault() as NumericUpDownWithUnit;
                    var cmbLinkedStyle = levelsContainer.Controls.Find("CmbLinkedStyle" + i, true).FirstOrDefault() as ComboBox;

                    if (cmbNumberStyle != null)
                    {
                        int styleIndex = GetNumberStyleIndex(listLevel.NumberStyle);
                        cmbNumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
                    }
                    
                    if (txtNumberFormat != null)
                        txtNumberFormat.Text = listLevel.NumberFormat.ToString();
                    
                    if (nudNumberIndent != null)
                        nudNumberIndent.Value = (decimal)app.PointsToCentimeters(listLevel.NumberPosition);
                    
                    if (nudTextIndent != null)
                        nudTextIndent.Value = (decimal)app.PointsToCentimeters(listLevel.TextPosition);
                    
                    if (nudAfterIndent != null)
                    {
                        if (listLevel.TabPosition != 9999999f)
                            nudAfterIndent.Value = (decimal)app.PointsToCentimeters(listLevel.TabPosition);
                        else
                            nudAfterIndent.Value = 0;
                    }
                    
                    if (cmbLinkedStyle != null)
                        cmbLinkedStyle.Text = string.IsNullOrEmpty(listLevel.LinkedStyle) ? "无" : listLevel.LinkedStyle;
                }

                // 清空快捷设置
                ClearQuickSettings();
                MessageBox.Show("已载入当前多级列表设置", "成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show("载入失败：" + ex.Message, "错误");
            }
        }

        private int GetNumberStyleIndex(WdListNumberStyle numberStyle)
        {
            var styles = new[] { 
                WdListNumberStyle.wdListNumberStyleArabic,
                WdListNumberStyle.wdListNumberStyleArabic,
                WdListNumberStyle.wdListNumberStyleUppercaseLetter,
                WdListNumberStyle.wdListNumberStyleLowercaseLetter,
                WdListNumberStyle.wdListNumberStyleUppercaseRoman,
                WdListNumberStyle.wdListNumberStyleLowercaseRoman,
                WdListNumberStyle.wdListNumberStyleCardinalText,
                WdListNumberStyle.wdListNumberStyleOrdinalText,
                WdListNumberStyle.wdListNumberStyleOrdinal,
                WdListNumberStyle.wdListNumberStyleOrdinal
            };
            
            for (int i = 0; i < styles.Length; i++)
            {
                if (styles[i] == numberStyle)
                    return i;
            }
            return -1;
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            try
            {
                int levelCount = currentLevels;
                int[] numberStyles = new int[levelCount];
                string[] numberFormats = new string[levelCount];
                string[] linkedStyles = new string[levelCount];
                float[] numberIndents = new float[levelCount];
                float[] textIndents = new float[levelCount];
                string[] afterNumberTypes = new string[levelCount];
                float[] tabPositions = new float[levelCount];

                // 收集数据
                for (int i = 0; i < levelCount; i++)
                {
                    var numberStyleCombo = levelsContainer.Controls.Find("CmbNumStyle" + (i + 1), true).FirstOrDefault() as ComboBox;
                    var numberFormatText = levelsContainer.Controls.Find("TextBoxNumFormat" + (i + 1), true).FirstOrDefault() as TextBox;
                    var numberIndentControl = levelsContainer.Controls.Find("TxtBoxNumIndent" + (i + 1), true).FirstOrDefault() as NumericUpDownWithUnit;
                    var textIndentControl = levelsContainer.Controls.Find("TxtBoxTextIndent" + (i + 1), true).FirstOrDefault() as NumericUpDownWithUnit;
                    var afterNumberCombo = levelsContainer.Controls.Find("CmbAfterNumber" + (i + 1), true).FirstOrDefault() as ComboBox;
                    var tabPositionControl = levelsContainer.Controls.Find("TxtBoxTabPosition" + (i + 1), true).FirstOrDefault() as NumericUpDownWithUnit;
                    var linkedStyleCombo = levelsContainer.Controls.Find("CmbLinkedStyle" + (i + 1), true).FirstOrDefault() as ComboBox;

                    if (numberStyleCombo != null)
                        numberStyles[i] = numberStyleCombo.SelectedIndex;
                    
                    if (numberFormatText != null)
                    {
                        if (!numberFormatText.Text.Contains("%" + (i + 1)))
                        {
                            MessageBox.Show("错误：第" + (i + 1) + "级编号格式未包含本级别的编号！");
                            return;
                        }
                        numberFormats[i] = numberFormatText.Text;
                    }
                    
                    if (linkedStyleCombo != null)
                    {
                        if (i == 0)
                        {
                            linkedStyles[i] = linkedStyleCombo.Text;
                        }
                        else
                        {
                            if (linkedStyles.Contains(linkedStyleCombo.Text) && linkedStyleCombo.Text != "无")
                            {
                                MessageBox.Show("错误：第" + (i + 1) + "级链接样式出现重复！");
                                return;
                            }
                            linkedStyles[i] = linkedStyleCombo.Text;
                        }
                    }
                    
                    if (numberIndentControl != null)
                    {
                        try
                        {
                            numberIndents[i] = (float)numberIndentControl.Value;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"第{i + 1}级编号缩进值错误：{ex.Message}", "错误");
                            return;
                        }
                    }
                    
                    if (textIndentControl != null)
                    {
                        try
                        {
                            textIndents[i] = (float)textIndentControl.Value;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"第{i + 1}级文本缩进值错误：{ex.Message}", "错误");
                            return;
                        }
                    }
                    
                    if (afterNumberCombo != null)
                    {
                        afterNumberTypes[i] = afterNumberCombo.Text;
                    }
                    
                    if (tabPositionControl != null)
                    {
                        try
                        {
                            tabPositions[i] = (float)tabPositionControl.Value;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"第{i + 1}级制表位位置值错误：{ex.Message}", "错误");
                            return;
                        }
                    }
                }

                // 创建多级列表模板
                CreateListTemplate(levelCount, numberStyles, numberFormats, numberIndents, textIndents, afterNumberTypes, tabPositions, linkedStyles);
            }
            catch (Exception ex)
            {
                MessageBox.Show("设置失败：" + ex.Message + "\n\n详细错误：" + ex.StackTrace, "错误");
            }
        }

        private void CreateListTemplate(int levelCount, int[] numberStyles, string[] numberFormats, 
            float[] numberIndents, float[] textIndents, string[] afterNumberTypes, float[] tabPositions, string[] linkedStyles)
        {
            // 验证参数
            if (levelCount <= 0 || levelCount > 9)
            {
                throw new ArgumentException($"级别数量无效: {levelCount}");
            }
            
            ListTemplate listTemplate;
            object Index;
            
            // 检查当前选区是否已有列表模板
            if (app.Selection.Range.ListFormat.ListTemplate != null)
            {
                listTemplate = app.Selection.Range.ListFormat.ListTemplate;
            }
            else
            {
                // 从ListGalleries获取模板
                ListTemplates listTemplates = app.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates;
                Index = 7;
                listTemplate = listTemplates[ref Index];
            }
            
            ListTemplate listTemplate2 = listTemplate;
            
            // 设置各级属性
            for (int i = 1; i <= levelCount; i++)
            {
                ListLevel listLevel = listTemplate2.ListLevels[i];
                
                if (numberStyles[i - 1] != -1)
                {
                    listLevel.NumberStyle = LevelNumStyle[numberStyles[i - 1]];
                }
                
                listLevel.NumberFormat = numberFormats[i - 1];
                
                // 设置链接样式
                if (linkedStyles[i - 1] != "无" && !string.IsNullOrEmpty(linkedStyles[i - 1]))
                {
                    // 将显示文本转换为Word样式名称
                    string styleName = GetWordStyleName(linkedStyles[i - 1]);
                    listLevel.LinkedStyle = styleName;
                    
                    // 添加调试信息到消息框
                    string debugInfo = $"级别 {i}: 设置链接样式为 '{styleName}'\n";
                    System.Diagnostics.Debug.WriteLine(debugInfo);
                }
                else
                {
                    listLevel.LinkedStyle = "";
                }
                
                listLevel.NumberPosition = app.CentimetersToPoints(numberIndents[i - 1]);
                listLevel.TextPosition = app.CentimetersToPoints(textIndents[i - 1]);
                
                // 设置编号之后的字符类型和制表位位置
                if (afterNumberTypes[i - 1] == "制表位" && tabPositions[i - 1] > 0f)
                {
                    listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
                    listLevel.TabPosition = app.CentimetersToPoints(tabPositions[i - 1]);
                }
                else if (afterNumberTypes[i - 1] == "空格")
                {
                    listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingSpace;
                }
                else
                {
                    listLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
                }
                
                listLevel.StartAt = 1;
                listLevel.ResetOnHigher = i - 1;
            }
            
            // 清空未使用的级别
            for (int j = levelCount + 1; j <= 9; j++)
            {
                ListLevel listLevel2 = listTemplate2.ListLevels[j];
                listLevel2.NumberFormat = "";
                listLevel2.NumberStyle = WdListNumberStyle.wdListNumberStyleNone;
            }
            
            // 先清除现有的列表格式
            app.Selection.Range.ListFormat.RemoveNumbers();
            
            // 应用多级列表到当前选区
            ListFormat listFormat = app.Selection.Range.ListFormat;
            Index = false;
            object ApplyTo = WdListApplyTo.wdListApplyToWholeList;
            object DefaultListBehavior = WdDefaultListBehavior.wdWord9ListBehavior;
            object ApplyLevel2 = levelCount;
            listFormat.ApplyListTemplateWithLevel(listTemplate2, ref Index, ref ApplyTo, ref DefaultListBehavior, ref ApplyLevel2);
        }

        private WdListNumberStyle GetNumberStyleByIndex(int index)
        {
            var styles = new[] { 
                WdListNumberStyle.wdListNumberStyleArabic,
                WdListNumberStyle.wdListNumberStyleArabic,
                WdListNumberStyle.wdListNumberStyleUppercaseLetter,
                WdListNumberStyle.wdListNumberStyleLowercaseLetter,
                WdListNumberStyle.wdListNumberStyleUppercaseRoman,
                WdListNumberStyle.wdListNumberStyleLowercaseRoman,
                WdListNumberStyle.wdListNumberStyleCardinalText,
                WdListNumberStyle.wdListNumberStyleOrdinalText,
                WdListNumberStyle.wdListNumberStyleOrdinal,
                WdListNumberStyle.wdListNumberStyleOrdinal
            };
            
            if (index < 0 || index >= styles.Length)
            {
                throw new ArgumentException($"编号样式索引无效: {index}, 有效范围: 0-{styles.Length - 1}");
            }
            
            return styles[index];
        }

        private void BtnApplySettings_Click(object sender, EventArgs e)
        {
            // 应用快捷设置
            ApplyQuickSettings();
        }

        private void ApplyQuickSettings()
        {
            for (int level = 1; level <= currentLevels; level++)
            {
                // 1. 统一缩进设置
                if (checkBox1.Checked) // 编号缩进
                {
                    var numberIndentControl = levelsContainer.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (numberIndentControl != null)
                        numberIndentControl.Value = numericUpDown1.Value;
                }
                
                if (checkBox2.Checked) // 文本缩进
                {
                    var textIndentControl = levelsContainer.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (textIndentControl != null)
                        textIndentControl.Value = numericUpDown4.Value; // 使用numericUpDown4（文本缩进输入框）
                }
                
                if (checkBox3.Checked) // 制表位位置
                {
                    var tabPositionControl = levelsContainer.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (tabPositionControl != null)
                        tabPositionControl.Value = numericUpDown5.Value; // 使用numericUpDown5（制表位位置输入框）
                }

                // 2. 递进缩进设置
                if (checkBox4.Checked) // 递进缩进设置
                {
                    var numberIndentControl = levelsContainer.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (numberIndentControl != null)
                    {
                        if (level == 1)
                        {
                            numberIndentControl.Value = numericUpDown2.Value; // 一级编号缩进
                        }
                        else
                        {
                            numberIndentControl.Value = numericUpDown2.Value + numericUpDown3.Value * (level - 1); // 递进计算
                        }
                    }
                }

                // 3. 链接标题样式
                if (checkBox5.Checked) // 链接到标题样式
                {
                    var linkedStyleControl = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as ComboBox;
                    if (linkedStyleControl != null)
                        linkedStyleControl.SelectedIndex = level;
                }
                else if (checkBox6.Checked) // 不链接标题样式
                {
                    var linkedStyleControl = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as ComboBox;
                    if (linkedStyleControl != null)
                        linkedStyleControl.SelectedIndex = 0;
                }
            }

            // 清空快捷设置
            ClearQuickSettings();
        }

        private void ClearQuickSettings()
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            numericUpDown1.Enabled = false;
            numericUpDown2.Enabled = false;
            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;
            numericUpDown5.Enabled = false;
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // 编号缩进使用numericUpDown1
            if (checkBox1.Checked) // 编号缩进
            {
                numericUpDown1.Enabled = true;
            }
            else if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked)
            {
                numericUpDown1.Enabled = false;
            }
            
            // 文本缩进使用numericUpDown4
            if (checkBox2.Checked) // 文本缩进
            {
                numericUpDown4.Enabled = true;
            }
            else if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked)
            {
                numericUpDown4.Enabled = false;
            }
            
            // 制表位位置使用numericUpDown5
            if (checkBox3.Checked) // 制表位位置
            {
                numericUpDown5.Enabled = true;
            }
            else if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked)
            {
                numericUpDown5.Enabled = false;
            }
        }

        private void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            // 递进缩进设置
            if (checkBox4.Checked)
            {
                numericUpDown2.Enabled = true;
                numericUpDown3.Enabled = true;
            }
            else
            {
                numericUpDown2.Enabled = false;
                numericUpDown3.Enabled = false;
            }
        }

        private void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                checkBox6.Checked = false;
            }
        }

        private void CheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                checkBox5.Checked = false;
            }
        }

        private string GetNumberStyleName(WdListNumberStyle numberStyle)
        {
            switch (numberStyle)
            {
                case WdListNumberStyle.wdListNumberStyleArabic: return "1,2,3...";
                case WdListNumberStyle.wdListNumberStyleUppercaseLetter: return "A,B,C...";
                case WdListNumberStyle.wdListNumberStyleLowercaseLetter: return "a,b,c...";
                case WdListNumberStyle.wdListNumberStyleUppercaseRoman: return "I,II,III...";
                case WdListNumberStyle.wdListNumberStyleLowercaseRoman: return "i,ii,iii...";
                case WdListNumberStyle.wdListNumberStyleCardinalText: return "一,二,三...";
                case WdListNumberStyle.wdListNumberStyleOrdinalText: return "壹,贰,叁...";
                case WdListNumberStyle.wdListNumberStyleOrdinal: return "①,②,③...";
                default: return "1,2,3...";
            }
        }

        private WdListNumberStyle GetNumberStyle(string styleName)
        {
            switch (styleName)
            {
                case "1,2,3...": return WdListNumberStyle.wdListNumberStyleArabic;
                case "01,02,03...": return WdListNumberStyle.wdListNumberStyleArabic;
                case "A,B,C...": return WdListNumberStyle.wdListNumberStyleUppercaseLetter;
                case "a,b,c...": return WdListNumberStyle.wdListNumberStyleLowercaseLetter;
                case "I,II,III...": return WdListNumberStyle.wdListNumberStyleUppercaseRoman;
                case "i,ii,iii...": return WdListNumberStyle.wdListNumberStyleLowercaseRoman;
                case "一,二,三...": return WdListNumberStyle.wdListNumberStyleCardinalText;
                case "壹,贰,叁...": return WdListNumberStyle.wdListNumberStyleOrdinalText;
                case "①,②,③...": return WdListNumberStyle.wdListNumberStyleOrdinal;
                case "⑴,⑵,⑶...": return WdListNumberStyle.wdListNumberStyleOrdinal;
                case "1),2),3)...": return WdListNumberStyle.wdListNumberStyleArabic;
                case "(1),(2),(3)...": return WdListNumberStyle.wdListNumberStyleArabic;
                default: return WdListNumberStyle.wdListNumberStyleArabic;
            }
        }

        private string GetWordStyleName(string displayName)
        {
            // 如果已经是"无"，直接返回空字符串
            if (displayName == "无")
            {
                return "";
            }

            // 提取级别数字
            int level = 0;
            if (!int.TryParse(displayName.Replace("标题 ", "").Replace("标题", ""), out level) || level < 1 || level > 9)
            {
                System.Diagnostics.Debug.WriteLine($"无法解析级别数字: '{displayName}'");
                return "";
            }

            // 首先尝试通过内置样式索引获取（最可靠的方法）
            try
            {
                var builtInStyle = app.ActiveDocument.Styles[WdBuiltinStyle.wdStyleHeading1 + level - 1];
                if (builtInStyle != null)
                {
                    System.Diagnostics.Debug.WriteLine($"找到内置样式: '{builtInStyle.NameLocal}'");
                    return builtInStyle.NameLocal;
                }
            }
            catch
            {
                // 内置样式不存在，继续尝试其他方法
            }

            // 尝试多种可能的样式名称格式
            string[] possibleNames = {
                // 中文格式
                displayName,                           // 标题 1, 标题 2, ...
                displayName.Replace("标题 ", "标题"),    // 标题1, 标题2, ...
                
                // 英文格式
                displayName.Replace("标题 ", "Heading "), // Heading 1, Heading 2, ...
                displayName.Replace("标题 ", "Heading"),  // Heading1, Heading2, ...
                
                // 其他可能的格式
                "Heading " + level,                    // Heading 1, Heading 2, ...
                "标题 " + level,                       // 标题 1, 标题 2, ...
                "Heading" + level,                     // Heading1, Heading2, ...
                "标题" + level,                        // 标题1, 标题2, ...
            };

            // 检查样式是否存在
            foreach (string styleName in possibleNames)
            {
                try
                {
                    var style = app.ActiveDocument.Styles[styleName];
                    if (style != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"找到样式: '{styleName}'");
                        return styleName;
                    }
                }
                catch
                {
                    // 样式不存在，继续尝试下一个
                }
            }

            // 如果都找不到，返回空字符串（表示不链接样式）
            System.Diagnostics.Debug.WriteLine($"未找到样式，返回空字符串: '{displayName}'");
            return "";
        }
    }
}