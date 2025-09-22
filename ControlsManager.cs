using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Point = System.Drawing.Point;
using Font = System.Drawing.Font;
using Color = System.Drawing.Color;

namespace WordMan_VSTO
{
    #region 数据结构定义

    /// <summary>
    /// 输入框值结构体
    /// </summary>
    public struct InputValues
    {
        public decimal NumberIndent { get; set; }
        public decimal TextIndent { get; set; }
        public decimal TabPosition { get; set; }
    }

    /// <summary>
    /// 单位转换器 - 使用磅作为中间单位进行转换
    /// </summary>
    public static class UnitConverter
    {
        #region 转换常量
        /// <summary>
        /// 厘米到磅的转换系数（1厘米 = 28.35磅）
        /// </summary>
        private static readonly double CentimetersToPointsRatio = 28.35;
        
        /// <summary>
        /// 磅到厘米的转换系数
        /// </summary>
        private static readonly double PointsToCentimetersRatio = 1.0 / CentimetersToPointsRatio;
        
        /// <summary>
        /// 字符到厘米的转换系数（1字符 ≈ 0.5厘米）
        /// </summary>
        private static readonly double CharactersToCentimetersRatio = 0.5;
        
        /// <summary>
        /// 行到磅的转换系数（1行 ≈ 12磅）
        /// </summary>
        private static readonly double LinesToPointsRatio = 12.0;
        #endregion

        /// <summary>
        /// 转换数值单位
        /// </summary>
        /// <param name="value">要转换的数值</param>
        /// <param name="sourceUnit">源单位</param>
        /// <param name="targetUnit">目标单位</param>
        /// <returns>转换后的数值</returns>
        public static double UnitConvert(double value, string sourceUnit, string targetUnit)
        {
            // 单位相同直接返回
            if (string.Equals(sourceUnit, targetUnit, StringComparison.OrdinalIgnoreCase))
                return value;

            var wordApp = Globals.ThisAddIn?.Application;
            
            // 如果Word Application不可用，使用默认转换
            if (wordApp == null)
            {
                return ConvertWithoutWordApp(value, sourceUnit, targetUnit);
            }
            
            // 第一步：将所有单位转换为中间单位（磅）
            double valueInPoints = ConvertToPoints(value, sourceUnit, wordApp);
            
            // 第二步：将中间单位（磅）转换为目标单位
            return ConvertFromPoints(valueInPoints, targetUnit, wordApp);
        }

        /// <summary>
        /// 不使用Word Application的简单单位转换（用于设计时）
        /// </summary>
        /// <param name="value">要转换的数值</param>
        /// <param name="sourceUnit">源单位</param>
        /// <param name="targetUnit">目标单位</param>
        /// <returns>转换后的数值</returns>
        private static double ConvertWithoutWordApp(double value, string sourceUnit, string targetUnit)
        {
            // 简化的转换逻辑，用于设计时
            
            // 转换为磅
            double valueInPoints;
            switch (sourceUnit)
            {
                case "磅":
                    valueInPoints = value;
                    break;
                case "厘米":
                    valueInPoints = value * CentimetersToPointsRatio;
                    break;
                case "字符":
                    valueInPoints = value * CentimetersToPointsRatio * CharactersToCentimetersRatio; // 1字符 ≈ 0.5厘米
                    break;
                case "行":
                    valueInPoints = value * LinesToPointsRatio; // 1行 ≈ 12磅
                    break;
                default:
                    valueInPoints = value;
                    break;
            }
            
            // 从磅转换
            switch (targetUnit)
            {
                case "磅":
                    return valueInPoints;
                case "厘米":
                    return valueInPoints * PointsToCentimetersRatio;
                case "字符":
                    return valueInPoints * PointsToCentimetersRatio * 2; // 1厘米 ≈ 2字符
                case "行":
                    return valueInPoints / LinesToPointsRatio; // 1磅 ≈ 1/12行
                default:
                    return valueInPoints;
            }
        }

        /// <summary>
        /// 将各种单位转换为磅
        /// </summary>
        /// <param name="value">要转换的数值</param>
        /// <param name="unit">源单位</param>
        /// <param name="wordApp">Word应用程序实例</param>
        /// <returns>转换后的磅值</returns>
        private static double ConvertToPoints(double value, string unit, Microsoft.Office.Interop.Word.Application wordApp)
        {
            switch (unit)
            {
                case "磅":
                    return value;
                case "厘米":
                    return wordApp.CentimetersToPoints((float)value);
                case "字符":
                    // 字符转换：1字符 ≈ 0.5厘米，1厘米 ≈ 28.35磅
                    return wordApp.CentimetersToPoints((float)(value * CharactersToCentimetersRatio));
                case "行":
                    // 行转换：1行 ≈ 12磅（标准行距）
                    return value * LinesToPointsRatio;
                default:
                    return value; // 默认返回原值
            }
        }

        /// <summary>
        /// 将磅转换为各种单位
        /// </summary>
        /// <param name="valueInPoints">磅值</param>
        /// <param name="unit">目标单位</param>
        /// <param name="wordApp">Word应用程序实例</param>
        /// <returns>转换后的数值</returns>
        private static double ConvertFromPoints(double valueInPoints, string unit, Microsoft.Office.Interop.Word.Application wordApp)
        {
            switch (unit)
            {
                case "磅":
                    return valueInPoints;
                case "厘米":
                    return wordApp.PointsToCentimeters((float)valueInPoints);
                case "字符":
                    // 字符转换：1磅 ≈ 0.035厘米，1厘米 ≈ 2字符
                    return wordApp.PointsToCentimeters((float)valueInPoints) * 2;
                case "行":
                    // 行转换：1磅 ≈ 1/12行
                    return valueInPoints / LinesToPointsRatio;
                default:
                    return valueInPoints; // 默认返回原值
            }
        }
    }


    #endregion


    #region 标准控件
    /// <summary>
    /// 标准样式按钮 - 简化版本
    /// </summary>
    [System.ComponentModel.DesignerCategory("")]
    public class StandardButton : Button
    {
        public enum ButtonType
        {
            Primary,    // 主要按钮（蓝色）
            Secondary,  // 次要按钮（黑色）
            Small       // 小按钮（导入导出）
        }

        /// <summary>
        /// 初始化标准按钮
        /// </summary>
        public StandardButton() : this(ButtonType.Secondary, "", null, null)
        {
        }

        /// <summary>
        /// 初始化标准按钮
        /// </summary>
        /// <param name="type">按钮类型</param>
        /// <param name="text">按钮文本</param>
        /// <param name="size">按钮大小</param>
        /// <param name="location">按钮位置</param>
        public StandardButton(ButtonType type = ButtonType.Secondary, string text = "", Size? size = null, Point? location = null)
        {
            // 基础样式设置
            this.FlatStyle = FlatStyle.Flat;
            this.Font = new Font("Microsoft YaHei", 10F, FontStyle.Bold);
            this.UseVisualStyleBackColor = false;
            this.Text = text;
            
            // 设置大小和位置
            if (size.HasValue)
                this.Size = size.Value;
            if (location.HasValue)
                this.Location = location.Value;

            // 根据按钮类型设置样式
            SetButtonStyle(type);
        }

        /// <summary>
        /// 设置按钮样式
        /// </summary>
        /// <param name="type">按钮类型</param>
        private void SetButtonStyle(ButtonType type)
        {
            // 通用样式
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.FlatAppearance.BorderSize = 1;

            switch (type)
            {
                case ButtonType.Primary:
                    this.FlatAppearance.BorderColor = Color.FromArgb(0, 123, 255);
                    this.ForeColor = Color.FromArgb(0, 123, 255);
                    break;
                case ButtonType.Secondary:
                    this.FlatAppearance.BorderColor = Color.FromArgb(10, 10, 10);
                    this.ForeColor = Color.Black;
                    break;
                case ButtonType.Small:
                    this.FlatAppearance.BorderColor = Color.FromArgb(10, 10, 10);
                    this.ForeColor = Color.Black;
                    this.Size = new Size(50, 35);
                    break;
            }
        }
    }

    /// <summary>
    /// 标准数值输入框 - 支持字符、厘米、磅单位转换
    /// </summary>
    [System.ComponentModel.DesignerCategory("")]
    public class StandardNumericUpDown : NumericUpDown, System.ComponentModel.ISupportInitialize
    {
        private Label _unitLabel;
        private string _currentUnit = "厘米";
        private readonly string[] _availableUnits = { "厘米", "磅", "字符", "行" };

        public string Unit
        {
            get => _currentUnit;
            set
            {
                if (!string.Equals(_currentUnit, value, StringComparison.OrdinalIgnoreCase))
                {
                    _currentUnit = value;
                    if (_unitLabel != null)
                    {
                        _unitLabel.Text = _currentUnit;
                        UpdatePosition();
                    }
                }
            }
        }

        /// <summary>
        /// 获取指定单位的数值
        /// </summary>
        public decimal GetValueInUnit(string targetUnit)
        {
            if (string.IsNullOrEmpty(targetUnit) || string.Equals(targetUnit, _currentUnit, StringComparison.OrdinalIgnoreCase))
                return Value;
            
            return (decimal)UnitConverter.UnitConvert((double)Value, _currentUnit, targetUnit);
        }

        /// <summary>
        /// 设置指定单位的数值
        /// </summary>
        public void SetValueInUnit(decimal value, string sourceUnit)
        {
            if (string.IsNullOrEmpty(sourceUnit) || string.Equals(sourceUnit, _currentUnit, StringComparison.OrdinalIgnoreCase))
            {
                Value = value;
                return;
            }
            
            Value = (decimal)UnitConverter.UnitConvert((double)value, sourceUnit, _currentUnit);
        }

        /// <summary>
        /// 获取厘米单位的数值（便捷方法）
        /// </summary>
        public decimal GetValueInCentimeters()
        {
            return GetValueInUnit("厘米");
        }

        /// <summary>
        /// 设置厘米单位的数值（便捷方法）
        /// </summary>
        public void SetValueInCentimeters(decimal value)
        {
            SetValueInUnit(value, "厘米");
        }

        /// <summary>
        /// 初始化标准数值输入框
        /// </summary>
        public StandardNumericUpDown()
        {
            _currentUnit = "厘米";
            InitializeComponent();
        }

        /// <summary>
        /// 初始化标准数值输入框
        /// </summary>
        /// <param name="wordApp">Word应用程序实例</param>
        /// <param name="unit">默认单位</param>
        public StandardNumericUpDown(Microsoft.Office.Interop.Word.Application wordApp = null, string unit = "厘米")
        {
            _currentUnit = unit;
            InitializeComponent();
        }

        /// <summary>
        /// 初始化组件
        /// </summary>
        private void InitializeComponent()
        {
            // 设置输入框样式
            this.TextAlign = HorizontalAlignment.Left;
            this.Width = 100;
            this.Height = 25;
            this.BackColor = Color.FromArgb(250, 250, 250);
            this.DecimalPlaces = 1;
            this.Increment = 0.1m;
            this.Minimum = decimal.MinValue; // 取消下限
            this.Maximum = decimal.MaxValue; // 取消上限
            this.BorderStyle = BorderStyle.FixedSingle;
            
            // 创建单位标签
            _unitLabel = new Label
            {
                Text = _currentUnit,
                AutoSize = true,
                Font = new Font("Microsoft YaHei", 8F, FontStyle.Bold),
                ForeColor = Color.FromArgb(10, 10, 10), 
                BackColor = Color.Transparent,
                BorderStyle = BorderStyle.None,
                Cursor = Cursors.Hand // 鼠标悬停时显示手型光标
            };
            
            // 添加点击事件
            _unitLabel.Click += UnitLabel_Click;
            _unitLabel.MouseEnter += UnitLabel_MouseEnter;
            _unitLabel.MouseLeave += UnitLabel_MouseLeave;
            
            this.Controls.Add(_unitLabel);
            
            // 立即更新标签位置
            UpdatePosition();
        }

        private void UpdatePosition()
        {
            if (_unitLabel != null && this.Width > 0 && this.Height > 0)
            {
                using (Graphics graphics = this.CreateGraphics())
                {
                    SizeF unitLabelTextSize = graphics.MeasureString(_unitLabel.Text, _unitLabel.Font);
                    int rightMargin = 20; 
                    int labelX = this.Width - (int)unitLabelTextSize.Width - rightMargin;
                    int labelY = (this.Height - (int)unitLabelTextSize.Height) / 2;
                    
                    // 确保标签在控件范围内
                    labelX = Math.Max(0, Math.Min(labelX, this.Width - (int)unitLabelTextSize.Width));
                    labelY = Math.Max(0, Math.Min(labelY, this.Height - (int)unitLabelTextSize.Height));
                    
                    _unitLabel.Location = new Point(labelX, labelY);
                    _unitLabel.BringToFront(); // 确保标签在最前面
                }
            }
        }

        /// <summary>
        /// 处理控件大小改变事件
        /// </summary>
        /// <param name="e">事件参数</param>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdatePosition();
        }

        /// <summary>
        /// 单位标签点击事件 - 切换单位
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">事件参数</param>
        private void UnitLabel_Click(object sender, EventArgs e)
        {
            // 获取当前单位在数组中的索引
            int currentIndex = Array.IndexOf(_availableUnits, _currentUnit);
            
            // 切换到下一个单位（循环）
            int nextIndex = (currentIndex + 1) % _availableUnits.Length;
            
            // 保存当前值（以厘米为单位）
            decimal currentValueInCm = GetValueInCentimeters();
            
            // 切换单位
            Unit = _availableUnits[nextIndex];
            
            // 恢复值（自动转换到新单位）
            SetValueInCentimeters(currentValueInCm);
        }

        /// <summary>
        /// 鼠标进入标签 - 高亮显示
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">事件参数</param>
        private void UnitLabel_MouseEnter(object sender, EventArgs e)
        {
            if (_unitLabel != null)
            {
                _unitLabel.ForeColor = Color.FromArgb(0, 123, 255); // 蓝色高亮
                _unitLabel.Font = new Font("Microsoft YaHei", 8F, FontStyle.Bold | FontStyle.Underline);
            }
        }

        /// <summary>
        /// 鼠标离开标签 - 恢复原色
        /// </summary>
        /// <param name="sender">事件发送者</param>
        /// <param name="e">事件参数</param>
        private void UnitLabel_MouseLeave(object sender, EventArgs e)
        {
            if (_unitLabel != null)
            {
                _unitLabel.ForeColor = Color.FromArgb(10, 10, 10); // 恢复原色
                _unitLabel.Font = new Font("Microsoft YaHei", 8F, FontStyle.Bold);
            }
        }

        #region ISupportInitialize 实现
        /// <summary>
        /// 开始初始化
        /// </summary>
        public void BeginInit()
        {
            // 设计器初始化开始
        }

        /// <summary>
        /// 结束初始化
        /// </summary>
        public void EndInit()
        {
            // 设计器初始化结束，确保单位标签正确初始化
            if (_unitLabel == null)
            {
                InitializeComponent();
            }
        }
        #endregion
    }
    

    /// <summary>
    /// 标准文本框 - 通用配置类
    /// </summary>
    [System.ComponentModel.DesignerCategory("")]
    public class StandardTextBox : TextBox
    {
        public string DefaultText { get; private set; }
        public bool IsReadOnly { get; private set; }
        public new int MaxLength { get; set; }

        public StandardTextBox() : this(null, false, 0)
        {
        }

        public StandardTextBox(string defaultText = null, bool readOnly = false, int maxLength = 0)
        {
            DefaultText = defaultText;
            IsReadOnly = readOnly;
            this.MaxLength = maxLength;
            
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // 基础样式设置
            this.Font = new Font("Microsoft YaHei", 9F);
            this.BackColor = Color.FromArgb(250, 250, 250);
            this.BorderStyle = BorderStyle.FixedSingle;
            this.Size = new Size(100, 25);
            
            // 设置属性
            if (!string.IsNullOrEmpty(DefaultText))
            {
                this.Text = DefaultText;
            }
            
            this.ReadOnly = IsReadOnly;
            
            if (this.MaxLength > 0)
            {
                base.MaxLength = this.MaxLength;
            }
        }


        /// <summary>
        /// 清空文本
        /// </summary>
        public void ClearText()
        {
            this.Text = string.Empty;
        }

        /// <summary>
        /// 检查是否为空
        /// </summary>
        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(this.Text);
        }

        /// <summary>
        /// 检查是否为空白
        /// </summary>
        public bool IsWhitespace()
        {
            return string.IsNullOrWhiteSpace(this.Text);
        }

        /// <summary>
        /// 设置只读状态
        /// </summary>
        public void SetReadOnly(bool readOnly)
        {
            this.ReadOnly = readOnly;
        }

        /// <summary>
        /// 设置最大长度
        /// </summary>
        public void SetMaxLength(int maxLength)
        {
            this.MaxLength = maxLength;
        }
    }


    /// <summary>
    /// 标准下拉框 - 通用配置类
    /// </summary>
    [System.ComponentModel.DesignerCategory("")]
    public class StandardComboBox : ComboBox
    {
        public string[] DataItems { get; private set; }
        public int DefaultSelectedIndex { get; private set; }
        public string DefaultSelectedText { get; private set; }
        private bool _allowInput = true;
        public bool AllowInput 
        { 
            get => _allowInput;
            set 
            {
                _allowInput = value;
                this.DropDownStyle = _allowInput ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList;
            }
        }

        public StandardComboBox() : this(null, 0, null, true)
        {
        }

        public StandardComboBox(string[] items, int defaultIndex = 0, string defaultText = null, bool allowInput = true)
        {
            DataItems = items;
            DefaultSelectedIndex = defaultIndex;
            DefaultSelectedText = defaultText;
            _allowInput = allowInput;
            
            InitializeComponent();
            LoadItems();
        }

        private void InitializeComponent()
        {
            // 基础样式设置
            this.Font = new Font("Microsoft YaHei", 9F);
            this.BackColor = Color.FromArgb(250, 250, 250);
            this.DropDownStyle = _allowInput ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList;
            this.Size = new Size(100, 25);
            this.FlatStyle = FlatStyle.Standard;
        }

        private void LoadItems()
        {
            this.Items.Clear();
            
            if (DataItems != null && DataItems.Length > 0)
            {
                foreach (string item in DataItems)
                {
                    this.Items.Add(item);
                }
                
                // 设置默认选中项
                if (!string.IsNullOrEmpty(DefaultSelectedText))
                {
                    SetSelectedItem(DefaultSelectedText);
                }
                else if (DefaultSelectedIndex >= 0 && DefaultSelectedIndex < this.Items.Count)
                {
                    this.SelectedIndex = DefaultSelectedIndex;
                }
                else if (this.Items.Count > 0)
                {
                    this.SelectedIndex = 0;
                }
            }
        }

        /// <summary>
        /// 设置选中项（通过文本）
        /// </summary>
        public void SetSelectedItem(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            
            for (int i = 0; i < this.Items.Count; i++)
            {
                if (string.Equals(this.Items[i].ToString(), text, StringComparison.OrdinalIgnoreCase))
                {
                    this.SelectedIndex = i;
                    return;
                }
            }
        }


        /// <summary>
        /// 获取当前选中项的文本
        /// </summary>
        public string GetSelectedText()
        {
            return this.SelectedItem?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// 添加自定义项目
        /// </summary>
        public void AddCustomItem(string item)
        {
            if (!string.IsNullOrEmpty(item) && !ContainsItem(item))
            {
                this.Items.Add(item);
            }
        }

        /// <summary>
        /// 批量添加自定义项目
        /// </summary>
        public void AddCustomItems(string[] items)
        {
            if (items != null)
            {
                foreach (string item in items)
                {
                    AddCustomItem(item);
                }
            }
        }


        /// <summary>
        /// 检查是否包含指定项目
        /// </summary>
        public bool ContainsItem(string item)
        {
            if (string.IsNullOrEmpty(item)) return false;
            
            foreach (var existingItem in this.Items)
            {
                if (string.Equals(existingItem.ToString(), item, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 获取项目数量
        /// </summary>
        public int GetItemCount()
        {
            return this.Items.Count;
        }

    }
    
    /// <summary>
    /// 切换按钮控件 - 支持按下/释放状态的按钮
    /// </summary>
    [System.ComponentModel.DesignerCategory("")]
    public class ToggleButton : Button
    {
        private bool isPressed;

        /// <summary>
        /// 获取或设置按钮是否被按下
        /// </summary>
        public bool Pressed
        {
            get
            {
                return isPressed;
            }
            set
            {
                if (isPressed != value)
                {
                    isPressed = value;
                    BackColor = (isPressed ? Color.DarkGray : Color.AliceBlue);
                    OnPressedChanged(EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// 按下状态改变事件
        /// </summary>
        public event EventHandler PressedChanged;

        /// <summary>
        /// 初始化切换按钮
        /// </summary>
        public ToggleButton()
        {
            isPressed = false;
            BackColor = Color.AliceBlue;
        }

        /// <summary>
        /// 处理点击事件
        /// </summary>
        /// <param name="e">事件参数</param>
        protected override void OnClick(EventArgs e)
        {
            isPressed = !isPressed;
            BackColor = (isPressed ? Color.DarkGray : Color.AliceBlue);
            OnPressedChanged(EventArgs.Empty);
            base.OnClick(e);
        }

        protected virtual void OnPressedChanged(EventArgs e)
        {
            PressedChanged?.Invoke(this, e);
        }
    }

    #endregion
    /// <summary>
    /// 多级列表控件工厂类 - 专门创建和管理多级列表相关的UI控件
    /// </summary>
    public static class MultiLevelListControlFactory
    {
        /// <summary>
        /// 创建标准数值输入框
        /// </summary>
        public static StandardNumericUpDown CreateNumericInput(Microsoft.Office.Interop.Word.Application wordApp, 
            string name, Point location, Size size, string unit = "厘米")
        {
            return new StandardNumericUpDown(wordApp, unit)
            {
                Name = name,
                Location = location,
                Size = size
            };
        }


        /// <summary>
        /// 创建标准下拉框 - 通用方法
        /// </summary>
        public static StandardComboBox CreateStandardCombo(string name, Point location, string[] items, int defaultIndex = 0, string defaultText = null, bool allowInput = true)
        {
            var combo = new StandardComboBox(items, defaultIndex, defaultText, allowInput);
            combo.Name = name;
            combo.Location = location;
            // 确保 AllowInput 设置生效
            combo.AllowInput = allowInput;
            return combo;
        }

        /// <summary>
        /// 创建编号样式下拉框（使用StandardComboBox）
        /// </summary>
        public static StandardComboBox CreateNumberStyleCombo(string name, Point location, int defaultIndex = 0, string defaultText = null, bool allowInput = false)
        {
            var combo = CreateStandardCombo(name, location, MultiLevelDataManager.ValidationConstants.ValidNumberStyles, defaultIndex, defaultText);
            combo.AllowInput = allowInput;
            return combo;
        }


        /// <summary>
        /// 创建标准文本框 - 通用方法
        /// </summary>
        public static StandardTextBox CreateStandardTextBox(string name, Point location, string defaultText = null, bool readOnly = false, int maxLength = 0)
        {
            return new StandardTextBox(defaultText, readOnly, maxLength)
            {
                Name = name,
                Location = location
            };
        }

        /// <summary>
        /// 创建编号格式文本框（使用StandardTextBox）
        /// </summary>
        public static StandardTextBox CreateNumberFormatTextBox(string name, Point location, string defaultText = null)
        {
            return CreateStandardTextBox(name, location, defaultText, false, 0);
        }

        /// <summary>
        /// 创建只读文本框
        /// </summary>
        public static StandardTextBox CreateReadOnlyTextBox(string name, Point location, string defaultText = null)
        {
            return CreateStandardTextBox(name, location, defaultText, true, 0);
        }

        /// <summary>
        /// 创建限制长度的文本框
        /// </summary>
        public static StandardTextBox CreateLimitedTextBox(string name, Point location, int maxLength, string defaultText = null)
        {
            return CreateStandardTextBox(name, location, defaultText, false, maxLength);
        }

        /// <summary>
        /// 创建编号之后下拉框（使用StandardComboBox）
        /// </summary>
        public static StandardComboBox CreateAfterNumberCombo(string name, Point location, int defaultIndex = 1, string defaultText = null, bool allowInput = false)
        {
            var combo = CreateStandardCombo(name, location, MultiLevelDataManager.ValidationConstants.ValidAfterNumberTypes, defaultIndex, defaultText);
            combo.AllowInput = allowInput;
            return combo;
        }

        /// <summary>
        /// 创建链接样式下拉框（使用StandardComboBox）
        /// </summary>
        public static StandardComboBox CreateLinkedStyleCombo(string name, Point location, int defaultIndex = 0, string defaultText = null, bool allowInput = false)
        {
            var combo = CreateStandardCombo(name, location, MultiLevelDataManager.ValidationConstants.ValidLinkedStyles, defaultIndex, defaultText);
            combo.AllowInput = allowInput;
            return combo;
        }

        /// <summary>
        /// 创建级别标签
        /// </summary>
        public static Label CreateLevelLabel(int level, Point location)
        {
            return new Label
            {
                Text = "第" + level + "级",
                Location = location,
                Size = new Size(50, 20),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 123, 255)
            };
        }

        /// <summary>
        /// 创建列标题标签
        /// </summary>
        public static Label CreateColumnHeader(string text, Point location, Size size)
        {
            return new Label
            {
                Text = text,
                Location = location,
                Size = size,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
        }

        /// <summary>
        /// 创建字体大小下拉框
        /// </summary>
        public static StandardComboBox CreateFontSizeCombo(string name, Point location, int defaultIndex = 0, string defaultText = null)
        {
            return CreateStandardCombo(name, location, MultiLevelDataManager.GetFontSizes(), defaultIndex, defaultText);
        }

        /// <summary>
        /// 创建字体族下拉框
        /// </summary>
        public static StandardComboBox CreateFontFamilyCombo(string name, Point location, int defaultIndex = 0, string defaultText = null)
        {
            return CreateStandardCombo(name, location, MultiLevelDataManager.GetSystemFonts().ToArray(), defaultIndex, defaultText);
        }

        /// <summary>
        /// 创建对齐方式下拉框
        /// </summary>
        /// <param name="name">控件名称</param>
        /// <param name="location">位置</param>
        /// <param name="defaultIndex">默认选中索引</param>
        /// <param name="defaultText">默认选中文本</param>
        /// <returns>标准下拉框</returns>
        public static StandardComboBox CreateAlignmentCombo(string name, Point location, int defaultIndex = 0, string defaultText = null)
        {
            var alignmentOptions = new string[] { "左对齐", "居中对齐", "右对齐", "两端对齐" };
            return CreateStandardCombo(name, location, alignmentOptions, defaultIndex, defaultText);
        }

        /// <summary>
        /// 创建行间距下拉框
        /// </summary>
        /// <param name="name">控件名称</param>
        /// <param name="location">位置</param>
        /// <param name="defaultIndex">默认选中索引</param>
        /// <param name="defaultText">默认选中文本</param>
        /// <returns>标准下拉框</returns>
        public static StandardComboBox CreateLineSpacingCombo(string name, Point location, int defaultIndex = 0, string defaultText = null)
        {
            var lineSpacingOptions = new string[] { "单倍行距", "1.5倍行距", "2倍行距", "最小值", "固定值", "多倍行距" };
            return CreateStandardCombo(name, location, lineSpacingOptions, defaultIndex, defaultText);
        }

        /// <summary>
        /// 创建自定义下拉框
        /// </summary>
        public static StandardComboBox CreateCustomCombo(string name, Point location, string[] items, int defaultIndex = 0, string defaultText = null)
        {
            return CreateStandardCombo(name, location, items, defaultIndex, defaultText);
        }

        /// <summary>
        /// 创建切换按钮
        /// </summary>
        public static ToggleButton CreateToggleButton(string name, Point location, Size size, string text = "", bool pressed = false)
        {
            return new ToggleButton
            {
                Name = name,
                Location = location,
                Size = size,
                Text = text,
                Pressed = pressed
            };
        }

        /// <summary>
        /// 创建字体样式切换按钮（粗体、斜体、下划线）
        /// </summary>
        public static ToggleButton CreateFontStyleToggleButton(string name, Point location, string text, bool pressed = false)
        {
            return CreateToggleButton(name, location, new Size(55, 30), text, pressed);
        }

        /// <summary>
        /// 批量设置输入框的值（使用Word API转换）
        /// </summary>
        /// <param name="container">控件容器</param>
        /// <param name="level">级别</param>
        /// <param name="numberIndent">编号缩进值</param>
        /// <param name="textIndent">文本缩进值</param>
        /// <param name="tabPosition">制表位位置</param>
        public static void SetInputValues(Control container, int level, decimal numberIndent, decimal textIndent, decimal tabPosition)
        {
            var numberIndentControl = container.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as StandardNumericUpDown;
            var textIndentControl = container.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as StandardNumericUpDown;
            var tabPositionControl = container.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as StandardNumericUpDown;

            if (numberIndentControl != null) 
            {
                // 直接设置Value，因为传入的已经是厘米值
                numberIndentControl.Value = numberIndent;
            }
            if (textIndentControl != null) 
            {
                // 直接设置Value，因为传入的已经是厘米值
                textIndentControl.Value = textIndent;
            }
            if (tabPositionControl != null) 
            {
                // 直接设置Value，因为传入的已经是厘米值
                tabPositionControl.Value = tabPosition;
            }
        }

        /// <summary>
        /// 批量获取输入框的值（使用Word API转换）
        /// </summary>
        /// <param name="container">控件容器</param>
        /// <param name="level">级别</param>
        /// <returns>输入框值结构体</returns>
        public static InputValues GetInputValues(Control container, int level)
        {
            var numberIndentControl = container.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as StandardNumericUpDown;
            var textIndentControl = container.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as StandardNumericUpDown;
            var tabPositionControl = container.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as StandardNumericUpDown;

            return new InputValues
            {
                NumberIndent = numberIndentControl?.GetValueInCentimeters() ?? 0,
                TextIndent = textIndentControl?.GetValueInCentimeters() ?? 0,
                TabPosition = tabPositionControl?.GetValueInCentimeters() ?? 0
            };
        }

        #region 常用下拉框项目获取方法

        /// <summary>
        /// 获取编号样式项目
        /// </summary>
        public static string[] GetNumberStyleItems()
        {
            return MultiLevelDataManager.ValidationConstants.ValidNumberStyles;
        }

        /// <summary>
        /// 获取编号之后项目
        /// </summary>
        public static string[] GetAfterNumberItems()
        {
            return MultiLevelDataManager.ValidationConstants.ValidAfterNumberTypes;
        }

        /// <summary>
        /// 获取链接样式项目
        /// </summary>
        public static string[] GetLinkedStyleItems()
        {
            return MultiLevelDataManager.ValidationConstants.ValidLinkedStyles;
        }

        /// <summary>
        /// 获取字体大小项目
        /// </summary>
        public static string[] GetFontSizeItems()
        {
            return MultiLevelDataManager.GetFontSizes();
        }

        /// <summary>
        /// 获取字体族项目
        /// </summary>
        public static string[] GetFontFamilyItems()
        {
            return MultiLevelDataManager.GetSystemFonts().ToArray();
        }

        /// <summary>
        /// 获取对齐方式项目
        /// </summary>
        /// <returns>对齐方式选项数组</returns>
        public static string[] GetAlignmentItems()
        {
            return new string[] { "左对齐", "居中对齐", "右对齐", "两端对齐" };
        }

        /// <summary>
        /// 获取行间距项目
        /// </summary>
        public static string[] GetLineSpacingItems()
        {
            return new string[] { "单倍行距", "1.5倍行距", "2倍行距", "最小值", "固定值", "多倍行距" };
        }

        #endregion

        #region 标准下拉框通用操作方法

        /// <summary>
        /// 批量设置标准下拉框的值
        /// </summary>
        /// <param name="container">控件容器</param>
        /// <param name="comboValues">下拉框值字典</param>
        public static void SetStandardComboBoxValues(Control container, Dictionary<string, string> comboValues)
        {
            foreach (var comboValuePair in comboValues)
            {
                var combo = container.Controls.Find(comboValuePair.Key, true).FirstOrDefault() as StandardComboBox;
                if (combo != null)
                {
                    combo.SetSelectedItem(comboValuePair.Value);
                }
            }
        }

        /// <summary>
        /// 批量获取标准下拉框的值
        /// </summary>
        public static Dictionary<string, string> GetStandardComboBoxValues(Control container, string[] comboNames)
        {
            var values = new Dictionary<string, string>();
            
            foreach (string name in comboNames)
            {
                var combo = container.Controls.Find(name, true).FirstOrDefault() as StandardComboBox;
                if (combo != null)
                {
                    values[name] = combo.GetSelectedText();
                }
            }
            
            return values;
        }

        /// <summary>
        /// 查找标准下拉框
        /// </summary>
        public static StandardComboBox FindStandardComboBox(Control container, string name)
        {
            return container.Controls.Find(name, true).FirstOrDefault() as StandardComboBox;
        }

        /// <summary>
        /// 批量查找标准下拉框
        /// </summary>
        public static Dictionary<string, StandardComboBox> FindStandardComboBoxes(Control container, string[] names)
        {
            var combos = new Dictionary<string, StandardComboBox>();
            
            foreach (string name in names)
            {
                var combo = FindStandardComboBox(container, name);
                if (combo != null)
                {
                    combos[name] = combo;
                }
            }
            
            return combos;
        }

        /// <summary>
        /// 设置下拉框的启用状态
        /// </summary>
        public static void SetComboBoxEnabled(Control container, string name, bool enabled)
        {
            var combo = FindStandardComboBox(container, name);
            if (combo != null)
            {
                combo.Enabled = enabled;
            }
        }

        /// <summary>
        /// 批量设置下拉框的启用状态
        /// </summary>
        /// <param name="container">控件容器</param>
        /// <param name="enabledStates">启用状态字典</param>
        public static void SetComboBoxesEnabled(Control container, Dictionary<string, bool> enabledStates)
        {
            foreach (var enabledStatePair in enabledStates)
            {
                SetComboBoxEnabled(container, enabledStatePair.Key, enabledStatePair.Value);
            }
        }

        #endregion

        #region 标准文本框通用操作方法

        /// <summary>
        /// 批量设置标准文本框的值
        /// </summary>
        /// <param name="container">控件容器</param>
        /// <param name="textValues">文本框值字典</param>
        public static void SetStandardTextBoxValues(Control container, Dictionary<string, string> textValues)
        {
            foreach (var textValuePair in textValues)
            {
                var textBox = container.Controls.Find(textValuePair.Key, true).FirstOrDefault() as StandardTextBox;
                if (textBox != null)
                {
                    textBox.Text = textValuePair.Value ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// 批量获取标准文本框的值
        /// </summary>
        public static Dictionary<string, string> GetStandardTextBoxValues(Control container, string[] textBoxNames)
        {
            var values = new Dictionary<string, string>();
            
            foreach (string name in textBoxNames)
            {
                var textBox = container.Controls.Find(name, true).FirstOrDefault() as StandardTextBox;
                if (textBox != null)
                {
                    values[name] = textBox.Text ?? string.Empty;
                }
            }
            
            return values;
        }

        /// <summary>
        /// 查找标准文本框
        /// </summary>
        public static StandardTextBox FindStandardTextBox(Control container, string name)
        {
            return container.Controls.Find(name, true).FirstOrDefault() as StandardTextBox;
        }

        /// <summary>
        /// 批量查找标准文本框
        /// </summary>
        public static Dictionary<string, StandardTextBox> FindStandardTextBoxes(Control container, string[] names)
        {
            var textBoxes = new Dictionary<string, StandardTextBox>();
            
            foreach (string name in names)
            {
                var textBox = FindStandardTextBox(container, name);
                if (textBox != null)
                {
                    textBoxes[name] = textBox;
                }
            }
            
            return textBoxes;
        }

        /// <summary>
        /// 设置文本框的启用状态
        /// </summary>
        public static void SetTextBoxEnabled(Control container, string name, bool enabled)
        {
            var textBox = FindStandardTextBox(container, name);
            if (textBox != null)
            {
                textBox.Enabled = enabled;
            }
        }

        /// <summary>
        /// 批量设置文本框的启用状态
        /// </summary>
        /// <param name="container">控件容器</param>
        /// <param name="enabledStates">启用状态字典</param>
        public static void SetTextBoxesEnabled(Control container, Dictionary<string, bool> enabledStates)
        {
            foreach (var enabledStatePair in enabledStates)
            {
                SetTextBoxEnabled(container, enabledStatePair.Key, enabledStatePair.Value);
            }
        }

        #endregion

    }
}