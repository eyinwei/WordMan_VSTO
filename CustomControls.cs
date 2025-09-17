using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Point = System.Drawing.Point;
using Font = System.Drawing.Font;

namespace WordMan_VSTO
{
    /// <summary>
    /// 标准样式按钮 - 简化版本
    /// </summary>
    public class StandardButton : Button
    {
        public enum ButtonType
        {
            Primary,    // 主要按钮（蓝色）
            Secondary,  // 次要按钮（黑色）
            Small       // 小按钮（导入导出）
        }

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
            switch (type)
            {
                case ButtonType.Primary:
                    this.BackColor = Color.FromArgb(245, 245, 245);
                    this.FlatAppearance.BorderColor = Color.FromArgb(0, 123, 255);
                    this.FlatAppearance.BorderSize = 1;
                    this.ForeColor = Color.FromArgb(0, 123, 255);
                    break;
                case ButtonType.Secondary:
                    this.BackColor = Color.FromArgb(245, 245, 245);
                    this.FlatAppearance.BorderColor = Color.FromArgb(0, 123, 255);
                    this.FlatAppearance.BorderSize = 1;
                    this.ForeColor = Color.Black;
                    break;
                case ButtonType.Small:
                    this.BackColor = Color.FromArgb(245, 245, 245);
                    this.FlatAppearance.BorderColor = Color.FromArgb(0, 123, 255);
                    this.FlatAppearance.BorderSize = 1;
                    this.ForeColor = Color.Black;
                    this.Size = new Size(50, 35);
                    break;
            }
        }
    }

    /// <summary>
    /// 带单位的数值输入框 - 支持字符、厘米、磅单位转换
    /// </summary>
    public class NumericUpDownWithUnit : NumericUpDown
    {
        private Label _unitLabel;
        private string _currentUnit = "厘米";
        private readonly string[] _availableUnits = { "厘米", "磅", "字符", "行" };

        public string Unit
        {
            get => _currentUnit;
            set
            {
                _currentUnit = value;
                if (_unitLabel != null)
                {
                    _unitLabel.Text = _currentUnit;
                    UpdatePosition();
                }
            }
        }

        public decimal ValueInCentimeters
        {
            get => (decimal)UnitConverter.UnitConvert((double)Value, _currentUnit, "厘米");
            set => Value = (decimal)UnitConverter.UnitConvert((double)value, "厘米", _currentUnit);
        }

        public NumericUpDownWithUnit(Microsoft.Office.Interop.Word.Application wordApp = null, string unit = "厘米")
        {
            _currentUnit = unit;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // 设置输入框样式
            this.TextAlign = HorizontalAlignment.Left;
            this.Width = 100;
            this.Height = 25;
            this.BackColor = Color.FromArgb(250, 250, 250);
            this.DecimalPlaces = 2;
            this.Increment = 0.01m;
            this.Minimum = decimal.MinValue; // 取消下限
            this.Maximum = decimal.MaxValue; // 取消上限
            
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
                using (Graphics g = this.CreateGraphics())
                {
                    SizeF textSize = g.MeasureString(_unitLabel.Text, _unitLabel.Font);
                    int rightMargin = 20; 
                    int x = this.Width - (int)textSize.Width - rightMargin;
                    int y = (this.Height - (int)textSize.Height) / 2;
                    
                    // 确保标签在控件范围内
                    x = Math.Max(0, Math.Min(x, this.Width - (int)textSize.Width));
                    y = Math.Max(0, Math.Min(y, this.Height - (int)textSize.Height));
                    
                    _unitLabel.Location = new Point(x, y);
                    _unitLabel.BringToFront(); // 确保标签在最前面
                }
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdatePosition();
        }

        /// <summary>
        /// 单位标签点击事件 - 切换单位
        /// </summary>
        private void UnitLabel_Click(object sender, EventArgs e)
        {
            // 获取当前单位在数组中的索引
            int currentIndex = Array.IndexOf(_availableUnits, _currentUnit);
            
            // 切换到下一个单位（循环）
            int nextIndex = (currentIndex + 1) % _availableUnits.Length;
            
            // 保存当前值（以厘米为单位）
            decimal currentValueInCm = ValueInCentimeters;
            
            // 切换单位
            Unit = _availableUnits[nextIndex];
            
            // 恢复值（自动转换到新单位）
            ValueInCentimeters = currentValueInCm;
        }

        /// <summary>
        /// 鼠标进入标签 - 高亮显示
        /// </summary>
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
        private void UnitLabel_MouseLeave(object sender, EventArgs e)
        {
            if (_unitLabel != null)
            {
                _unitLabel.ForeColor = Color.FromArgb(10, 10, 10); // 恢复原色
                _unitLabel.Font = new Font("Microsoft YaHei", 8F, FontStyle.Bold);
            }
        }
    }
    
    /// <summary>
    /// 单位转换器 - 使用磅作为中间单位进行转换
    /// </summary>
    public static class UnitConverter
    {
        /// <summary>
        /// 转换数值单位
        /// </summary>
        public static double UnitConvert(double value, string fromUnit, string toUnit)
        {
            // 单位相同直接返回
            if (fromUnit == toUnit)
                return value;

            var wordApp = Globals.ThisAddIn.Application;
            
            // 第一步：将所有单位转换为中间单位（磅）
            double valueInPoints = ConvertToPoints(value, fromUnit, wordApp);
            
            // 第二步：将中间单位（磅）转换为目标单位
            return ConvertFromPoints(valueInPoints, toUnit, wordApp);
        }

        /// 将各种单位转换为磅
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
                    return wordApp.CentimetersToPoints((float)(value * 0.5));
                case "行":
                    // 行转换：1行 ≈ 12磅（标准行距）
                    return value * 12;
                default:
                    return value; // 默认返回原值
            }
        }

        /// 将磅转换为各种单位
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
                    return valueInPoints / 12;
                default:
                    return valueInPoints; // 默认返回原值
            }
        }
    }

    /// <summary>
    /// 样式化文本框
    /// </summary>
    public class StyledTextBox : TextBox
    {
        public StyledTextBox()
        {
            this.Font = new Font("Microsoft YaHei", 9F);
            this.BackColor = Color.FromArgb(250, 250, 250); // 更浅的白灰色
            this.BorderStyle = BorderStyle.FixedSingle;
            this.Size = new Size(100, 25);
        }
    }

    /// <summary>
    /// 样式化下拉框
    /// </summary>
    public class StyledComboBox : ComboBox
    {
        public StyledComboBox()
        {
            this.Font = new Font("Microsoft YaHei", 9F);
            this.BackColor = Color.FromArgb(250, 250, 250); // 更浅的白灰色
            this.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Size = new Size(100, 25);
        }
    }








}