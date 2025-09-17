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
    /// 输入框辅助类 - 封装所有输入框相关功能
    /// </summary>
    public static class InputHelper
    {
        /// <summary>
        /// 创建带单位的数值输入框
        /// </summary>
        public static NumericUpDownWithUnit CreateNumericInput(Microsoft.Office.Interop.Word.Application wordApp, 
            string name, Point location, Size size, string unit = "厘米")
        {
            return new NumericUpDownWithUnit(wordApp, unit)
            {
                Name = name,
                Location = location,
                Size = size
            };
        }

        /// <summary>
        /// 创建编号样式下拉框
        /// </summary>
        public static StyledComboBox CreateNumberStyleCombo(string name, Point location, string[] items = null)
        {
            var combo = new StyledComboBox
            {
                Name = name,
                Location = location
            };
            
            // 添加项目到下拉框
            var itemsToAdd = items ?? ValidationConstants.ValidNumberStyles;
            
            foreach (string item in itemsToAdd)
            {
                combo.Items.Add(item);
            }
            
            combo.SelectedIndex = 0;
            return combo;
        }

        /// <summary>
        /// 创建编号格式文本框
        /// </summary>
        public static StyledTextBox CreateNumberFormatTextBox(string name, Point location)
        {
            return new StyledTextBox
            {
                Name = name,
                Location = location
            };
        }

        /// <summary>
        /// 创建编号之后下拉框
        /// </summary>
        public static StyledComboBox CreateAfterNumberCombo(string name, Point location)
        {
            var combo = new StyledComboBox
            {
                Name = name,
                Location = location
            };
            
            // 添加项目到下拉框
            foreach (string item in ValidationConstants.ValidAfterNumberTypes)
            {
                combo.Items.Add(item);
            }
            
            combo.SelectedIndex = 1; // 默认选择"空格"
            return combo;
        }

        /// <summary>
        /// 创建链接样式下拉框
        /// </summary>
        public static StyledComboBox CreateLinkedStyleCombo(string name, Point location)
        {
            var combo = new StyledComboBox
            {
                Name = name,
                Location = location
            };
            
            // 添加项目到下拉框
            foreach (string item in ValidationConstants.ValidLinkedStyles)
            {
                combo.Items.Add(item);
            }
            
            combo.SelectedIndex = 0; // 默认选择"无"
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
        /// 批量设置输入框的值（使用Word API转换）
        /// </summary>
        public static void SetInputValues(Control container, int level, decimal numberIndent, decimal textIndent, decimal tabPosition)
        {
            var nudNumberIndent = container.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var nudTextIndent = container.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var nudTabPosition = container.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;

            if (nudNumberIndent != null) 
            {
                // 直接设置Value，因为传入的已经是厘米值
                nudNumberIndent.Value = numberIndent;
            }
            if (nudTextIndent != null) 
            {
                // 直接设置Value，因为传入的已经是厘米值
                nudTextIndent.Value = textIndent;
            }
            if (nudTabPosition != null) 
            {
                // 直接设置Value，因为传入的已经是厘米值
                nudTabPosition.Value = tabPosition;
            }
        }

        /// <summary>
        /// 批量获取输入框的值（使用Word API转换）
        /// </summary>
        public static InputValues GetInputValues(Control container, int level)
        {
            var nudNumberIndent = container.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var nudTextIndent = container.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
            var nudTabPosition = container.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;

            return new InputValues
            {
                NumberIndent = nudNumberIndent?.ValueInCentimeters ?? 0,
                TextIndent = nudTextIndent?.ValueInCentimeters ?? 0,
                TabPosition = nudTabPosition?.ValueInCentimeters ?? 0
            };
        }

        /// <summary>
        /// 切换所有输入框的单位
        /// </summary>
        public static void ChangeAllUnits(Control container, string newUnit, int levelCount)
        {
            for (int level = 1; level <= levelCount; level++)
            {
                var nudNumberIndent = container.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                var nudTextIndent = container.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                var nudTabPosition = container.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;

                if (nudNumberIndent != null)
                {
                    decimal currentValue = nudNumberIndent.ValueInCentimeters;
                    nudNumberIndent.Unit = newUnit;
                    nudNumberIndent.ValueInCentimeters = currentValue;
                }

                if (nudTextIndent != null)
                {
                    decimal currentValue = nudTextIndent.ValueInCentimeters;
                    nudTextIndent.Unit = newUnit;
                    nudTextIndent.ValueInCentimeters = currentValue;
                }

                if (nudTabPosition != null)
                {
                    decimal currentValue = nudTabPosition.ValueInCentimeters;
                    nudTabPosition.Unit = newUnit;
                    nudTabPosition.ValueInCentimeters = currentValue;
                }
            }
        }
    }
}
