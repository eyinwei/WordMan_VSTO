using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Color = System.Drawing.Color;

namespace WordMan_VSTO
{
    public partial class TypesettingTaskPane : UserControl
    {
        // 静态变量：存储唯一的任务窗格实例
        private static Microsoft.Office.Tools.CustomTaskPane _uniqueTaskPane;

        // 统一的样式常量
        private const int BUTTON_HEIGHT = 44;
        private const int LABEL_HEIGHT = 40;
        private const int SECTION_MARGIN = 20;
        private static readonly Font BUTTON_FONT = new Font("黑体", 11.5f, FontStyle.Bold);
        private static readonly Font LABEL_FONT = new Font("黑体", 12f, FontStyle.Bold);
        private static readonly Font TIP_FONT = new Font("黑体", 9.5f, FontStyle.Regular);
        
        // 简化配色方案
        private static readonly Color BACKGROUND_COLOR = Color.FromArgb(250, 251, 252);
        private static readonly Color BUTTON_BACKCOLOR = Color.White;
        private static readonly Color BUTTON_FORECOLOR = Color.FromArgb(52, 58, 64);
        private static readonly Color BUTTON_BORDERCOLOR = Color.FromArgb(206, 212, 218);
        private static readonly Color BUTTON_HOVER_COLOR = Color.FromArgb(248, 249, 250);
        private static readonly Color BUTTON_PRESSED_COLOR = Color.FromArgb(241, 243, 245);
        private static readonly Color LABEL_BACKCOLOR = Color.FromArgb(73, 80, 87);
        private static readonly Color LABEL_FORECOLOR = Color.White;
        private static readonly Color SEPARATOR_COLOR = Color.FromArgb(233, 236, 239);
        private static readonly Color TIP_BACKCOLOR = Color.FromArgb(40, 167, 69);
        private static readonly Color TIP_FORECOLOR = Color.White;
        private static readonly Color CLOSE_BUTTON_BACKCOLOR = Color.FromArgb(220, 53, 69);
        private static readonly Color CLOSE_BUTTON_FORECOLOR = Color.White;
        private static readonly Color CLOSE_BUTTON_HOVER_COLOR = Color.FromArgb(200, 35, 51);

        public TypesettingTaskPane()
        {
            InitializeComponent();
            this.BackColor = BACKGROUND_COLOR;
            this.Padding = new Padding(12, 8, 12, 8);
            this.AutoScroll = true;

            // 创建三个主要板块（注意：由于使用Dock.Top，需要逆序添加）
            CreateCloseButton();
            CreateSettingsSection();
            CreateTextStylesSection();
            CreateTitleStylesSection();
        }

        // 创建标题样式板块
        private void CreateTitleStylesSection()
        {
            // 添加分隔线
            this.Controls.Add(CreateSeparator());

            // 创建标题按钮（逆序添加，从六级到一级）
            string[] titleNames = { "六级标题", "五级标题", "四级标题", "三级标题", "二级标题", "一级标题" };
            
            for (int i = 0; i < titleNames.Length; i++)
            {
                var button = CreateStyleButton(titleNames[i]);
                int level = 6 - i; // 调整级别计算
                button.Click += (sender, e) => ApplyHeadingStyle(level);
                this.Controls.Add(button);
            }

            // 板块标题
            var titleLabel = CreateSectionLabel("标题样式");
            titleLabel.Margin = new Padding(0, 0, 0, 6);
            this.Controls.Add(titleLabel);
        }

        // 创建文本样式板块
        private void CreateTextStylesSection()
        {
            // 添加分隔线
            this.Controls.Add(CreateSeparator());

            // 创建文本样式按钮（逆序添加）
            var captionButton = CreateStyleButton("题注");
            captionButton.Click += (sender, e) => ApplyCaptionStyle();
            this.Controls.Add(captionButton);

            var tableTextButton = CreateStyleButton("表中文本");
            tableTextButton.Click += (sender, e) => ApplyTableTextStyle();
            this.Controls.Add(tableTextButton);

            var bodyIndentButton = CreateStyleButton("正文(缩进)");
            bodyIndentButton.Click += (sender, e) => ApplyBodyTextStyle(true);
            this.Controls.Add(bodyIndentButton);

            var bodyButton = CreateStyleButton("正文");
            bodyButton.Click += (sender, e) => ApplyBodyTextStyle(false);
            this.Controls.Add(bodyButton);

            // 板块标题
            var textLabel = CreateSectionLabel("文本样式");
            textLabel.Margin = new Padding(0, 0, 0, 6);
            this.Controls.Add(textLabel);
        }

        // 创建样式设置板块
        private void CreateSettingsSection()
        {
            // 多级列表设置按钮
            var multiLevelListButton = CreateStyleButton("多级列表");
            multiLevelListButton.Click += (sender, e) => OpenMultiLevelListForm();
            this.Controls.Add(multiLevelListButton);

            // 样式设置按钮
            var styleSettingsButton = CreateStyleButton("样式设置");
            styleSettingsButton.Click += (sender, e) => OpenStyleSettingsForm();
            this.Controls.Add(styleSettingsButton);

            // 板块标题
            var settingsLabel = CreateSectionLabel("样式设置");
            settingsLabel.Margin = new Padding(0, 0, 0, 6);
            this.Controls.Add(settingsLabel);
        }

        // 创建关闭窗格按钮
        private void CreateCloseButton()
        {
            // 添加分隔线
            this.Controls.Add(CreateSeparator());

            // 创建关闭按钮
            var closeButton = CreateCloseStyleButton("关闭窗格");
            closeButton.Click += (sender, e) => CloseTaskPane();
            this.Controls.Add(closeButton);
        }

        // 创建统一风格的按钮
        private Button CreateStyleButton(string text)
        {
            var button = new Button
            {
                Text = text,
                Height = BUTTON_HEIGHT,
                FlatStyle = FlatStyle.Flat,
                Font = BUTTON_FONT,
                BackColor = BUTTON_BACKCOLOR,
                ForeColor = BUTTON_FORECOLOR,
                Margin = new Padding(0, 2, 0, 2),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Cursor = Cursors.Hand
            };

            // 设置边框样式
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = BUTTON_BORDERCOLOR;
            button.FlatAppearance.MouseOverBackColor = BUTTON_HOVER_COLOR;
            button.FlatAppearance.MouseDownBackColor = BUTTON_PRESSED_COLOR;

                // 添加悬停效果
            button.MouseEnter += (sender, e) =>
                {
                    var btn = sender as Button;
                btn.BackColor = BUTTON_HOVER_COLOR;
                btn.FlatAppearance.BorderColor = Color.FromArgb(173, 181, 189);
                };

            button.MouseLeave += (sender, e) =>
                {
                    var btn = sender as Button;
                    btn.BackColor = BUTTON_BACKCOLOR;
                btn.FlatAppearance.BorderColor = BUTTON_BORDERCOLOR;
            };

            button.MouseDown += (sender, e) =>
                {
                    var btn = sender as Button;
                btn.BackColor = BUTTON_PRESSED_COLOR;
                };

            button.MouseUp += (sender, e) =>
                {
                    var btn = sender as Button;
                btn.BackColor = BUTTON_HOVER_COLOR;
            };

            return button;
        }

        // 创建关闭按钮样式
        private Button CreateCloseStyleButton(string text)
        {
            var button = new Button
            {
                Text = text,
                Height = BUTTON_HEIGHT,
                FlatStyle = FlatStyle.Flat,
                Font = BUTTON_FONT,
                BackColor = CLOSE_BUTTON_BACKCOLOR,
                ForeColor = CLOSE_BUTTON_FORECOLOR,
                Margin = new Padding(0, 2, 0, 2),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Cursor = Cursors.Hand
            };

            // 设置边框样式
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = Color.FromArgb(200, 35, 51);
            button.FlatAppearance.MouseOverBackColor = CLOSE_BUTTON_HOVER_COLOR;
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(180, 20, 36);

            // 添加悬停效果
            button.MouseEnter += (sender, e) =>
            {
                var btn = sender as Button;
                btn.BackColor = CLOSE_BUTTON_HOVER_COLOR;
                btn.FlatAppearance.BorderColor = Color.FromArgb(180, 20, 36);
            };

            button.MouseLeave += (sender, e) =>
            {
                var btn = sender as Button;
                btn.BackColor = CLOSE_BUTTON_BACKCOLOR;
                btn.FlatAppearance.BorderColor = Color.FromArgb(200, 35, 51);
            };

            button.MouseDown += (sender, e) =>
            {
                var btn = sender as Button;
                btn.BackColor = Color.FromArgb(180, 20, 36);
            };

            button.MouseUp += (sender, e) =>
            {
                var btn = sender as Button;
                btn.BackColor = CLOSE_BUTTON_HOVER_COLOR;
            };

            return button;
        }

        // 创建板块标题标签
        private Label CreateSectionLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = LABEL_FONT,
                ForeColor = LABEL_FORECOLOR,
                BackColor = LABEL_BACKCOLOR,
                Height = LABEL_HEIGHT,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Margin = new Padding(0, 0, 0, 8)
            };
        }

        // 创建分隔线
        private Control CreateSeparator()
        {
            return new Panel
            {
                Height = 2,
                BackColor = SEPARATOR_COLOR,
                Dock = DockStyle.Top,
                Margin = new Padding(8, 12, 8, 12)
            };
        }

        // 应用标题样式
        private void ApplyHeadingStyle(int level)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                Document activeDoc = wordApp.ActiveDocument;
                Selection selection = wordApp.Selection;

                if (activeDoc == null || selection == null)
                {
                    ShowTemporaryMessage("无法获取文档或选择内容");
                    return;
                }

                // 选中光标所在的整行
                selection.Expand(WdUnits.wdLine);

                // 设置标题样式
                WdBuiltinStyle styleId;
                switch (level)
                {
                    case 1:
                        styleId = WdBuiltinStyle.wdStyleHeading1;
                        break;
                    case 2:
                        styleId = WdBuiltinStyle.wdStyleHeading2;
                        break;
                    case 3:
                        styleId = WdBuiltinStyle.wdStyleHeading3;
                        break;
                    case 4:
                        styleId = WdBuiltinStyle.wdStyleHeading4;
                        break;
                    case 5:
                        styleId = WdBuiltinStyle.wdStyleHeading5;
                        break;
                    case 6:
                        styleId = WdBuiltinStyle.wdStyleHeading6;
                        break;
                    default:
                        styleId = WdBuiltinStyle.wdStyleNormal;
                        break;
                }

                // 应用样式
                selection.set_Style(styleId);

                // 恢复光标位置到行尾
                selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                ShowTemporaryMessage($"已应用{level}级标题样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用标题样式失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 应用正文样式
        private void ApplyBodyTextStyle(bool useIndent)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                Document activeDoc = wordApp.ActiveDocument;
                Selection selection = wordApp.Selection;

                if (activeDoc == null || selection == null)
                {
                    ShowTemporaryMessage("无法获取文档或选择内容");
                    return;
                }

                // 选中光标所在的整行
                selection.Expand(WdUnits.wdLine);

                // 应用Word默认正文样式
                selection.set_Style(WdBuiltinStyle.wdStyleNormal);

                // 如果使用缩进，设置首行缩进两个字符
                if (useIndent)
                {
                    selection.ParagraphFormat.FirstLineIndent = 24f; // 2个字符的缩进（12pt字体对应24pt缩进）
                }
                else
                {
                    selection.ParagraphFormat.FirstLineIndent = 0f;
                }

                // 恢复光标位置到行尾
                selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                ShowTemporaryMessage(useIndent ? "已应用正文（缩进）样式" : "已应用正文样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用正文样式失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 应用表中文本样式
        private void ApplyTableTextStyle()
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                Document activeDoc = wordApp.ActiveDocument;
                Selection selection = wordApp.Selection;

                if (activeDoc == null || selection == null)
                {
                    ShowTemporaryMessage("无法获取文档或选择内容");
                    return;
                }

                // 选中光标所在的整行
                selection.Expand(WdUnits.wdLine);

            const string tableTextStyleName = "表中文本";

                // 检查是否已存在"表中文本"样式
                Style tableTextStyle = null;
                bool styleExists = false;

                foreach (Style s in activeDoc.Styles)
                {
                    if (s.NameLocal == tableTextStyleName)
                    {
                        tableTextStyle = s;
                        styleExists = true;
                        break;
                    }
                }

                // 如果不存在则创建新样式，基于正文样式
                if (!styleExists)
                {
                    tableTextStyle = activeDoc.Styles.Add(tableTextStyleName, WdStyleType.wdStyleTypeParagraph);
                    object baseStyle = activeDoc.Styles[WdBuiltinStyle.wdStyleNormal];
                    tableTextStyle.set_BaseStyle(ref baseStyle);
                }

                // 应用表中文本样式
                selection.set_Style(tableTextStyle);

                // 恢复光标位置到行尾
                selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                ShowTemporaryMessage("已应用表中文本样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用表中文本样式失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 应用题注样式
        private void ApplyCaptionStyle()
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                Document activeDoc = wordApp.ActiveDocument;
                Selection selection = wordApp.Selection;

                if (activeDoc == null || selection == null)
                {
                    ShowTemporaryMessage("无法获取文档或选择内容");
                    return;
                }

                // 选中光标所在的整行
                selection.Expand(WdUnits.wdLine);

                // 应用内置题注样式
                selection.set_Style(WdBuiltinStyle.wdStyleCaption);

                // 恢复光标位置到行尾
                selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                ShowTemporaryMessage("已应用题注样式");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用题注样式失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 打开样式设置窗体
        private void OpenStyleSettingsForm()
                {
                    try
                    {
            using (var styleForm = new StyleSettings())
            {
                var result = styleForm.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ShowTemporaryMessage("样式设置已更新");
                }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开样式设置窗体失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 打开多级列表设置窗体
        private void OpenMultiLevelListForm()
        {
            try
            {
                // 使用反射来创建窗体，避免编译时依赖问题
                var multiLevelListType = Type.GetType("WordMan_VSTO.MultiLevelListForm");
                if (multiLevelListType != null)
                {
                    using (var multiLevelForm = (Form)Activator.CreateInstance(multiLevelListType))
                    {
                        multiLevelForm.ShowDialog();
                    }
                }
                else
                {
                    MessageBox.Show("无法找到多级列表窗体类", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开多级列表设置窗体失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 关闭任务窗格
        private void CloseTaskPane()
        {
            if (_uniqueTaskPane != null)
            {
                _uniqueTaskPane.Visible = false;
            }
        }


        // 显示临时提示消息
        private void ShowTemporaryMessage(string message)
        {
            var tipLabel = new Label
            {
                Text = message,
                BackColor = TIP_BACKCOLOR,
                ForeColor = TIP_FORECOLOR,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Bottom,
                Height = 34,
                Visible = true,
                Font = TIP_FONT,
                Margin = new Padding(0, 0, 0, 0)
            };

            this.Controls.Add(tipLabel);
            this.Controls.SetChildIndex(tipLabel, 0);

            // 2.5秒后自动隐藏提示
            var timer = new Timer { Interval = 2500 };
            timer.Tick += (sender, e) =>
            {
                this.Controls.Remove(tipLabel);
                timer.Dispose();
            };
            timer.Start();
        }

        // 对外暴露的静态方法：供Ribbon调用
        public static void TriggerShowOrHide()
        {
            if (_uniqueTaskPane == null)
            {
                _uniqueTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(
                  control: new TypesettingTaskPane(),
                  title: "排版工具"
                );

                _uniqueTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _uniqueTaskPane.Width = 240;
            }

            _uniqueTaskPane.Visible = !_uniqueTaskPane.Visible;
        }
    }
}