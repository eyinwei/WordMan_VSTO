using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Color = System.Drawing.Color;
using Point = System.Drawing.Point;

namespace WordMan
{
    public partial class TypesettingTaskPane : UserControl
    {
        // 静态变量：存储唯一的任务窗格实例
        private static Microsoft.Office.Tools.CustomTaskPane _uniqueTaskPane;

        // 动态标题数量配置（默认显示6个标题）
        private static int _maxHeadingLevels = 6;
        
        // 兼容性设置：是否启用消息提示（老版本Office建议设为false）
        private static bool _enableMessageTips = true;

        // 统一的样式常量
        private const int BUTTON_HEIGHT = 44;
        private const int LABEL_HEIGHT = 40;
        private const int SECTION_MARGIN = 20;
        private const int TASK_PANE_WIDTH = 200;
        private const int TITLE_LABEL_WIDTH = 80;
        private const int COMBO_WIDTH = 30;
        private const int COMBO_SPACING = 6;
        private const int COMBO_HEIGHT = 25;
        private const int TIP_DISPLAY_TIME = 1500;
        
        private static readonly Font BUTTON_FONT = new Font("黑体", 11.5f, FontStyle.Bold);
        private static readonly Font LABEL_FONT = new Font("黑体", 12f, FontStyle.Bold);
        private static readonly Font TIP_FONT = new Font("黑体", 9.5f, FontStyle.Regular);
        private static readonly Font COMBO_FONT = new Font("黑体", 9f, FontStyle.Regular);
        
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

            // 动态创建标题按钮
            CreateHeadingButtons();

            // 创建标题样式标签容器（类似文本样式的布局）
            var titleLabelPanel = new Panel
            {
                Height = LABEL_HEIGHT,
                Dock = DockStyle.Top,
                BackColor = LABEL_BACKCOLOR,
                Margin = new Padding(0, 0, 0, 8)
            };

            // 创建整体居中的容器
            var centerContainer = new Panel
            {
                Height = LABEL_HEIGHT,
                Dock = DockStyle.Fill,
                BackColor = LABEL_BACKCOLOR
            };

            // 计算整体宽度和居中位置
            int totalWidth = TITLE_LABEL_WIDTH + COMBO_SPACING + COMBO_WIDTH;

            // 创建标签
            var titleLabel = CreateTitleLabel();

            // 创建下拉框
            var levelCountComboBox = CreateLevelCountComboBox();

            // 添加级别数量选项到下拉框
            for (int i = 1; i <= 9; i++)
            {
                levelCountComboBox.Items.Add($"{i}");
            }

            // 设置当前选中的级别数量
            levelCountComboBox.SelectedIndex = _maxHeadingLevels - 1;

            // 下拉框选择事件
            levelCountComboBox.SelectedIndexChanged += (sender, e) =>
            {
                if (levelCountComboBox.SelectedIndex >= 0)
                {
                    int newLevels = levelCountComboBox.SelectedIndex + 1;
                    if (newLevels != _maxHeadingLevels)
                    {
                        SetMaxHeadingLevels(newLevels);
                        RefreshTitleStylesSection();
                        ShowTemporaryMessage($"已设置为显示{newLevels}个标题");
                    }
                }
            };

            // 将标签和下拉框添加到居中容器
            centerContainer.Controls.Add(titleLabel);
            centerContainer.Controls.Add(levelCountComboBox);

            // 居中容器的事件处理
            centerContainer.Resize += (sender, e) => UpdateTitleControlsPosition(centerContainer, titleLabel, levelCountComboBox);

            // 将居中容器添加到主面板
            titleLabelPanel.Controls.Add(centerContainer);
            this.Controls.Add(titleLabelPanel);
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

        // 创建标题标签
        private Label CreateTitleLabel()
        {
            return new Label
            {
                Text = "标题样式",
                Font = LABEL_FONT,
                ForeColor = LABEL_FORECOLOR,
                BackColor = LABEL_BACKCOLOR,
                Height = LABEL_HEIGHT,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(0, 0),
                Size = new Size(TITLE_LABEL_WIDTH, LABEL_HEIGHT)
            };
        }

        // 创建级别数量下拉框
        private ComboBox CreateLevelCountComboBox()
        {
            return new ComboBox
            {
                Font = COMBO_FONT,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(52, 58, 64),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(TITLE_LABEL_WIDTH + COMBO_SPACING, 8),
                Size = new Size(COMBO_WIDTH, COMBO_HEIGHT),
                Cursor = Cursors.Hand
            };
        }

        // 更新标题控件位置（居中）
        private void UpdateTitleControlsPosition(Panel container, Label titleLabel, ComboBox comboBox)
        {
            int totalWidth = TITLE_LABEL_WIDTH + COMBO_SPACING + COMBO_WIDTH;
            int centerX = (container.Width - totalWidth) / 2;
            
            titleLabel.Location = new Point(centerX, 0);
            comboBox.Location = new Point(centerX + TITLE_LABEL_WIDTH + COMBO_SPACING, 8);
        }

        // 创建标题按钮
        private void CreateHeadingButtons()
        {
            // 逆序添加标题按钮（从最高级到一级，因为使用Dock.Top）
            for (int i = _maxHeadingLevels; i >= 1; i--)
            {
                string titleName = GetHeadingLevelName(i);
                var button = CreateStyleButton(titleName);
                int level = i; // 直接使用级别
                button.Click += (sender, e) => ApplyHeadingStyle(level);
                this.Controls.Add(button);
            }
        }

        // 刷新标题按钮
        private void RefreshHeadingButtons(int insertIndex)
        {
            for (int i = 1; i <= _maxHeadingLevels; i++)
            {
                string titleName = GetHeadingLevelName(i);
                var button = CreateStyleButton(titleName);
                int level = i;
                button.Click += (sender, e) => ApplyHeadingStyle(level);
                this.Controls.Add(button);
                this.Controls.SetChildIndex(button, insertIndex);
            }
        }

        // 获取标题级别名称
        private string GetHeadingLevelName(int level)
        {
            switch (level)
            {
                case 1: return "一级标题";
                case 2: return "二级标题";
                case 3: return "三级标题";
                case 4: return "四级标题";
                case 5: return "五级标题";
                case 6: return "六级标题";
                case 7: return "七级标题";
                case 8: return "八级标题";
                case 9: return "九级标题";
                default: return $"{level}级标题";
            }
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

                // 保存当前光标位置
                Range cursorRange = selection.Range;

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
                    case 7:
                        styleId = WdBuiltinStyle.wdStyleHeading7;
                        break;
                    case 8:
                        styleId = WdBuiltinStyle.wdStyleHeading8;
                        break;
                    case 9:
                        styleId = WdBuiltinStyle.wdStyleHeading9;
                        break;
                    default:
                        styleId = WdBuiltinStyle.wdStyleNormal;
                        break;
                }

                // 对当前段落应用样式
                selection.Paragraphs.set_Style(styleId);

                // 恢复光标位置
                selection.SetRange(cursorRange.Start, cursorRange.Start);

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

                // 保存当前光标位置
                Range cursorRange = selection.Range;

                // 对当前段落应用样式
                selection.Paragraphs.set_Style(WdBuiltinStyle.wdStyleNormal);

                // 如果使用缩进，设置首行缩进两个字符
                if (useIndent)
                {
                    selection.ParagraphFormat.FirstLineIndent = 24f; // 2个字符的缩进（12pt字体对应24pt缩进）
                }
                else
                {
                    selection.ParagraphFormat.FirstLineIndent = 0f;
                }

                // 恢复光标位置
                selection.SetRange(cursorRange.Start, cursorRange.Start);

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

                // 保存当前光标位置
                Range cursorRange = selection.Range;

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

                // 对当前段落应用样式
                selection.Paragraphs.set_Style(tableTextStyle);

                // 恢复光标位置
                selection.SetRange(cursorRange.Start, cursorRange.Start);

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

                // 保存当前光标位置
                Range cursorRange = selection.Range;

                // 对当前段落应用样式
                selection.Paragraphs.set_Style(WdBuiltinStyle.wdStyleCaption);

                // 恢复光标位置
                selection.SetRange(cursorRange.Start, cursorRange.Start);

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
                var multiLevelListType = Type.GetType("WordMan.MultiLevelListForm");
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


        // 显示临时提示消息（优化兼容性）
        private void ShowTemporaryMessage(string message)
        {
            // 如果禁用了消息提示，直接返回
            if (!_enableMessageTips)
                return;

            try
            {
                // 优先使用任务窗格内提示，避免Word失去焦点
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

                // 自动隐藏提示
                var timer = new Timer { Interval = TIP_DISPLAY_TIME };
                timer.Tick += (sender, e) =>
                {
                    try
                    {
                        if (this.Controls.Contains(tipLabel))
                        {
                            this.Controls.Remove(tipLabel);
                        }
                        timer.Dispose();
                    }
                    catch
                    {
                        // 忽略清理时的异常
                    }
                };
                timer.Start();
            }
            catch
            {
                // 如果任务窗格内提示失败，静默处理，不显示消息框
                // 这样可以避免在老版本Office中导致焦点丢失
            }
        }

        // 设置最大标题级别数量
        public static void SetMaxHeadingLevels(int maxLevels)
        {
            if (maxLevels < 1 || maxLevels > 9)
            {
                throw new ArgumentException("标题级别数量必须在1-9之间", nameof(maxLevels));
            }

            _maxHeadingLevels = maxLevels;
        }

        // 刷新标题样式板块
        private void RefreshTitleStylesSection()
        {
            // 移除现有的标题按钮（保留板块标题和分隔线）
            var controlsToRemove = new List<Control>();
            Panel titleLabelPanel = null;
            
            foreach (Control control in this.Controls)
            {
                if (control is Button button && button.Text.Contains("级标题"))
                {
                    controlsToRemove.Add(control);
                }
                else if (control is Panel panel && panel.Controls.Count > 0)
                {
                    // 检查是否是标题样式面板（现在有两层Panel结构）
                    foreach (Control child in panel.Controls)
                    {
                        if (child is Panel centerContainer)
                        {
                            foreach (Control grandChild in centerContainer.Controls)
                            {
                                if (grandChild is Label label && label.Text == "标题样式")
                                {
                                    titleLabelPanel = panel;
                                    break;
                                }
                            }
                        }
                        if (titleLabelPanel != null) break;
                    }
                }
            }

            // 移除标题按钮
            foreach (var control in controlsToRemove)
            {
                this.Controls.Remove(control);
                control.Dispose();
            }

            // 更新下拉框选中状态
            if (titleLabelPanel != null)
            {
                foreach (Control control in titleLabelPanel.Controls)
                {
                    if (control is Panel centerContainer)
                    {
                        foreach (Control child in centerContainer.Controls)
                        {
                            if (child is ComboBox comboBox)
                            {
                                // 更新选中状态为当前级别数量
                                comboBox.SelectedIndex = _maxHeadingLevels - 1;
                                break;
                            }
                        }
                    }
                }
            }

            // 重新创建标题按钮
            if (titleLabelPanel != null)
            {
                // 在标题面板之前插入新的按钮
                var titlePanelIndex = -1;
                for (int i = 0; i < this.Controls.Count; i++)
                {
                    if (this.Controls[i] == titleLabelPanel)
                    {
                        titlePanelIndex = i;
                        break;
                    }
                }

                if (titlePanelIndex >= 0)
                {
                    // 重新创建标题按钮
                    RefreshHeadingButtons(titlePanelIndex);
                }
            }
        }

        // 获取当前最大标题级别数量
        public static int GetMaxHeadingLevels()
        {
            return _maxHeadingLevels;
        }

        // 设置是否启用消息提示（兼容性设置）
        public static void SetMessageTipsEnabled(bool enabled)
        {
            _enableMessageTips = enabled;
        }

        // 获取消息提示状态
        public static bool GetMessageTipsEnabled()
        {
            return _enableMessageTips;
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
                _uniqueTaskPane.Width = TASK_PANE_WIDTH;
            }

            _uniqueTaskPane.Visible = !_uniqueTaskPane.Visible;
        }
    }
}