using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;
using WordMan_VSTO.MultiLevel;

namespace WordMan_VSTO
{





    public partial class MultiLevelListForm : Form
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

        public MultiLevelListForm()
        {
            InitializeComponent();
            app = WordAPIHelper.GetWordApplication();
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
                string numberFormat = "";
                decimal numberIndent = 0m;
                decimal textIndent = 0m; // 默认文本缩进为0
                string afterNumberType = "空格";
                decimal tabPosition = 0m;
                string linkedStyle = "无";
                
                // 设置默认编号格式
                if (i == 1)
                {
                    numberFormat = "第%1章";
                    linkedStyle = "标题 1";
                }
                else if (i == 2)
                {
                    numberFormat = "%1.%2";
                    linkedStyle = "标题 2";
                }
                else if (i == 3)
                {
                    numberFormat = "%1.%2.%3";
                    linkedStyle = "标题 3";
                }
                else if (i == 4)
                {
                    numberFormat = "%4.";
                    numberIndent = 0.8m; // 四到九级编号缩进0.8厘米
                    linkedStyle = "标题 4";
                }
                else if (i == 5)
                {
                    numberFormat = "(%5)";
                    numberIndent = 0.8m;
                    linkedStyle = "标题 5";
                }
                else if (i == 6)
                {
                    numberStyle = "A,B,C...";
                    numberFormat = "%6.";
                    numberIndent = 0.8m;
                    linkedStyle = "标题 6";
                }
                else if (i == 7)
                {
                    numberStyle = "A,B,C...";
                    numberFormat = "(%7)";
                    numberIndent = 0.8m;
                    linkedStyle = "标题 7";
                }
                else if (i == 8)
                {
                    numberStyle = "a,b,c...";
                    numberFormat = "%8.";
                    numberIndent = 0.8m;
                    linkedStyle = "标题 8";
                }
                else if (i == 9)
                {
                    numberStyle = "a,b,c...";
                    numberFormat = "(%9)";
                    numberIndent = 0.8m;
                    linkedStyle = "标题 9";
                }
                
                levelDataList.Add(new LevelData
                {
                    Level = i,
                    NumberStyle = numberStyle,
                    NumberFormat = numberFormat,
                    NumberIndent = numberIndent,
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

            // 添加列标题 - 使用InputHelper创建
            var headerPanel = new Panel();
            headerPanel.Height = 30;
            headerPanel.Dock = DockStyle.Top;
            headerPanel.BackColor = Color.Transparent;
            
            var lblLevel = InputHelper.CreateColumnHeader("级别", new Point(10, 8), new Size(50, 20));
            var lblNumberStyle = InputHelper.CreateColumnHeader("编号样式", new Point(70, 8), new Size(100, 20));
            var lblNumberFormat = InputHelper.CreateColumnHeader("编号格式", new Point(180, 8), new Size(100, 20));
            var lblNumberIndent = InputHelper.CreateColumnHeader("编号缩进", new Point(290, 8), new Size(100, 20));
            var lblTextIndent = InputHelper.CreateColumnHeader("文本缩进", new Point(400, 8), new Size(100, 20));
            var lblAfterNumber = InputHelper.CreateColumnHeader("编号之后", new Point(510, 8), new Size(100, 20));
            var lblTabPosition = InputHelper.CreateColumnHeader("制表位位置", new Point(620, 8), new Size(100, 20));
            var lblLinkedStyle = InputHelper.CreateColumnHeader("链接样式", new Point(730, 8), new Size(100, 20));
            
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

            // 使用InputHelper创建控件
            var lblLevel = InputHelper.CreateLevelLabel(level, new Point(10, 8));
            var cmbNumberStyle = InputHelper.CreateNumberStyleCombo("CmbNumStyle" + level, new Point(70, 5));
            var txtNumberFormat = InputHelper.CreateNumberFormatTextBox("TextBoxNumFormat" + level, new Point(180, 5));
            var nudNumberIndent = InputHelper.CreateNumericInput(app, "TxtBoxNumIndent" + level, new Point(290, 5), new Size(100, 25));
            var nudTextIndent = InputHelper.CreateNumericInput(app, "TxtBoxTextIndent" + level, new Point(400, 5), new Size(100, 25));
            var cmbAfterNumber = InputHelper.CreateAfterNumberCombo("CmbAfterNumber" + level, new Point(510, 5));
            var nudTabPosition = InputHelper.CreateNumericInput(app, "TxtBoxTabPosition" + level, new Point(620, 5), new Size(100, 25));
            var cmbLinkedStyle = InputHelper.CreateLinkedStyleCombo("CmbLinkedStyle" + level, new Point(730, 5));

            // 从levelDataList中获取数据并设置
            var levelData = levelDataList[level - 1];
            
            // 设置编号样式
            string[] styleOptions = { "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,贰,叁...", "甲,乙,丙...", "正规编号" };
            int styleIndex = Array.IndexOf(styleOptions, levelData.NumberStyle);
            cmbNumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;

            // 设置编号格式
            txtNumberFormat.Text = levelData.NumberFormat;

            // 设置数值（使用Word API转换）
            nudNumberIndent.ValueInCentimeters = levelData.NumberIndent;
            nudTextIndent.ValueInCentimeters = levelData.TextIndent;
            nudTabPosition.ValueInCentimeters = levelData.TabPosition;

            // 设置编号之后类型
            string[] afterNumberOptions = { "无", "空格", "制表位" };
            int afterNumberIndex = Array.IndexOf(afterNumberOptions, levelData.AfterNumberType);
            cmbAfterNumber.SelectedIndex = afterNumberIndex >= 0 ? afterNumberIndex : 1;

            // 设置制表位位置启用状态
            nudTabPosition.Enabled = (levelData.AfterNumberType == "制表位");

            // 设置链接样式
            string[] linkedStyleOptions = { "无", "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9" };
            int linkedStyleIndex = Array.IndexOf(linkedStyleOptions, levelData.LinkedStyle);
            cmbLinkedStyle.SelectedIndex = linkedStyleIndex >= 0 ? linkedStyleIndex : 0;

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
            var cmbNumberStyle = levelsContainer.Controls.Find("CmbNumStyle" + level, true).FirstOrDefault() as StyledComboBox;
            var txtNumberFormat = levelsContainer.Controls.Find("TextBoxNumFormat" + level, true).FirstOrDefault() as StyledTextBox;
            var cmbAfterNumber = levelsContainer.Controls.Find("CmbAfterNumber" + level, true).FirstOrDefault() as StyledComboBox;
            var cmbLinkedStyle = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as StyledComboBox;

            // 使用InputHelper获取数值（自动使用Word API转换）
            var inputValues = InputHelper.GetInputValues(levelsContainer, level);
            var numberIndent = inputValues.NumberIndent;
            var textIndent = inputValues.TextIndent;
            var tabPosition = inputValues.TabPosition;

            if (cmbNumberStyle != null) levelData.NumberStyle = cmbNumberStyle.Text;
            if (txtNumberFormat != null) levelData.NumberFormat = txtNumberFormat.Text;
            if (cmbAfterNumber != null) levelData.AfterNumberType = cmbAfterNumber.Text;
            if (cmbLinkedStyle != null) levelData.LinkedStyle = cmbLinkedStyle.Text;
            
            // 使用Word API转换的数值
            levelData.NumberIndent = numberIndent;
            levelData.TextIndent = textIndent;
            levelData.TabPosition = tabPosition;
        }

        private void UpdateTabPositionEnabled(int level)
        {
            var cmbAfterNumber = levelsContainer.Controls.Find("CmbAfterNumber" + level, true).FirstOrDefault() as StyledComboBox;
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
            btnImport.Click += btnImport_Click;
            btnExport.Click += btnExport_Click;
            btnApplySettings.Click += BtnApplySettings_Click;

            // 右侧快捷设置事件
            chkNumberIndent.CheckedChanged += CheckBox_CheckedChanged; // 编号缩进
            chkTextIndent.CheckedChanged += CheckBox_CheckedChanged; // 文本缩进
            chkTabPosition.CheckedChanged += CheckBox_CheckedChanged; // 制表位位置
            chkProgressiveIndent.CheckedChanged += ProgressiveIndent_CheckedChanged; // 递进缩进设置
            chkLinkTitles.CheckedChanged += LinkTitles_CheckedChanged; // 链接到标题样式
            chkUnlinkTitles.CheckedChanged += UnlinkTitles_CheckedChanged; // 不链接标题样式
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

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = ConfigurationManager.ShowImportDialog();
                if (filePath != null)
                {
                    LoadConfigurationFromFile(filePath);
                    MessageBox.Show("配置导入成功！", "导入成功");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败：{ex.Message}", "导入错误");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = ConfigurationManager.ShowExportDialog();
                if (filePath != null)
                {
                    SaveConfigurationToFile(filePath);
                    MessageBox.Show("配置导出成功！", "导出成功");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "导出错误");
            }
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

                // 载入数据到控件 - 使用InputHelper简化
                for (int i = 1; i <= maxLevel; i++)
                {
                    ListLevel listLevel = listTemplate.ListLevels[i];
                    
                    var cmbNumberStyle = levelsContainer.Controls.Find("CmbNumStyle" + i, true).FirstOrDefault() as StyledComboBox;
                    var txtNumberFormat = levelsContainer.Controls.Find("TextBoxNumFormat" + i, true).FirstOrDefault() as StyledTextBox;
                    var cmbLinkedStyle = levelsContainer.Controls.Find("CmbLinkedStyle" + i, true).FirstOrDefault() as StyledComboBox;

                    if (cmbNumberStyle != null)
                    {
                        int styleIndex = GetNumberStyleIndex(listLevel.NumberStyle);
                        cmbNumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
                    }
                    
                    if (txtNumberFormat != null)
                        txtNumberFormat.Text = listLevel.NumberFormat.ToString();
                    
                    if (cmbLinkedStyle != null)
                        cmbLinkedStyle.Text = string.IsNullOrEmpty(listLevel.LinkedStyle) ? "无" : listLevel.LinkedStyle;
                    
                    // 使用InputHelper设置数值（自动使用Word API转换）
                    decimal numberIndent = (decimal)app.PointsToCentimeters(listLevel.NumberPosition);
                    decimal textIndent = (decimal)app.PointsToCentimeters(listLevel.TextPosition);
                    decimal tabPosition = listLevel.TabPosition != 9999999f ? (decimal)app.PointsToCentimeters(listLevel.TabPosition) : 0;
                    
                    InputHelper.SetInputValues(levelsContainer, i, numberIndent, textIndent, tabPosition);
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

                // 收集数据 - 使用InputHelper简化
                for (int i = 0; i < levelCount; i++)
                {
                    var numberStyleCombo = levelsContainer.Controls.Find("CmbNumStyle" + (i + 1), true).FirstOrDefault() as StyledComboBox;
                    var numberFormatText = levelsContainer.Controls.Find("TextBoxNumFormat" + (i + 1), true).FirstOrDefault() as StyledTextBox;
                    var afterNumberCombo = levelsContainer.Controls.Find("CmbAfterNumber" + (i + 1), true).FirstOrDefault() as StyledComboBox;
                    var linkedStyleCombo = levelsContainer.Controls.Find("CmbLinkedStyle" + (i + 1), true).FirstOrDefault() as StyledComboBox;

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
                    
                    if (afterNumberCombo != null)
                    {
                        afterNumberTypes[i] = afterNumberCombo.Text;
                    }
                    
                    // 使用InputHelper获取数值（自动使用Word API转换）
                        try
                        {
                        var inputValues = InputHelper.GetInputValues(levelsContainer, i + 1);
                        numberIndents[i] = (float)inputValues.NumberIndent;
                        textIndents[i] = (float)inputValues.TextIndent;
                        tabPositions[i] = (float)inputValues.TabPosition;
                        }
                        catch (Exception ex)
                        {
                        MessageBox.Show($"第{i + 1}级数值转换错误：{ex.Message}", "错误");
                            return;
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
                
                // 设置链接样式 - 直接使用Word样式对象
                if (linkedStyles[i - 1] != "无" && !string.IsNullOrEmpty(linkedStyles[i - 1]))
                {
                    try
                    {
                        // 提取级别数字
                        int level = 0;
                        if (int.TryParse(linkedStyles[i - 1].Replace("标题 ", "").Replace("标题", ""), out level) && level >= 1 && level <= 9)
                        {
                            // 直接使用WdBuiltinStyle枚举引用内置样式
                            WdBuiltinStyle builtInStyleEnum;
                            switch (level)
                            {
                                case 1: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading1; break;
                                case 2: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading2; break;
                                case 3: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading3; break;
                                case 4: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading4; break;
                                case 5: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading5; break;
                                case 6: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading6; break;
                                case 7: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading7; break;
                                case 8: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading8; break;
                                case 9: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading9; break;
                                default: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading1; break;
                            }
                            
                            var style = app.ActiveDocument.Styles[builtInStyleEnum];
                            
                            if (style != null)
                            {
                                listLevel.LinkedStyle = style.NameLocal;
                            }
                            else
                            {
                                listLevel.LinkedStyle = "";
                            }
                        }
                        else
                        {
                            // 如果不是标题样式，尝试通过名称查找
                            string styleName = GetWordStyleName(linkedStyles[i - 1]);
                            listLevel.LinkedStyle = styleName;
                        }
                    }
                    catch (Exception)
                    {
                        listLevel.LinkedStyle = "";
                    }
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
                if (chkNumberIndent.Checked) // 编号缩进
                {
                    var numberIndentControl = levelsContainer.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (numberIndentControl != null)
                        numberIndentControl.ValueInCentimeters = numericUpDownWithUnit1.ValueInCentimeters;
                }
                
                if (chkTextIndent.Checked) // 文本缩进
                {
                    var textIndentControl = levelsContainer.Controls.Find("TxtBoxTextIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (textIndentControl != null)
                        textIndentControl.ValueInCentimeters = numericUpDownWithUnit4.ValueInCentimeters; // 使用numericUpDownWithUnit4（文本缩进输入框）
                }
                
                if (chkTabPosition.Checked) // 制表位位置
                {
                    var tabPositionControl = levelsContainer.Controls.Find("TxtBoxTabPosition" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (tabPositionControl != null)
                        tabPositionControl.ValueInCentimeters = numericUpDownWithUnit5.ValueInCentimeters; // 使用numericUpDownWithUnit5（制表位位置输入框）
                }

                // 2. 递进缩进设置
                if (chkProgressiveIndent.Checked) // 递进缩进设置
                {
                    var numberIndentControl = levelsContainer.Controls.Find("TxtBoxNumIndent" + level, true).FirstOrDefault() as NumericUpDownWithUnit;
                    if (numberIndentControl != null)
                    {
                        if (level == 1)
                        {
                            numberIndentControl.ValueInCentimeters = numericUpDownWithUnit2.ValueInCentimeters; // 一级编号缩进
                        }
                        else
                        {
                            // 使用Word API进行递进计算
                            decimal baseIndent = numericUpDownWithUnit2.ValueInCentimeters;
                            decimal increment = numericUpDownWithUnit3.ValueInCentimeters;
                            numberIndentControl.ValueInCentimeters = baseIndent + increment * (level - 1);
                        }
                    }
                }

                // 3. 链接标题样式
                if (chkLinkTitles.Checked) // 链接到标题样式
                {
                    var linkedStyleControl = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as StyledComboBox;
                    if (linkedStyleControl != null)
                        linkedStyleControl.SelectedIndex = level;
                }
                else if (chkUnlinkTitles.Checked) // 不链接标题样式
                {
                    var linkedStyleControl = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as StyledComboBox;
                    if (linkedStyleControl != null)
                        linkedStyleControl.SelectedIndex = 0;
                }
            }

            // 清空快捷设置
            ClearQuickSettings();
        }

        private void ClearQuickSettings()
        {
            chkNumberIndent.Checked = false;
            chkTextIndent.Checked = false;
            chkTabPosition.Checked = false;
            chkProgressiveIndent.Checked = false;
            chkLinkTitles.Checked = false;
            chkUnlinkTitles.Checked = false;
            numericUpDownWithUnit1.Enabled = false;
            numericUpDownWithUnit2.Enabled = false;
            numericUpDownWithUnit3.Enabled = false;
            numericUpDownWithUnit4.Enabled = false;
            numericUpDownWithUnit5.Enabled = false;
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // 编号缩进使用numericUpDownWithUnit1
            if (chkNumberIndent.Checked) // 编号缩进
            {
                numericUpDownWithUnit1.Enabled = true;
            }
            else if (!chkNumberIndent.Checked && !chkTextIndent.Checked && !chkTabPosition.Checked)
            {
                numericUpDownWithUnit1.Enabled = false;
            }
            
            // 文本缩进使用numericUpDownWithUnit4
            if (chkTextIndent.Checked) // 文本缩进
            {
                numericUpDownWithUnit4.Enabled = true;
            }
            else if (!chkNumberIndent.Checked && !chkTextIndent.Checked && !chkTabPosition.Checked)
            {
                numericUpDownWithUnit4.Enabled = false;
            }
            
            // 制表位位置使用numericUpDownWithUnit5
            if (chkTabPosition.Checked) // 制表位位置
            {
                numericUpDownWithUnit5.Enabled = true;
            }
            else if (!chkNumberIndent.Checked && !chkTextIndent.Checked && !chkTabPosition.Checked)
            {
                numericUpDownWithUnit5.Enabled = false;
            }
        }

        private void ProgressiveIndent_CheckedChanged(object sender, EventArgs e)
        {
            // 递进缩进设置
            if (chkProgressiveIndent.Checked)
            {
                numericUpDownWithUnit2.Enabled = true;
                numericUpDownWithUnit3.Enabled = true;
            }
            else
            {
                numericUpDownWithUnit2.Enabled = false;
                numericUpDownWithUnit3.Enabled = false;
            }
        }

        private void LinkTitles_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLinkTitles.Checked)
            {
                chkUnlinkTitles.Checked = false;
            }
        }

        private void UnlinkTitles_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUnlinkTitles.Checked)
            {
                chkLinkTitles.Checked = false;
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
                return "";
            }

            // 直接使用Word内置样式对象引用
            try
            {
                // 使用WdBuiltinStyle枚举直接引用内置样式
                WdBuiltinStyle builtInStyleEnum;
                switch (level)
                {
                    case 1: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading1; break;
                    case 2: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading2; break;
                    case 3: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading3; break;
                    case 4: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading4; break;
                    case 5: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading5; break;
                    case 6: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading6; break;
                    case 7: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading7; break;
                    case 8: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading8; break;
                    case 9: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading9; break;
                    default: builtInStyleEnum = WdBuiltinStyle.wdStyleHeading1; break;
                }
                
                var builtInStyle = app.ActiveDocument.Styles[builtInStyleEnum];
                
                if (builtInStyle != null)
                {
                    return builtInStyle.NameLocal;
                }
            }
            catch (Exception)
            {
                // 内置样式不可用，继续尝试其他方法
            }

            // 如果内置样式不可用，尝试通过名称查找
            try
            {
                // 尝试英文名称
                string englishName = "Heading " + level;
                var style = app.ActiveDocument.Styles[englishName];
                if (style != null)
                {
                    return style.NameLocal;
                }
            }
            catch
            {
                // 英文样式不存在，继续尝试中文
            }

            try
            {
                // 尝试中文名称
                string chineseName = "标题 " + level;
                var style = app.ActiveDocument.Styles[chineseName];
                if (style != null)
                {
                    return style.NameLocal;
                }
            }
            catch
            {
                // 中文样式不存在
            }

            // 如果都找不到，返回空字符串（表示不链接样式）
            return "";
        }

        /// <summary>
        /// 保存配置到文件
        /// </summary>
        private void SaveConfigurationToFile(string filePath)
        {
            ConfigurationManager.SaveConfigurationToFile(filePath, levelDataList, currentLevels);
        }


        /// <summary>
        /// 从文件加载配置
        /// </summary>
        private void LoadConfigurationFromFile(string filePath)
        {
            ConfigurationManager.LoadConfigurationFromFile(filePath, out levelDataList, out currentLevels);
            
            // 更新界面
            cmbLevelCount.SelectedItem = currentLevels.ToString();
            CreateLevelControls();
            RefreshLevelControls();
            
            // 强制刷新界面
            this.Refresh();
            levelsContainer.Refresh();
        }

        /// <summary>
        /// 刷新级别控件显示
        /// </summary>
        private void RefreshLevelControls()
        {
            for (int level = 1; level <= currentLevels; level++)
            {
                var levelData = levelDataList[level - 1];
                
                // 更新控件显示
                var cmbNumberStyle = levelsContainer.Controls.Find("CmbNumStyle" + level, true).FirstOrDefault() as StyledComboBox;
                var txtNumberFormat = levelsContainer.Controls.Find("TextBoxNumFormat" + level, true).FirstOrDefault() as StyledTextBox;
                var cmbAfterNumber = levelsContainer.Controls.Find("CmbAfterNumber" + level, true).FirstOrDefault() as StyledComboBox;
                var cmbLinkedStyle = levelsContainer.Controls.Find("CmbLinkedStyle" + level, true).FirstOrDefault() as StyledComboBox;

                if (cmbNumberStyle != null)
                {
                    string[] styleOptions = { "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,贰,叁...", "甲,乙,丙...", "正规编号" };
                    int styleIndex = Array.IndexOf(styleOptions, levelData.NumberStyle);
                    cmbNumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
                }

                if (txtNumberFormat != null)
                    txtNumberFormat.Text = levelData.NumberFormat;

                if (cmbAfterNumber != null)
                {
                    string[] afterNumberOptions = { "无", "空格", "制表位" };
                    int afterNumberIndex = Array.IndexOf(afterNumberOptions, levelData.AfterNumberType);
                    cmbAfterNumber.SelectedIndex = afterNumberIndex >= 0 ? afterNumberIndex : 1;
                }

                if (cmbLinkedStyle != null)
                {
                    string[] linkedStyleOptions = { "无", "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9" };
                    int linkedStyleIndex = Array.IndexOf(linkedStyleOptions, levelData.LinkedStyle);
                    cmbLinkedStyle.SelectedIndex = linkedStyleIndex >= 0 ? linkedStyleIndex : 0;
                }

                // 使用InputHelper设置数值
                InputHelper.SetInputValues(levelsContainer, level, levelData.NumberIndent, levelData.TextIndent, levelData.TabPosition);
                
                // 更新制表位位置的启用状态
                UpdateTabPositionEnabled(level);
            }
        }

        /// <summary>
        /// 设置每级样式按钮点击事件
        /// </summary>
        private void btnSetLevelStyle_Click(object sender, EventArgs e)
        {
            try
            {
                // 创建多级段落设置窗体
                WordMan_VSTO.MultiLevel.LevelStyleSettingsForm levelStyleSettingsForm = new WordMan_VSTO.MultiLevel.LevelStyleSettingsForm(currentLevels);
                
                // 显示窗体
                levelStyleSettingsForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开多级段落设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}