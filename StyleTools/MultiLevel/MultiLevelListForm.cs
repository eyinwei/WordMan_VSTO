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
using WordMan;
using WordMan.MultiLevel;

namespace WordMan
{

    public partial class MultiLevelListForm : Form
    {
        #region 常量和配置
        private const int DEFAULT_LEVELS = 0;
        private const int MAX_LEVELS = 9;
        private const decimal DEFAULT_INDENT = 0.8m;
        
        // 魔法数字常量定义
        private const float INVALID_TAB_POSITION = 9999999f;
        private const int DEFAULT_FONT_SIZE = 12;
        private const int MAX_INDENT_CM = 50;
        
        // 使用MultiLevelDataManager中的统一常量定义
        private static readonly WdListNumberStyle[] LevelNumStyle = new WdListNumberStyle[]
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
        
        private static readonly Dictionary<int, LevelConfig> DefaultLevelConfigs = new Dictionary<int, LevelConfig>
        {
            { 1, new LevelConfig("第%1章", "标题 1", 0m) },
            { 2, new LevelConfig("%1.%2", "标题 2", 0m) },
            { 3, new LevelConfig("%1.%2.%3", "标题 3", 0m) },
            { 4, new LevelConfig("%4.", "标题 4", DEFAULT_INDENT) },
            { 5, new LevelConfig("(%5)", "标题 5", DEFAULT_INDENT) },
            { 6, new LevelConfig("%6.", "标题 6", DEFAULT_INDENT, "A,B,C...") },
            { 7, new LevelConfig("(%7)", "标题 7", DEFAULT_INDENT, "A,B,C...") },
            { 8, new LevelConfig("%8.", "标题 8", DEFAULT_INDENT, "a,b,c...") },
            { 9, new LevelConfig("(%9)", "标题 9", DEFAULT_INDENT, "a,b,c...") }
        };
        
        // 使用MultiLevelDataManager中的统一常量定义
        private static readonly string[] NumberStyleOptions = MultiLevelDataManager.ValidationConstants.ValidNumberStyles;
        private static readonly string[] AfterNumberOptions = MultiLevelDataManager.ValidationConstants.ValidAfterNumberTypes;
        private static readonly string[] LinkedStyleOptions = MultiLevelDataManager.ValidationConstants.ValidLinkedStyles;
        #endregion
        
        #region 私有字段
        private int currentLevels = 0;
        private List<LevelData> levelDataList = new List<LevelData>();
        private Microsoft.Office.Interop.Word.Application app;
        private List<WordStyleInfo> levelStyleSettings = new List<WordStyleInfo>();
        #endregion
        
        #region 内部类
        
        /// <summary>
        /// 级别配置类 - 存储每个级别的默认配置信息
        /// </summary>
        private class LevelConfig
        {
            /// <summary>
            /// 编号格式（如：第%1章、%1.%2等）
            /// </summary>
            public string NumberFormat { get; }
            
            /// <summary>
            /// 链接的样式名称
            /// </summary>
            public string LinkedStyle { get; }
            
            /// <summary>
            /// 编号缩进值（厘米）
            /// </summary>
            public decimal NumberIndent { get; }
            
            /// <summary>
            /// 编号样式（如：1,2,3...、A,B,C...等）
            /// </summary>
            public string NumberStyle { get; }
            
            /// <summary>
            /// 初始化级别配置
            /// </summary>
            /// <param name="numberFormat">编号格式</param>
            /// <param name="linkedStyle">链接样式</param>
            /// <param name="numberIndent">编号缩进</param>
            /// <param name="numberStyle">编号样式</param>
            public LevelConfig(string numberFormat, string linkedStyle, decimal numberIndent, string numberStyle = "1,2,3...")
            {
                NumberFormat = numberFormat;
                LinkedStyle = linkedStyle;
                NumberIndent = numberIndent;
                NumberStyle = numberStyle;
            }
        }
        #endregion

        /// <summary>
        /// 初始化多级列表表单
        /// </summary>
        public MultiLevelListForm()
        {
            InitializeComponent();
            
            // 添加空引用检查
            if (Globals.ThisAddIn?.Application == null)
            {
                throw new InvalidOperationException("Word应用程序不可用，无法初始化多级列表表单。");
            }
            
            app = Globals.ThisAddIn.Application;
            
            // 初始化表单数据、事件处理和控件
            InitializeData();
            SetupEventHandlers();
            CreateLevelControls();
        }


        /// <summary>
        /// 初始化表单数据 - 创建默认的级别数据配置
        /// </summary>
        private void InitializeData()
        {
            // 为每个级别创建默认配置数据
            for (int i = 1; i <= MAX_LEVELS; i++)
            {
                var config = DefaultLevelConfigs.ContainsKey(i) ? DefaultLevelConfigs[i] : new LevelConfig("", "无", 0m);
                
                levelDataList.Add(new LevelData
                {
                    Level = i,
                    NumberStyle = config.NumberStyle,
                    NumberFormat = config.NumberFormat,
                    NumberIndent = config.NumberIndent,
                    TextIndent = 0m,
                    AfterNumberType = "空格",
                    TabPosition = 0m,
                    LinkedStyle = config.LinkedStyle
                });
            }

            // 设置默认不显示任何级别
            cmbLevelCount.SetSelectedItem("0");
            currentLevels = 0;
        }

        /// <summary>
        /// 创建级别控件 - 动态生成多级列表配置界面
        /// </summary>
        private void CreateLevelControls()
        {
            levelsContainer.Controls.Clear();

            // 动态创建级别控件 - 按正确顺序（标题在上，1级在下）
            for (int i = currentLevels; i >= 1; i--)
            {
                CreateLevelRow(i);
            }

            // 添加列标题 - 使用MultiLevelListControlFactory创建
            var headerPanel = new Panel();
            headerPanel.Height = 30;
            headerPanel.Dock = DockStyle.Top;
            headerPanel.BackColor = Color.Transparent;
            
            var lblLevel = MultiLevelListControlFactory.CreateColumnHeader("级别", new Point(10, 8), new Size(50, 20));
            var lblNumberStyle = MultiLevelListControlFactory.CreateColumnHeader("编号样式", new Point(70, 8), new Size(100, 20));
            var lblNumberFormat = MultiLevelListControlFactory.CreateColumnHeader("编号格式", new Point(180, 8), new Size(100, 20));
            var lblNumberIndent = MultiLevelListControlFactory.CreateColumnHeader("编号缩进", new Point(290, 8), new Size(100, 20));
            var lblTextIndent = MultiLevelListControlFactory.CreateColumnHeader("文本缩进", new Point(400, 8), new Size(100, 20));
            var lblAfterNumber = MultiLevelListControlFactory.CreateColumnHeader("编号之后", new Point(510, 8), new Size(100, 20));
            var lblTabPosition = MultiLevelListControlFactory.CreateColumnHeader("制表位位置", new Point(620, 8), new Size(100, 20));
            var lblLinkedStyle = MultiLevelListControlFactory.CreateColumnHeader("链接样式", new Point(730, 8), new Size(100, 20));
            
            headerPanel.Controls.AddRange(new Control[] { lblLevel, lblNumberStyle, lblNumberFormat, lblNumberIndent, lblTextIndent, lblAfterNumber, lblTabPosition, lblLinkedStyle });
            levelsContainer.Controls.Add(headerPanel);
            
            // 设置所有级别的制表位位置启用状态
            for (int i = 1; i <= currentLevels; i++)
            {
                var controls = GetLevelControls(i);
                if (controls.AfterNumber != null && controls.TabPosition != null)
                    controls.TabPosition.Enabled = (controls.AfterNumber.Text == "制表位");
            }
        }


        /// <summary>
        /// 创建单个级别的控件行
        /// </summary>
        /// <param name="level">级别编号（1-9）</param>
        private void CreateLevelRow(int level)
        {
            var rowPanel = new Panel();
            rowPanel.Height = 35;
            rowPanel.Dock = DockStyle.Top;
            rowPanel.BackColor = Color.Transparent;
            rowPanel.BorderStyle = BorderStyle.None;

            // 使用MultiLevelListControlFactory创建控件
            var lblLevel = MultiLevelListControlFactory.CreateLevelLabel(level, new Point(10, 8));
            var cmbNumberStyle = MultiLevelListControlFactory.CreateNumberStyleCombo("CmbNumStyle" + level, new Point(70, 5), 0, null, false);
            var txtNumberFormat = MultiLevelListControlFactory.CreateNumberFormatTextBox("TextBoxNumFormat" + level, new Point(180, 5), "");
            var nudNumberIndent = MultiLevelListControlFactory.CreateNumericInput(app, "TxtBoxNumIndent" + level, new Point(290, 5), new Size(100, 25));
            var nudTextIndent = MultiLevelListControlFactory.CreateNumericInput(app, "TxtBoxTextIndent" + level, new Point(400, 5), new Size(100, 25));
            var cmbAfterNumber = MultiLevelListControlFactory.CreateAfterNumberCombo("CmbAfterNumber" + level, new Point(510, 5), 1, null, false);
            var nudTabPosition = MultiLevelListControlFactory.CreateNumericInput(app, "TxtBoxTabPosition" + level, new Point(620, 5), new Size(100, 25));
            var cmbLinkedStyle = MultiLevelListControlFactory.CreateLinkedStyleCombo("CmbLinkedStyle" + level, new Point(730, 5), 0, null, false);

            // 从levelDataList中获取数据并设置
            var levelData = levelDataList[level - 1];
            
            // 设置控件值
            int styleIndex = Array.IndexOf(NumberStyleOptions, levelData.NumberStyle);
            cmbNumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
            txtNumberFormat.Text = levelData.NumberFormat;
            int afterIndex = Array.IndexOf(AfterNumberOptions, levelData.AfterNumberType);
            cmbAfterNumber.SelectedIndex = afterIndex >= 0 ? afterIndex : 1;
            int linkedIndex = Array.IndexOf(LinkedStyleOptions, levelData.LinkedStyle);
            cmbLinkedStyle.SelectedIndex = linkedIndex >= 0 ? linkedIndex : 0;

            // 设置数值（使用Word API转换）
            nudNumberIndent.SetValueInCentimeters(levelData.NumberIndent);
            nudTextIndent.SetValueInCentimeters(levelData.TextIndent);
            nudTabPosition.SetValueInCentimeters(levelData.TabPosition);
            nudTabPosition.Enabled = (levelData.AfterNumberType == "制表位");

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
                var controls = GetLevelControls(level);
                if (controls.AfterNumber != null && controls.TabPosition != null)
                    controls.TabPosition.Enabled = (controls.AfterNumber.Text == "制表位");
            };
            nudTabPosition.ValueChanged += (s, e) => UpdateLevelData(level);
            cmbLinkedStyle.SelectedIndexChanged += (s, e) => UpdateLevelData(level);

            levelsContainer.Controls.Add(rowPanel);
        }

        #region 控件操作辅助方法
        private T FindControl<T>(string name, int level) where T : Control
        {
            return levelsContainer.Controls.Find(name + level, true).FirstOrDefault() as T;
        }
        
        private LevelControls GetLevelControls(int level)
        {
            return new LevelControls
            {
                NumberStyle = FindControl<StandardComboBox>("CmbNumStyle", level),
                NumberFormat = FindControl<StandardTextBox>("TextBoxNumFormat", level),
                AfterNumber = FindControl<StandardComboBox>("CmbAfterNumber", level),
                LinkedStyle = FindControl<StandardComboBox>("CmbLinkedStyle", level),
                NumberIndent = FindControl<StandardNumericUpDown>("TxtBoxNumIndent", level),
                TextIndent = FindControl<StandardNumericUpDown>("TxtBoxTextIndent", level),
                TabPosition = FindControl<StandardNumericUpDown>("TxtBoxTabPosition", level)
            };
        }
        
        
        private class LevelControls
        {
            public StandardComboBox NumberStyle { get; set; }
            public StandardTextBox NumberFormat { get; set; }
            public StandardComboBox AfterNumber { get; set; }
            public StandardComboBox LinkedStyle { get; set; }
            public StandardNumericUpDown NumberIndent { get; set; }
            public StandardNumericUpDown TextIndent { get; set; }
            public StandardNumericUpDown TabPosition { get; set; }
        }
        #endregion
        
        private void UpdateLevelData(int level)
        {
            if (level < 1 || level > levelDataList.Count) return;

            var levelData = levelDataList[level - 1];
            var controls = GetLevelControls(level);
            
            // 使用MultiLevelListControlFactory获取数值（自动使用Word API转换）
            var inputValues = MultiLevelListControlFactory.GetInputValues(levelsContainer, level);
            
            // 更新数据
            levelData.NumberStyle = controls.NumberStyle?.Text ?? levelData.NumberStyle;
            levelData.NumberFormat = controls.NumberFormat?.Text ?? levelData.NumberFormat;
            levelData.AfterNumberType = controls.AfterNumber?.Text ?? levelData.AfterNumberType;
            levelData.LinkedStyle = controls.LinkedStyle?.Text ?? levelData.LinkedStyle;
            levelData.NumberIndent = inputValues.NumberIndent;
            levelData.TextIndent = inputValues.TextIndent;
            levelData.TabPosition = inputValues.TabPosition;
        }




        private void SetupEventHandlers()
        {
            // 底部控制按钮事件
            cmbLevelCount.SelectedIndexChanged += CmbLevelCount_SelectedIndexChanged;
            // btnSetLevelStyle.Click 事件已在设计器中绑定
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
            currentLevels = int.Parse(cmbLevelCount.GetSelectedText());
            CreateLevelControls();
        }



        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string filePath = ConfigurationManager.ShowImportDialog();
            if (!string.IsNullOrEmpty(filePath))
            {
                LoadConfigurationFromFile(filePath);
                MessageBox.Show("配置导入成功！", "导入成功");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string filePath = ConfigurationManager.ShowExportDialog();
            if (!string.IsNullOrEmpty(filePath))
            {
                SaveConfigurationToFile(filePath);
                MessageBox.Show("配置导出成功！", "导出成功");
            }
        }

        private void BtnLoadCurrentList_Click(object sender, EventArgs e)
        {
            // 添加空引用检查
            if (app?.Selection == null)
            {
                MessageBox.Show("Word应用程序或选区不可用！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
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

            cmbLevelCount.SetSelectedItem(maxLevel.ToString());
            currentLevels = maxLevel;
            CreateLevelControls();

            for (int i = 1; i <= maxLevel; i++)
            {
                ListLevel listLevel = listTemplate.ListLevels[i];
                var controls = GetLevelControls(i);

                if (controls.NumberStyle != null)
                {
                    int styleIndex = GetNumberStyleIndex(listLevel.NumberStyle);
                    controls.NumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
                }
                
                if (controls.NumberFormat != null)
                    controls.NumberFormat.Text = listLevel.NumberFormat.ToString();
                
                if (controls.LinkedStyle != null)
                    controls.LinkedStyle.Text = string.IsNullOrEmpty(listLevel.LinkedStyle) ? "无" : listLevel.LinkedStyle;
                
                decimal numberIndent = (decimal)app.PointsToCentimeters(listLevel.NumberPosition);
                decimal textIndent = (decimal)app.PointsToCentimeters(listLevel.TextPosition);
                decimal tabPosition = listLevel.TabPosition != INVALID_TAB_POSITION ? (decimal)app.PointsToCentimeters(listLevel.TabPosition) : 0;
                
                MultiLevelListControlFactory.SetInputValues(levelsContainer, i, numberIndent, textIndent, tabPosition);
            }

            // 清空快捷设置
            ClearQuickSettings();
            MessageBox.Show("已载入当前多级列表设置", "成功");
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
            int levelCount = currentLevels;
            int[] numberStyles = new int[levelCount];
            string[] numberFormats = new string[levelCount];
            string[] linkedStyles = new string[levelCount];
            float[] numberIndents = new float[levelCount];
            float[] textIndents = new float[levelCount];
            string[] afterNumberTypes = new string[levelCount];
            float[] tabPositions = new float[levelCount];

            for (int i = 0; i < levelCount; i++)
            {
                var controls = GetLevelControls(i + 1);

                if (controls.NumberStyle != null)
                    numberStyles[i] = controls.NumberStyle.SelectedIndex;
                
                if (controls.NumberFormat != null)
                {
                    if (!controls.NumberFormat.Text.Contains("%" + (i + 1)))
                    {
                        MessageBox.Show("错误：第" + (i + 1) + "级编号格式未包含本级别的编号！");
                        return;
                    }
                    numberFormats[i] = controls.NumberFormat.Text;
                }
                
                if (controls.LinkedStyle != null)
                {
                    if (i == 0)
                    {
                        linkedStyles[i] = controls.LinkedStyle.Text;
                    }
                    else
                    {
                        if (linkedStyles.Contains(controls.LinkedStyle.Text) && controls.LinkedStyle.Text != "无")
                        {
                            MessageBox.Show("错误：第" + (i + 1) + "级链接样式出现重复！");
                            return;
                        }
                        linkedStyles[i] = controls.LinkedStyle.Text;
                    }
                }
                
                if (controls.AfterNumber != null)
                    afterNumberTypes[i] = controls.AfterNumber.Text;
                
                var inputValues = MultiLevelListControlFactory.GetInputValues(levelsContainer, i + 1);
                numberIndents[i] = (float)inputValues.NumberIndent;
                textIndents[i] = (float)inputValues.TextIndent;
                tabPositions[i] = (float)inputValues.TabPosition;
            }

            // 创建多级列表模板
            CreateListTemplate(levelCount, numberStyles, numberFormats, numberIndents, textIndents, afterNumberTypes, tabPositions, linkedStyles);
        }

        private void CreateListTemplate(int levelCount, int[] numberStyles, string[] numberFormats, 
            float[] numberIndents, float[] textIndents, string[] afterNumberTypes, float[] tabPositions, string[] linkedStyles)
        {
            // 验证参数
            if (levelCount <= 0 || levelCount > MAX_LEVELS)
            {
                throw new ArgumentException($"级别数量无效: {levelCount}，有效范围：1-{MAX_LEVELS}");
            }
            
            // 添加空引用检查
            if (app?.ActiveDocument == null)
            {
                throw new InvalidOperationException("Word文档不可用，无法创建多级列表模板。");
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
            
            // 应用样式设置（如果有的话）
            ApplyLevelStyleSettings();
        }

        /// <summary>
        /// 应用级别样式设置
        /// </summary>
        private void ApplyLevelStyleSettings()
        {
            if (levelStyleSettings == null || levelStyleSettings.Count == 0)
                return;

            try
            {
                var doc = app.ActiveDocument;
                string errorText = string.Empty;
                
                foreach (WordStyleInfo style in levelStyleSettings)
                {
                    if (!style.SetStyle(doc))
                    {
                        errorText += style.StyleName + ";";
                    }
                }
                
                if (!string.IsNullOrEmpty(errorText))
                {
                    MessageBox.Show("样式：" + errorText.TrimEnd(';') + " 应用时出现错误，请检查设置值是否正确！", "样式应用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                // 移除成功提示，静默完成应用
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式设置时出现错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                var controls = GetLevelControls(level);
                
                // 1. 统一缩进设置
                if (chkNumberIndent.Checked && controls.NumberIndent != null)
                    controls.NumberIndent.SetValueInCentimeters(nudNumberIndent.GetValueInCentimeters());
                
                if (chkTextIndent.Checked && controls.TextIndent != null)
                    controls.TextIndent.SetValueInCentimeters(nudTextIndent.GetValueInCentimeters());
                
                if (chkTabPosition.Checked && controls.TabPosition != null)
                    controls.TabPosition.SetValueInCentimeters(nudTabPosition.GetValueInCentimeters());

                // 2. 递进缩进设置
                if (chkProgressiveIndent.Checked && controls.NumberIndent != null)
                {
                    if (level == 1)
                    {
                        controls.NumberIndent.SetValueInCentimeters(nudFirstLevelIndent.GetValueInCentimeters());
                    }
                    else
                    {
                        // 使用Word API进行递进计算
                        decimal baseIndent = nudFirstLevelIndent.GetValueInCentimeters();
                        decimal increment = nudIncrementIndent.GetValueInCentimeters();
                        controls.NumberIndent.SetValueInCentimeters(baseIndent + increment * (level - 1));
                    }
                }

                // 3. 链接标题样式
                if (chkLinkTitles.Checked && controls.LinkedStyle != null)
                    controls.LinkedStyle.SelectedIndex = level;
                else if (chkUnlinkTitles.Checked && controls.LinkedStyle != null)
                    controls.LinkedStyle.SelectedIndex = 0;
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
            nudNumberIndent.Enabled = false;
            nudFirstLevelIndent.Enabled = false;
            nudIncrementIndent.Enabled = false;
            nudTextIndent.Enabled = false;
            nudTabPosition.Enabled = false;
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // 编号缩进使用nudNumberIndent
            if (chkNumberIndent.Checked) // 编号缩进
            {
                nudNumberIndent.Enabled = true;
            }
            else if (!chkNumberIndent.Checked && !chkTextIndent.Checked && !chkTabPosition.Checked)
            {
                nudNumberIndent.Enabled = false;
            }
            
            // 文本缩进使用nudTextIndent
            if (chkTextIndent.Checked) // 文本缩进
            {
                nudTextIndent.Enabled = true;
            }
            else if (!chkNumberIndent.Checked && !chkTextIndent.Checked && !chkTabPosition.Checked)
            {
                nudTextIndent.Enabled = false;
            }
            
            // 制表位位置使用nudTabPosition
            if (chkTabPosition.Checked) // 制表位位置
            {
                nudTabPosition.Enabled = true;
            }
            else if (!chkNumberIndent.Checked && !chkTextIndent.Checked && !chkTabPosition.Checked)
            {
                nudTabPosition.Enabled = false;
            }
        }

        private void ProgressiveIndent_CheckedChanged(object sender, EventArgs e)
        {
            // 递进缩进设置
            if (chkProgressiveIndent.Checked)
            {
                nudFirstLevelIndent.Enabled = true;
                nudIncrementIndent.Enabled = true;
            }
            else
            {
                nudFirstLevelIndent.Enabled = false;
                nudIncrementIndent.Enabled = false;
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
            cmbLevelCount.SetSelectedItem(currentLevels.ToString());
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
                var controls = GetLevelControls(level);
                
                // 更新控件显示
                int styleIndex = Array.IndexOf(NumberStyleOptions, levelData.NumberStyle);
                controls.NumberStyle.SelectedIndex = styleIndex >= 0 ? styleIndex : 0;
                controls.NumberFormat.Text = levelData.NumberFormat;
                int afterIndex = Array.IndexOf(AfterNumberOptions, levelData.AfterNumberType);
                controls.AfterNumber.SelectedIndex = afterIndex >= 0 ? afterIndex : 1;
                int linkedIndex = Array.IndexOf(LinkedStyleOptions, levelData.LinkedStyle);
                controls.LinkedStyle.SelectedIndex = linkedIndex >= 0 ? linkedIndex : 0;

                // 使用MultiLevelListControlFactory设置数值
                MultiLevelListControlFactory.SetInputValues(levelsContainer, level, levelData.NumberIndent, levelData.TextIndent, levelData.TabPosition);
                
                // 更新制表位位置的启用状态
                if (controls.AfterNumber != null && controls.TabPosition != null)
                    controls.TabPosition.Enabled = (controls.AfterNumber.Text == "制表位");
            }
        }

        /// <summary>
        /// 设置每级样式按钮点击事件
        /// </summary>
        private void BtnSetLevelStyle_Click(object sender, EventArgs e)
        {
            // 方案1：打开样式设置窗体（当前实现）
            OpenLevelStyleSettingsForm();
            
            // 方案2：直接应用默认样式（注释掉）
            // ApplyDefaultLevelStyles();
            
            // 方案3：打开Word内置样式对话框（注释掉）
            // OpenWordStyleDialog();
        }

        /// <summary>
        /// 打开级别样式设置窗体
        /// </summary>
        private void OpenLevelStyleSettingsForm()
        {
            var levelStyleSettingsForm = new WordMan.MultiLevel.LevelStyleSettingsForm(currentLevels);
            
            // 如果有现有的样式设置，传递给窗体
            if (levelStyleSettings != null && levelStyleSettings.Count > 0)
            {
                levelStyleSettingsForm.LoadExistingStyles(levelStyleSettings);
            }
            
            if (levelStyleSettingsForm.ShowDialog(this) == DialogResult.OK)
            {
                // 保存样式设置，但不立即应用
                levelStyleSettings = levelStyleSettingsForm.GetLevelStyles();
            }
        }

        /// <summary>
        /// 应用默认级别样式（替代方案）
        /// </summary>
        private void ApplyDefaultLevelStyles()
        {
            var doc = app.ActiveDocument;
            
            // 为每个级别应用默认样式
            for (int i = 1; i <= currentLevels; i++)
            {
                var styleName = $"标题 {i}";
                var style = doc.Styles[styleName];
                
                // 设置字体
                style.Font.Name = "微软雅黑";
                style.Font.Size = 16 - (i - 1) * 2; // 标题1=16磅，标题2=14磅...
                style.Font.Bold = (int)WdConstants.wdToggle;
                
                // 设置段落格式
                style.ParagraphFormat.LeftIndent = app.CentimetersToPoints(i * 0.5f); // 递进缩进
                style.ParagraphFormat.SpaceAfter = 6; // 段后间距
                style.ParagraphFormat.LineSpacing = 1.5f; // 行距
            }
            
        }

        /// <summary>
        /// 打开Word内置样式对话框（替代方案）
        /// </summary>
        private void OpenWordStyleDialog()
        {
            app.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFormatStyle].Show();
        }
    }
}