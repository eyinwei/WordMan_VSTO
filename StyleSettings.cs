using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using CheckBox = System.Windows.Forms.CheckBox;

namespace WordMan_VSTO
{
    public partial class StyleSettings : Form
    {
        // UI设计器
        private StyleSettingsUIDesigner _uiDesigner;

        // 保存所有样式设置
        public Hashtable Title1Settings { get; set; }
        public Hashtable Title2Settings { get; set; }
        public Hashtable Title3Settings { get; set; }
        public Hashtable Title4Settings { get; set; }
        public Hashtable Title5Settings { get; set; }
        public Hashtable Title6Settings { get; set; }
        public Hashtable BodyTextSettings { get; set; }
        public Hashtable BodyTextIndentSettings { get; set; }
        public Hashtable CaptionSettings { get; set; }
        public Hashtable TableTextSettings { get; set; }

        // 页面设置相关属性（固定为GB/T 9704-2012标准）
        public WdPaperSize PaperSize { get; set; } = WdPaperSize.wdPaperA4;
        public WdOrientation PaperDirection { get; set; } = WdOrientation.wdOrientPortrait;
        public float[] PageMargin { get; set; } = new float[] { 3.7f, 3.5f, 2.8f, 2.6f }; // 上、下、左、右（GB/T 9704-2012标准）
        public bool SetGutter { get; set; } = false;
        public WdGutterStyle GutterPosition { get; set; } = WdGutterStyle.wdGutterPosLeft;
        public float GutterValue { get; set; } = 0f;

        // 样式列表与预设样式
        private List<string> _styleNames = new List<string>
        {
            "正文（无缩进）", "正文（缩进）", "一级标题", "二级标题", "三级标题",
            "四级标题", "五级标题", "六级标题", "表中文本", "题注"
        };
        private string[] _presetStyles = { "公文风格" };

        public StyleSettings()
        {
            InitializeComponent();
            this.Text = "样式设置";
            this.Size = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.Font = new Font("微软雅黑", 9F);
            
            // 初始化UI设计器
            _uiDesigner = new StyleSettingsUIDesigner(this);
            _uiDesigner.InitializeAllControls();
            
            // 绑定事件
            BindEvents();

            // 初始化默认样式
            InitializeDefaultStyles();
        }

        /// <summary>
        /// 绑定事件
        /// </summary>
        private void BindEvents()
        {
            // 样式列表选择事件
            var styleList = _uiDesigner.GetControl<ListBox>("lstStyleList");
            if (styleList != null)
                styleList.SelectedIndexChanged += StyleList_SelectedIndexChanged;

            // 应用样式按钮事件
            var applyBtn = _uiDesigner.GetControl<Button>("btnApplyStyle");
            if (applyBtn != null)
                applyBtn.Click += ApplyButton_Click;

            // 读取样式按钮事件
            var loadBtn = _uiDesigner.GetControl<Button>("btnLoadStyle");
            if (loadBtn != null)
                loadBtn.Click += LoadStyle_Click;

            // 保存样式按钮事件
            var saveBtn = _uiDesigner.GetControl<Button>("btnSaveStyle");
            if (saveBtn != null)
                saveBtn.Click += SaveStyle_Click;

            // 添加样式按钮事件
            var addBtn = _uiDesigner.GetControl<Button>("btnAddStyle");
            if (addBtn != null)
                addBtn.Click += AddStyle_Click;

            // 删除样式按钮事件
            var delBtn = _uiDesigner.GetControl<Button>("btnDeleteStyle");
            if (delBtn != null)
                delBtn.Click += DeleteStyle_Click;

            // 选择内置样式按钮事件
            var selectBuiltInBtn = _uiDesigner.GetControl<Button>("btnSelectBuiltIn");
            if (selectBuiltInBtn != null)
                selectBuiltInBtn.Click += SelectBuiltInStyle_Click;

            // 预设样式选择事件
            var presetCmb = _uiDesigner.GetControl<ComboBox>("cmbPresetStyle");
            if (presetCmb != null)
                presetCmb.SelectedIndexChanged += PresetStyle_SelectedIndexChanged;

            // 字体颜色按钮事件
            var fontColorBtn = _uiDesigner.GetControl<Button>("btnFontColor");
            if (fontColorBtn != null)
                fontColorBtn.Click += FontColor_Click;

            // 绑定控件值变化事件
            BindControlValueChangedEvents();
        }

        // 初始化默认样式（按照GB/T 9704-2012标准）
        private void InitializeDefaultStyles()
        {
            // 初始化公文风格样式
            InitializeOfficialDocumentStyle();
        }





        // 预设样式选择事件
        private void PresetStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            var cmb = sender as ComboBox;
            if (cmb == null) return;

            // 只支持公文风格
            if (cmb.SelectedItem.ToString() == "公文风格")
            {
                    InitializeOfficialDocumentStyle();
            }
        }

        // 公文风格初始化（按照GB/T 9704-2012标准）
        private void InitializeOfficialDocumentStyle()
        {
            // 正文样式（GB/T 9704-2012：3号仿宋体字）
            BodyTextSettings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", 16f}, // 三号=16磅
                {"isBold", false}, {"alignment", WdParagraphAlignment.wdAlignParagraphJustify},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f} // 28磅行距（每页22行）
            };

            // 正文缩进样式
            BodyTextIndentSettings = new Hashtable
            {
                {"leftIndent", 0f}, {"firstLineIndent", 32f} // 首行缩进2个字符
            };

            // 一级标题（GB/T 9704-2012：2号小标宋体字）
            Title1Settings = new Hashtable
            {
                {"enFont", "小标宋"}, {"cnFont", "小标宋"}, {"fontSize", 22f}, // 二号=22磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 二级标题（GB/T 9704-2012：3号黑体字）
            Title2Settings = new Hashtable
            {
                {"enFont", "黑体"}, {"cnFont", "黑体"}, {"fontSize", 16f}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 三级标题（GB/T 9704-2012：3号楷体字）
            Title3Settings = new Hashtable
            {
                {"enFont", "楷体"}, {"cnFont", "楷体"}, {"fontSize", 16f}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 四级标题（GB/T 9704-2012：3号仿宋体字）
            Title4Settings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", 16f}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 五级标题（GB/T 9704-2012：3号仿宋体字）
            Title5Settings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", 16f}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 六级标题（GB/T 9704-2012：3号仿宋体字）
            Title6Settings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", 16f}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 表中文本（GB/T 9704-2012：3号仿宋体字）
            TableTextSettings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", 16f}, // 三号=16磅
                {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"lineSpacing", 28f}
            };

            // 题注（GB/T 9704-2012：3号仿宋体字）
            CaptionSettings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", 16f}, // 三号=16磅
                {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            LoadSettings(); // 刷新界面
        }


        // 应用预设样式
        private void ApplyPresetStyle_Click(object sender, EventArgs e)
        {
            var cmb = FindControl<ComboBox>("cmbPreset");
            if (cmb.SelectedIndex == -1)
            {
                MessageBox.Show("请先选择预设样式！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            PresetStyle_SelectedIndexChanged(cmb, EventArgs.Empty);
            MessageBox.Show($"已应用「{cmb.SelectedItem}」预设样式！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 样式列表选择事件
        private void StyleList_SelectionChanged(object sender, EventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null || dgv.SelectedRows.Count == 0) return;

            var selectedRow = dgv.SelectedRows[0];
            var styleName = selectedRow.Cells["Col_StyleName"].Value?.ToString();

            // 根据选择的样式加载对应设置
            switch (styleName)
            {
                case "正文（无缩进）":
                    LoadBodyTextSettings(false);
                    break;
                case "正文（缩进）":
                    LoadBodyTextSettings(true);
                    break;
                case "一级标题":
                    LoadTitleSettings("Title1", Title1Settings);
                    break;
                case "二级标题":
                    LoadTitleSettings("Title2", Title2Settings);
                    break;
                case "三级标题":
                    LoadTitleSettings("Title3", Title3Settings);
                    break;
                case "四级标题":
                    LoadTitleSettings("Title4", Title4Settings);
                    break;
                case "五级标题":
                    LoadTitleSettings("Title5", Title5Settings);
                    break;
                case "六级标题":
                    LoadTitleSettings("Title6", Title6Settings);
                    break;
                case "表中文本":
                    LoadTableTextSettings();
                    break;
                case "题注":
                    LoadCaptionSettings();
                    break;
            }
        }

        // 加载所有样式
        private void LoadSettings()
        {
            // 初始化DataGridView数据
            InitializeStyleListData();

            // 默认加载正文（无缩进）
            LoadBodyTextSettings(false);
        }

        // 初始化样式列表数据
        private void InitializeStyleListData()
        {
            if (_uiDesigner == null) return;
            
            var dgv = _uiDesigner.GetControl<DataGridView>("dgvStyleList");
            if (dgv == null) return;

            // 清空现有数据
            dgv.Rows.Clear();

            // 添加样式数据
            var styleNames = new List<string>
            {
                "正文（无缩进）", "正文（缩进）", "一级标题", "二级标题", "三级标题",
                "四级标题", "五级标题", "六级标题", "表中文本", "题注"
            };

            foreach (var styleName in styleNames)
            {
                dgv.Rows.Add(styleName, "仿宋", "仿宋", "16", false, false, false, "28磅", "0磅", "0磅", "两端对齐");
            }

            // 默认选择第一行
            if (dgv.Rows.Count > 0)
            {
                dgv.Rows[0].Selected = true;
            }
        }

        // 标题加载逻辑
        private void LoadTitleSettings(string prefix, Hashtable settings)
        {
            if (settings == null || _uiDesigner == null) return;

            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            if (cmbEngFont != null)
                cmbEngFont.Text = settings["enFont"].ToString();
                
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            if (cmbChnFont != null)
                cmbChnFont.Text = settings["cnFont"].ToString();
                
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            if (cmbFontSize != null)
                cmbFontSize.Text = settings["fontSize"].ToString();
                
            var boldBtn = FindControl<Button>("btnBold");
            if (boldBtn != null)
                boldBtn.Text = Convert.ToBoolean(settings["isBold"]) ? "是" : "否";

            // 对齐方式
            var cmbAlign = FindControl<ComboBox>("cmbHAlignment");
            if (cmbAlign != null)
            {
            switch ((WdParagraphAlignment)settings["alignment"])
            {
                case WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbAlign.Text = "左对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbAlign.Text = "居中";
                    break;
                case WdParagraphAlignment.wdAlignParagraphRight:
                        cmbAlign.Text = "右对齐";
                        break;
                    case WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbAlign.Text = "两端对齐";
                    break;
                }
            }

            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            if (cmbSpaceBefore != null)
                cmbSpaceBefore.Text = settings["spaceBefore"].ToString() + "磅";
                
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            if (cmbSpaceAfter != null)
                cmbSpaceAfter.Text = settings["spaceAfter"].ToString() + "磅";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = settings["lineSpacing"].ToString() + "磅";
                
            var txtLeftIndent = FindControl<TextBox>("txtLeftIndent");
            if (txtLeftIndent != null)
                txtLeftIndent.Text = "0";
                
            var txtRightIndent = FindControl<TextBox>("txtRightIndent");
            if (txtRightIndent != null)
                txtRightIndent.Text = "0";
        }

        // 加载正文样式（区分缩进/无缩进）
        private void LoadBodyTextSettings(bool useIndent)
        {
            if (BodyTextSettings == null || _uiDesigner == null) return;

            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            if (cmbEngFont != null)
                cmbEngFont.Text = BodyTextSettings["enFont"].ToString();
                
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            if (cmbChnFont != null)
                cmbChnFont.Text = BodyTextSettings["cnFont"].ToString();
                
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            if (cmbFontSize != null)
                cmbFontSize.Text = BodyTextSettings["fontSize"].ToString();
                
            var boldBtn = FindControl<Button>("btnBold");
            if (boldBtn != null)
                boldBtn.Text = Convert.ToBoolean(BodyTextSettings["isBold"]) ? "是" : "否";

            // 对齐方式
            var cmbAlign = FindControl<ComboBox>("cmbHAlignment");
            if (cmbAlign != null)
            {
            switch ((WdParagraphAlignment)BodyTextSettings["alignment"])
            {
                case WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbAlign.Text = "左对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbAlign.Text = "居中";
                    break;
                case WdParagraphAlignment.wdAlignParagraphRight:
                        cmbAlign.Text = "右对齐";
                        break;
                    case WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbAlign.Text = "两端对齐";
                    break;
                }
            }

            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            if (cmbSpaceBefore != null)
                cmbSpaceBefore.Text = BodyTextSettings["spaceBefore"].ToString() + "磅";
                
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            if (cmbSpaceAfter != null)
                cmbSpaceAfter.Text = BodyTextSettings["spaceAfter"].ToString() + "磅";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = BodyTextSettings["lineSpacing"].ToString() + "磅";

            // 缩进设置
            var txtLeftIndent = FindControl<TextBox>("txtLeftIndent");
            var txtRightIndent = FindControl<TextBox>("txtRightIndent");
            
            if (useIndent && BodyTextIndentSettings != null)
            {
                if (txtLeftIndent != null)
                    txtLeftIndent.Text = BodyTextIndentSettings["leftIndent"].ToString();
                if (txtRightIndent != null)
                    txtRightIndent.Text = BodyTextIndentSettings["firstLineIndent"].ToString();
            }
            else
            {
                if (txtLeftIndent != null)
                    txtLeftIndent.Text = "0";
                if (txtRightIndent != null)
                    txtRightIndent.Text = "0";
            }
        }

        // 表中文本加载逻辑
        private void LoadTableTextSettings()
        {
            if (TableTextSettings == null || _uiDesigner == null) return;

            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            if (cmbEngFont != null)
                cmbEngFont.Text = TableTextSettings["enFont"].ToString();
                
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            if (cmbChnFont != null)
                cmbChnFont.Text = TableTextSettings["cnFont"].ToString();
                
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            if (cmbFontSize != null)
                cmbFontSize.Text = TableTextSettings["fontSize"].ToString();
                
            var boldBtn = FindControl<Button>("btnBold");
            if (boldBtn != null)
                boldBtn.Text = "否";

            var cmbAlign = FindControl<ComboBox>("cmbHAlignment");
            if (cmbAlign != null)
            {
            switch ((WdParagraphAlignment)TableTextSettings["alignment"])
            {
                case WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbAlign.Text = "左对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbAlign.Text = "居中";
                    break;
                case WdParagraphAlignment.wdAlignParagraphRight:
                        cmbAlign.Text = "右对齐";
                        break;
                    case WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbAlign.Text = "两端对齐";
                    break;
                }
            }

            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            if (cmbSpaceBefore != null)
                cmbSpaceBefore.Text = "0磅";
                
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            if (cmbSpaceAfter != null)
                cmbSpaceAfter.Text = "0磅";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = TableTextSettings["lineSpacing"].ToString() + "磅";
                
            var txtLeftIndent = FindControl<TextBox>("txtLeftIndent");
            if (txtLeftIndent != null)
                txtLeftIndent.Text = "0";
                
            var txtRightIndent = FindControl<TextBox>("txtRightIndent");
            if (txtRightIndent != null)
                txtRightIndent.Text = "0";
        }

        // 题注加载逻辑
        private void LoadCaptionSettings()
        {
            if (CaptionSettings == null || _uiDesigner == null) return;

            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            if (cmbEngFont != null)
                cmbEngFont.Text = CaptionSettings["enFont"].ToString();
                
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            if (cmbChnFont != null)
                cmbChnFont.Text = CaptionSettings["cnFont"].ToString();
                
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            if (cmbFontSize != null)
                cmbFontSize.Text = CaptionSettings["fontSize"].ToString();
                
            var boldBtn = FindControl<Button>("btnBold");
            if (boldBtn != null)
                boldBtn.Text = "否";

            var cmbAlign = FindControl<ComboBox>("cmbHAlignment");
            if (cmbAlign != null)
            {
            switch ((WdParagraphAlignment)CaptionSettings["alignment"])
            {
                case WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbAlign.Text = "左对齐";
                    break;
                case WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbAlign.Text = "居中";
                    break;
                case WdParagraphAlignment.wdAlignParagraphRight:
                        cmbAlign.Text = "右对齐";
                        break;
                    case WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbAlign.Text = "两端对齐";
                    break;
                }
            }

            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            if (cmbSpaceBefore != null)
                cmbSpaceBefore.Text = CaptionSettings["spaceBefore"].ToString() + "磅";
                
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            if (cmbSpaceAfter != null)
                cmbSpaceAfter.Text = CaptionSettings["spaceAfter"].ToString() + "磅";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = CaptionSettings["lineSpacing"].ToString() + "磅";
                
            var txtLeftIndent = FindControl<TextBox>("txtLeftIndent");
            if (txtLeftIndent != null)
                txtLeftIndent.Text = "0";
                
            var txtRightIndent = FindControl<TextBox>("txtRightIndent");
            if (txtRightIndent != null)
                txtRightIndent.Text = "0";
        }

        // 查找控件
        private T FindControl<T>(string name) where T : Control
        {
            return _uiDesigner.GetControl<T>(name);
        }

        // 应用按钮事件
        private void ApplyButton_Click(object sender, EventArgs e)
        {
            // 保存标题样式
            Title1Settings = GetTitleSettings();
            Title2Settings = GetTitleSettings();
            Title3Settings = GetTitleSettings();
            Title4Settings = GetTitleSettings();
            Title5Settings = GetTitleSettings();
            Title6Settings = GetTitleSettings();

            // 保存正文样式
            BodyTextSettings = GetBodyTextSettings();
            if (FindControl<ListBox>("lstStyles").SelectedItem.ToString() == "正文（缩进）")
            {
                BodyTextIndentSettings = GetBodyTextIndentSettings();
            }

            // 保存表中文本和题注样式
            TableTextSettings = GetTableTextSettings();
            CaptionSettings = GetCaptionSettings();

            // 应用到文档
            ApplySettingsToDocument();

            MessageBox.Show("样式设置已应用到文档！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 获取标题样式设置
        private Hashtable GetTitleSettings()
        {
            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            var boldBtn = FindControl<Button>("btnBold");
            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"isBold", boldBtn?.Text == "是"},
                {"alignment", GetAlignmentValue()},
                {"spaceBefore", GetSpaceValue(cmbSpaceBefore?.Text)},
                {"spaceAfter", GetSpaceValue(cmbSpaceAfter?.Text)},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取正文样式设置
        private Hashtable GetBodyTextSettings()
        {
            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            var boldBtn = FindControl<Button>("btnBold");
            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"isBold", boldBtn?.Text == "是"},
                {"alignment", GetAlignmentValue()},
                {"spaceBefore", GetSpaceValue(cmbSpaceBefore?.Text)},
                {"spaceAfter", GetSpaceValue(cmbSpaceAfter?.Text)},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取正文缩进设置
        private Hashtable GetBodyTextIndentSettings()
        {
            var txtLeftIndent = FindControl<TextBox>("txtLeftIndent");
            var txtRightIndent = FindControl<TextBox>("txtRightIndent");

            return new Hashtable
            {
                {"leftIndent", float.TryParse(txtLeftIndent?.Text, out float leftIndent) ? leftIndent : 0f},
                {"firstLineIndent", float.TryParse(txtRightIndent?.Text, out float rightIndent) ? rightIndent : 0f}
            };
        }

        // 获取表中文本样式
        private Hashtable GetTableTextSettings()
        {
            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"alignment", GetAlignmentValue()},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取题注样式
        private Hashtable GetCaptionSettings()
        {
            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            var cmbSpaceBefore = FindControl<ComboBox>("cmbSpaceBefore");
            var cmbSpaceAfter = FindControl<ComboBox>("cmbSpaceAfter");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"alignment", GetAlignmentValue()},
                {"spaceBefore", GetSpaceValue(cmbSpaceBefore?.Text)},
                {"spaceAfter", GetSpaceValue(cmbSpaceAfter?.Text)},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取间距值（去除"磅"单位）
        private float GetSpaceValue(string spaceText)
        {
            if (string.IsNullOrEmpty(spaceText)) return 0f;
            var cleanText = spaceText.Replace("磅", "").Trim();
            return float.TryParse(cleanText, out float value) ? value : 0f;
        }

        // 获取对齐方式对应的枚举值
        private WdParagraphAlignment GetAlignmentValue()
        {
            var cmbAlign = FindControl<ComboBox>("cmbHAlignment");
            if (cmbAlign == null) return WdParagraphAlignment.wdAlignParagraphLeft;
            
            switch (cmbAlign.Text)
            {
                case "左对齐":
                    return WdParagraphAlignment.wdAlignParagraphLeft;
                case "居中":
                    return WdParagraphAlignment.wdAlignParagraphCenter;
                case "右对齐":
                    return WdParagraphAlignment.wdAlignParagraphRight;
                case "两端对齐":
                    return WdParagraphAlignment.wdAlignParagraphJustify;
                default:
                    return WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }


        // 将设置应用到文档
        private void ApplySettingsToDocument()
        {
            // 获取当前Word应用程序和文档
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null) return;

            // 应用页面设置（固定为GB/T 9704-2012标准）
            doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4; // A4纸张
            doc.PageSetup.Orientation = WdOrientation.wdOrientPortrait; // 竖向
            doc.PageSetup.TopMargin = app.InchesToPoints(3.7f / 2.54f); // 上边距37mm
            doc.PageSetup.BottomMargin = app.InchesToPoints(3.5f / 2.54f); // 下边距35mm
            doc.PageSetup.LeftMargin = app.InchesToPoints(2.8f / 2.54f); // 左边距28mm
            doc.PageSetup.RightMargin = app.InchesToPoints(2.6f / 2.54f); // 右边距26mm
            doc.PageSetup.Gutter = 0; // 无装订线

            // 应用样式设置
            ApplyStyleToDocument("标题 1", Title1Settings, app);
            ApplyStyleToDocument("标题 2", Title2Settings, app);
            ApplyStyleToDocument("标题 3", Title3Settings, app);
            ApplyStyleToDocument("标题 4", Title4Settings, app);
            ApplyStyleToDocument("标题 5", Title5Settings, app);
            ApplyStyleToDocument("标题 6", Title6Settings, app);
            ApplyStyleToDocument("正文", BodyTextSettings, app);
            ApplyStyleToDocument("表中文本", TableTextSettings, app);
            ApplyStyleToDocument("题注", CaptionSettings, app);
        }

        // 应用单个样式到文档
        private void ApplyStyleToDocument(string styleName, Hashtable settings, Microsoft.Office.Interop.Word.Application app)
        {
            if (settings == null) return;

            try
            {
                var style = app.ActiveDocument.Styles[styleName];

                // 设置字体
                style.Font.Name = settings["cnFont"].ToString();
                style.Font.NameAscii = settings["enFont"].ToString();
                style.Font.Size = float.Parse(settings["fontSize"].ToString());
                style.Font.Bold = Convert.ToBoolean(settings["isBold"]) ? 1 : 0;

                // 设置段落格式
                style.ParagraphFormat.Alignment = (WdParagraphAlignment)settings["alignment"];
                style.ParagraphFormat.SpaceBefore = float.Parse(settings["spaceBefore"].ToString());
                style.ParagraphFormat.SpaceAfter = float.Parse(settings["spaceAfter"].ToString());
                style.ParagraphFormat.LineSpacing = float.Parse(settings["lineSpacing"].ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式「{styleName}」时出错：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 当前样式按钮事件
        private void CurrentStyle_Click(object sender, EventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null) return;

            try
            {
                // 从当前文档加载样式
                Title1Settings = GetStyleFromDocument(doc.Styles["标题 1"]);
                Title2Settings = GetStyleFromDocument(doc.Styles["标题 2"]);
                Title3Settings = GetStyleFromDocument(doc.Styles["标题 3"]);
                BodyTextSettings = GetStyleFromDocument(doc.Styles["正文"]);

                // 加载页面设置
                LoadPageSettingsFromDocument(doc);

                LoadSettings();
                MessageBox.Show("已加载当前文档样式！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取当前样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 从文档获取样式设置
        private Hashtable GetStyleFromDocument(Style style)
        {
            return new Hashtable
            {
                {"enFont", style.Font.NameAscii},
                {"cnFont", style.Font.Name},
                {"fontSize", style.Font.Size},
                {"isBold", style.Font.Bold},
                {"alignment", style.ParagraphFormat.Alignment},
                {"spaceBefore", style.ParagraphFormat.SpaceBefore},
                {"spaceAfter", style.ParagraphFormat.SpaceAfter},
                {"lineSpacing", style.ParagraphFormat.LineSpacing}
            };
        }

        // 从文档加载页面设置（固定为GB/T 9704-2012标准）
        private void LoadPageSettingsFromDocument(Document doc)
        {
            // 固定为GB/T 9704-2012标准页面设置
            PageMargin = new float[] { 3.7f, 3.5f, 2.8f, 2.6f }; // 上、下、左、右
            PaperSize = WdPaperSize.wdPaperA4;
            PaperDirection = WdOrientation.wdOrientPortrait;
            SetGutter = false;
            GutterValue = 0f;
            GutterPosition = WdGutterStyle.wdGutterPosLeft;
        }

        // 取消按钮事件
        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 获取当前样式名称
        /// </summary>
        private string GetCurrentStyleName()
        {
            var currentStyleLabel = _uiDesigner.GetControl<Label>("lblCurrentStyleName");
            return currentStyleLabel?.Text ?? "公文风格";
        }

        /// <summary>
        /// 更新当前样式名称显示
        /// </summary>
        private void UpdateCurrentStyleName(string styleName)
        {
            var currentStyleLabel = _uiDesigner.GetControl<Label>("lblCurrentStyleName");
            if (currentStyleLabel != null)
            {
                currentStyleLabel.Text = styleName;
            }
        }


        /// <summary>
        /// 获取对齐方式文本
        /// </summary>
        private string GetAlignmentText(WdParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case WdParagraphAlignment.wdAlignParagraphLeft:
                    return "左对齐";
                case WdParagraphAlignment.wdAlignParagraphCenter:
                    return "居中";
                case WdParagraphAlignment.wdAlignParagraphRight:
                    return "右对齐";
                case WdParagraphAlignment.wdAlignParagraphJustify:
                    return "两端对齐";
                default:
                    return "左对齐";
            }
        }

        /// <summary>
        /// 获取对齐方式枚举值
        /// </summary>
        private WdParagraphAlignment GetAlignmentValue(string alignmentText)
        {
            switch (alignmentText)
            {
                case "左对齐":
                    return WdParagraphAlignment.wdAlignParagraphLeft;
                case "居中":
                    return WdParagraphAlignment.wdAlignParagraphCenter;
                case "右对齐":
                    return WdParagraphAlignment.wdAlignParagraphRight;
                case "两端对齐":
                    return WdParagraphAlignment.wdAlignParagraphJustify;
                default:
                    return WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        #region 缺失的事件处理方法

        /// <summary>
        /// 样式列表选择变化事件
        /// </summary>
        private void StyleList_SelectedIndexChanged(object sender, EventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox == null || listBox.SelectedIndex == -1) return;

            var selectedStyle = listBox.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedStyle)) return;

            // 根据选择的样式加载对应设置
            switch (selectedStyle)
            {
                case "正文（无缩进）":
                    LoadBodyTextSettings(false);
                    break;
                case "正文（缩进）":
                    LoadBodyTextSettings(true);
                    break;
                case "一级标题":
                    LoadTitleSettings("Title1", Title1Settings);
                    break;
                case "二级标题":
                    LoadTitleSettings("Title2", Title2Settings);
                    break;
                case "三级标题":
                    LoadTitleSettings("Title3", Title3Settings);
                    break;
                case "四级标题":
                    LoadTitleSettings("Title4", Title4Settings);
                    break;
                case "五级标题":
                    LoadTitleSettings("Title5", Title5Settings);
                    break;
                case "六级标题":
                    LoadTitleSettings("Title6", Title6Settings);
                    break;
                case "表中文本":
                    LoadTableTextSettings();
                    break;
                case "题注":
                    LoadCaptionSettings();
                    break;
            }
        }

        /// <summary>
        /// 加载样式按钮点击事件
        /// </summary>
        private void LoadStyle_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("请先打开一个Word文档！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 从当前文档加载样式
                Title1Settings = GetStyleFromDocument(doc.Styles["标题 1"]);
                Title2Settings = GetStyleFromDocument(doc.Styles["标题 2"]);
                Title3Settings = GetStyleFromDocument(doc.Styles["标题 3"]);
                Title4Settings = GetStyleFromDocument(doc.Styles["标题 4"]);
                Title5Settings = GetStyleFromDocument(doc.Styles["标题 5"]);
                Title6Settings = GetStyleFromDocument(doc.Styles["标题 6"]);
                BodyTextSettings = GetStyleFromDocument(doc.Styles["正文"]);

                // 加载页面设置
                LoadPageSettingsFromDocument(doc);

                LoadSettings();
                MessageBox.Show("已加载当前文档样式！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 保存样式按钮点击事件
        /// </summary>
        private void SaveStyle_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取当前设置
                var currentStyle = GetCurrentStyleName();
                
                // 保存标题样式
                Title1Settings = GetTitleSettings();
                Title2Settings = GetTitleSettings();
                Title3Settings = GetTitleSettings();
                Title4Settings = GetTitleSettings();
                Title5Settings = GetTitleSettings();
                Title6Settings = GetTitleSettings();

                // 保存正文样式
                BodyTextSettings = GetBodyTextSettings();
                BodyTextIndentSettings = GetBodyTextIndentSettings();

                // 保存表中文本和题注样式
                TableTextSettings = GetTableTextSettings();
                CaptionSettings = GetCaptionSettings();

                // 这里可以添加保存到文件的逻辑
                MessageBox.Show($"样式「{currentStyle}」已保存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 添加样式按钮点击事件
        /// </summary>
        private void AddStyle_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取新样式名称
                string newStyleName = Microsoft.VisualBasic.Interaction.InputBox(
                    "请输入新样式名称：", "添加样式", "新样式", -1, -1);

                if (string.IsNullOrEmpty(newStyleName))
                    return;

                // 检查样式名称是否已存在
                if (_styleNames.Contains(newStyleName))
                {
                    MessageBox.Show("样式名称已存在！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 添加到样式列表
                _styleNames.Add(newStyleName);
                
                // 刷新样式列表显示
                var styleList = _uiDesigner.GetControl<ListBox>("lstStyleList");
                if (styleList != null)
                {
                    styleList.Items.Add(newStyleName);
                }

                MessageBox.Show($"样式「{newStyleName}」已添加！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 删除样式按钮点击事件
        /// </summary>
        private void DeleteStyle_Click(object sender, EventArgs e)
        {
            try
            {
                var styleList = _uiDesigner.GetControl<ListBox>("lstStyleList");
                if (styleList == null || styleList.SelectedIndex == -1)
                {
                    MessageBox.Show("请先选择要删除的样式！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedStyle = styleList.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selectedStyle))
                    return;

                // 确认删除
                var result = MessageBox.Show($"确定要删除样式「{selectedStyle}」吗？", "确认删除", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    // 从列表中移除
                    _styleNames.Remove(selectedStyle);
                    styleList.Items.Remove(selectedStyle);
                    
                    MessageBox.Show($"样式「{selectedStyle}」已删除！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 选择内置样式按钮点击事件
        /// </summary>
        private void SelectBuiltInStyle_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("请先打开一个Word文档！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 显示内置样式选择对话框
                var styleNames = new List<string>();
                foreach (Style style in doc.Styles)
                {
                    if (style.BuiltIn)
                    {
                        styleNames.Add(style.NameLocal);
                    }
                }

                if (styleNames.Count == 0)
                {
                    MessageBox.Show("未找到内置样式！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 这里可以添加样式选择对话框的逻辑
                MessageBox.Show($"找到 {styleNames.Count} 个内置样式", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择内置样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 字体颜色按钮点击事件
        /// </summary>
        private void FontColor_Click(object sender, EventArgs e)
        {
            try
            {
                using (var colorDialog = new ColorDialog())
                {
                    colorDialog.Color = Color.Black; // 默认颜色
                    colorDialog.FullOpen = true;
                    
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        // 更新字体颜色显示
                        var fontColorBtn = sender as Button;
                        if (fontColorBtn != null)
                        {
                            fontColorBtn.BackColor = colorDialog.Color;
                            fontColorBtn.Text = $"颜色: {colorDialog.Color.Name}";
                        }
                        
                        MessageBox.Show($"已选择颜色：{colorDialog.Color.Name}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择字体颜色时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 绑定控件值变化事件
        /// </summary>
        private void BindControlValueChangedEvents()
        {
            try
            {
                // 绑定字体相关控件
                var cmbEngFont = _uiDesigner.GetControl<ComboBox>("cmbEngFontName");
                if (cmbEngFont != null)
                    cmbEngFont.SelectedIndexChanged += ControlValueChanged;

                var cmbChnFont = _uiDesigner.GetControl<ComboBox>("cmbChnFontName");
                if (cmbChnFont != null)
                    cmbChnFont.SelectedIndexChanged += ControlValueChanged;

                var cmbFontSize = _uiDesigner.GetControl<ComboBox>("cmbFontSize");
                if (cmbFontSize != null)
                    cmbFontSize.SelectedIndexChanged += ControlValueChanged;

                // 绑定对齐方式控件
                var cmbAlign = _uiDesigner.GetControl<ComboBox>("cmbHAlignment");
                if (cmbAlign != null)
                    cmbAlign.SelectedIndexChanged += ControlValueChanged;

                // 绑定间距控件
                var cmbSpaceBefore = _uiDesigner.GetControl<ComboBox>("cmbSpaceBefore");
                if (cmbSpaceBefore != null)
                    cmbSpaceBefore.SelectedIndexChanged += ControlValueChanged;

                var cmbSpaceAfter = _uiDesigner.GetControl<ComboBox>("cmbSpaceAfter");
                if (cmbSpaceAfter != null)
                    cmbSpaceAfter.SelectedIndexChanged += ControlValueChanged;

                var cmbLineSpace = _uiDesigner.GetControl<ComboBox>("cmbLineSpace");
                if (cmbLineSpace != null)
                    cmbLineSpace.SelectedIndexChanged += ControlValueChanged;

                // 绑定缩进控件
                var txtLeftIndent = _uiDesigner.GetControl<TextBox>("txtLeftIndent");
                if (txtLeftIndent != null)
                    txtLeftIndent.TextChanged += ControlValueChanged;

                var txtRightIndent = _uiDesigner.GetControl<TextBox>("txtRightIndent");
                if (txtRightIndent != null)
                    txtRightIndent.TextChanged += ControlValueChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"绑定控件事件时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 控件值变化事件处理
        /// </summary>
        private void ControlValueChanged(object sender, EventArgs e)
        {
            try
            {
                // 这里可以添加实时预览逻辑
                // 当用户修改样式设置时，可以实时更新预览
                UpdateStylePreview();
            }
            catch (Exception ex)
            {
                // 静默处理，避免频繁弹窗
                System.Diagnostics.Debug.WriteLine($"控件值变化处理出错：{ex.Message}");
            }
        }

        /// <summary>
        /// 更新样式预览
        /// </summary>
        private void UpdateStylePreview()
        {
            try
            {
                // 这里可以添加样式预览逻辑
                // 例如：更新预览区域的显示
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"更新样式预览出错：{ex.Message}");
            }
        }

        #endregion
    }
}