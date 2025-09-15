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
            "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "正文", "题注", "表内文字"
        };
        private string[] _presetStyles = { "公文风格", "学术风格" };

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
                
            // 预设样式选择事件
            var presetCombo = _uiDesigner.GetControl<ComboBox>("cmbPresetStyle");
            if (presetCombo != null)
                presetCombo.SelectedIndexChanged += PresetStyle_SelectedIndexChanged;

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

            switch (cmb.SelectedItem?.ToString())
            {
                case "公文风格":
                    InitializeOfficialDocumentStyle();
                    break;
                case "学术风格":
                    InitializeAcademicStyle();
                    break;
            }
            
            // 更新UI显示
            UpdateUIFromCurrentStyles();
            // 更新预览
            _uiDesigner.UpdateStylePreview();
        }

        /// <summary>
        /// 根据当前样式设置更新UI显示
        /// </summary>
        private void UpdateUIFromCurrentStyles()
        {
            // 获取当前选中的样式
            var styleList = _uiDesigner.GetControl<ListBox>("lstStyleList");
            if (styleList?.SelectedItem != null)
            {
                string selectedStyle = styleList.SelectedItem.ToString();
                // 重新加载当前选中的样式到UI
                switch (selectedStyle)
                {
                    case "标题 1":
                        LoadTitleSettings("Title1", Title1Settings);
                        break;
                    case "标题 2":
                        LoadTitleSettings("Title2", Title2Settings);
                        break;
                    case "标题 3":
                        LoadTitleSettings("Title3", Title3Settings);
                        break;
                    case "标题 4":
                        LoadTitleSettings("Title4", Title4Settings);
                        break;
                    case "标题 5":
                        LoadTitleSettings("Title5", Title5Settings);
                        break;
                    case "标题 6":
                        LoadTitleSettings("Title6", Title6Settings);
                        break;
                    case "正文":
                        LoadBodyTextSettings(false);
                        break;
                    case "表内文字":
                        LoadTableTextSettings();
                        break;
                    case "题注":
                        LoadCaptionSettings();
                        break;
                }
            }
        }

        // 公文风格初始化（按照GB/T 9704-2012标准）
        private void InitializeOfficialDocumentStyle()
        {
            // 正文样式（GB/T 9704-2012：3号仿宋体字）
            BodyTextSettings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", "16"}, // 三号=16磅
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
                {"enFont", "小标宋"}, {"cnFont", "小标宋"}, {"fontSize", "22"}, // 二号=22磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 二级标题（GB/T 9704-2012：3号黑体字）
            Title2Settings = new Hashtable
            {
                {"enFont", "黑体"}, {"cnFont", "黑体"}, {"fontSize", "16"}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 三级标题（GB/T 9704-2012：3号楷体字）
            Title3Settings = new Hashtable
            {
                {"enFont", "楷体"}, {"cnFont", "楷体"}, {"fontSize", "16"}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 四级标题（GB/T 9704-2012：3号仿宋体字）
            Title4Settings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", "16"}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 五级标题（GB/T 9704-2012：3号仿宋体字）
            Title5Settings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", "16"}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 六级标题（GB/T 9704-2012：3号仿宋体字）
            Title6Settings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", "16"}, // 三号=16磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            // 表中文本（GB/T 9704-2012：3号仿宋体字）
            TableTextSettings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", "16"}, // 三号=16磅
                {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"lineSpacing", 28f}
            };

            // 题注（GB/T 9704-2012：3号仿宋体字）
            CaptionSettings = new Hashtable
            {
                {"enFont", "仿宋"}, {"cnFont", "仿宋"}, {"fontSize", "16"}, // 三号=16磅
                {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 28f}
            };

            LoadSettings(); // 刷新界面
        }

        // 学术风格初始化（按照学术论文标准）
        private void InitializeAcademicStyle()
        {
            // 正文样式（学术论文：小四宋体字）
            BodyTextSettings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "14"}, // 小四=14磅
                {"isBold", false}, {"alignment", WdParagraphAlignment.wdAlignParagraphJustify},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 20f} // 1.3倍行距
            };

            // 正文缩进样式
            BodyTextIndentSettings = new Hashtable
            {
                {"leftIndent", 0f}, {"firstLineIndent", 28f} // 首行缩进2个字符
            };

            // 一级标题（学术论文：四号黑体字）
            Title1Settings = new Hashtable
            {
                {"enFont", "黑体"}, {"cnFont", "黑体"}, {"fontSize", "18"}, // 四号=18磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 12f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
            };

            // 二级标题（学术论文：小四黑体字）
            Title2Settings = new Hashtable
            {
                {"enFont", "黑体"}, {"cnFont", "黑体"}, {"fontSize", "14"}, // 小四=14磅
                {"isBold", true}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 12f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
            };

            // 三级标题（学术论文：小四宋体字）
            Title3Settings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "14"}, // 小四=14磅
                {"isBold", false}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 12f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
            };

            // 四级标题（学术论文：小四宋体字）
            Title4Settings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "14"}, // 小四=14磅
                {"isBold", false}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 12f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
            };

            // 五级标题（学术论文：小四宋体字）
            Title5Settings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "14"}, // 小四=14磅
                {"isBold", false}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 12f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
            };

            // 六级标题（学术论文：小四宋体字）
            Title6Settings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "14"}, // 小四=14磅
                {"isBold", false}, {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"spaceBefore", 12f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
            };

            // 表中文本（学术论文：五号宋体字）
            TableTextSettings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "12"}, // 五号=12磅
                {"alignment", WdParagraphAlignment.wdAlignParagraphLeft},
                {"lineSpacing", 16f} // 单倍行距
            };

            // 题注（学术论文：小四宋体字）
            CaptionSettings = new Hashtable
            {
                {"enFont", "宋体"}, {"cnFont", "宋体"}, {"fontSize", "14"}, // 小四=14磅
                {"alignment", WdParagraphAlignment.wdAlignParagraphCenter},
                {"spaceBefore", 0f}, {"spaceAfter", 0f}, {"lineSpacing", 20f}
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
                "一级标题", "二级标题", "三级标题", "四级标题", "五级标题", "六级标题",
                "正文（无缩进）", "正文（缩进）", "表中文本", "题注"
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
                
            var chkBold = FindControl<CheckBox>("chkBold");
            if (chkBold != null)
                chkBold.Checked = Convert.ToBoolean(settings["isBold"]);

            // 对齐方式
            var cmbAlign = FindControl<ComboBox>("cmbAlignment");
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

            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            if (txtSpaceBefore != null)
                txtSpaceBefore.Text = settings["spaceBefore"].ToString() + "行";
                
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            if (txtSpaceAfter != null)
                txtSpaceAfter.Text = settings["spaceAfter"].ToString() + "行";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = settings["lineSpacing"].ToString() + "磅";
                
            var txtFirstIndent = FindControl<TextBox>("txtFirstIndent");
            if (txtFirstIndent != null)
                txtFirstIndent.Text = "0字符";
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
                
            var chkBold = FindControl<CheckBox>("chkBold");
            if (chkBold != null)
                chkBold.Checked = Convert.ToBoolean(BodyTextSettings["isBold"]);

            // 对齐方式
            var cmbAlign = FindControl<ComboBox>("cmbAlignment");
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

            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            if (txtSpaceBefore != null)
                txtSpaceBefore.Text = BodyTextSettings["spaceBefore"].ToString() + "行";
                
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            if (txtSpaceAfter != null)
                txtSpaceAfter.Text = BodyTextSettings["spaceAfter"].ToString() + "行";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = BodyTextSettings["lineSpacing"].ToString() + "磅";

            // 缩进设置
            var txtFirstIndent = FindControl<TextBox>("txtFirstIndent");
            
            if (useIndent && BodyTextIndentSettings != null)
            {
                if (txtFirstIndent != null)
                    txtFirstIndent.Text = BodyTextIndentSettings["firstLineIndent"].ToString() + "字符";
            }
            else
            {
                if (txtFirstIndent != null)
                    txtFirstIndent.Text = "0字符";
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
                
            var chkBold = FindControl<CheckBox>("chkBold");
            if (chkBold != null)
                chkBold.Checked = false;

            var cmbAlign = FindControl<ComboBox>("cmbAlignment");
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

            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            if (txtSpaceBefore != null)
                txtSpaceBefore.Text = "0.00行";
                
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            if (txtSpaceAfter != null)
                txtSpaceAfter.Text = "0.00行";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = TableTextSettings["lineSpacing"].ToString() + "磅";
                
            var txtFirstIndent = FindControl<TextBox>("txtFirstIndent");
            if (txtFirstIndent != null)
                txtFirstIndent.Text = "0字符";
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
                
            var chkBold = FindControl<CheckBox>("chkBold");
            if (chkBold != null)
                chkBold.Checked = false;

            var cmbAlign = FindControl<ComboBox>("cmbAlignment");
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

            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            if (txtSpaceBefore != null)
                txtSpaceBefore.Text = CaptionSettings["spaceBefore"].ToString() + "行";
                
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            if (txtSpaceAfter != null)
                txtSpaceAfter.Text = CaptionSettings["spaceAfter"].ToString() + "行";
                
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");
            if (cmbLineSpace != null)
                cmbLineSpace.Text = CaptionSettings["lineSpacing"].ToString() + "磅";
                
            var txtFirstIndent = FindControl<TextBox>("txtFirstIndent");
            if (txtFirstIndent != null)
                txtFirstIndent.Text = "0字符";
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
            // 正文样式统一处理，不再区分缩进/无缩进

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
            var chkBold = FindControl<CheckBox>("chkBold");
            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"isBold", chkBold?.Checked ?? false},
                {"alignment", GetAlignmentValue()},
                {"spaceBefore", GetSpaceValue(txtSpaceBefore?.Text)},
                {"spaceAfter", GetSpaceValue(txtSpaceAfter?.Text)},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取正文样式设置
        private Hashtable GetBodyTextSettings()
        {
            var cmbEngFont = FindControl<ComboBox>("cmbEngFontName");
            var cmbChnFont = FindControl<ComboBox>("cmbChnFontName");
            var cmbFontSize = FindControl<ComboBox>("cmbFontSize");
            var chkBold = FindControl<CheckBox>("chkBold");
            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"isBold", chkBold?.Checked ?? false},
                {"alignment", GetAlignmentValue()},
                {"spaceBefore", GetSpaceValue(txtSpaceBefore?.Text)},
                {"spaceAfter", GetSpaceValue(txtSpaceAfter?.Text)},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取正文缩进设置
        private Hashtable GetBodyTextIndentSettings()
        {
            var txtFirstIndent = FindControl<TextBox>("txtFirstIndent");

            return new Hashtable
            {
                {"leftIndent", 0f},
                {"firstLineIndent", GetFirstLineIndentValue(txtFirstIndent?.Text)}
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
            var txtSpaceBefore = FindControl<TextBox>("txtSpaceBefore");
            var txtSpaceAfter = FindControl<TextBox>("txtSpaceAfter");
            var cmbLineSpace = FindControl<ComboBox>("cmbLineSpace");

            return new Hashtable
            {
                {"enFont", cmbEngFont?.Text ?? "仿宋"},
                {"cnFont", cmbChnFont?.Text ?? "仿宋"},
                {"fontSize", float.TryParse(cmbFontSize?.Text, out float fontSize) ? fontSize : 16f},
                {"alignment", GetAlignmentValue()},
                {"spaceBefore", GetSpaceValue(txtSpaceBefore?.Text)},
                {"spaceAfter", GetSpaceValue(txtSpaceAfter?.Text)},
                {"lineSpacing", GetSpaceValue(cmbLineSpace?.Text)}
            };
        }

        // 获取间距值（去除"磅"和"行"单位）
        private float GetSpaceValue(string spaceText)
        {
            if (string.IsNullOrEmpty(spaceText)) return 0f;
            var cleanText = spaceText.Replace("磅", "").Replace("行", "").Trim();
            return float.TryParse(cleanText, out float value) ? value : 0f;
        }

        // 获取首行缩进值（字符转换为磅）
        private float GetFirstLineIndentValue(string indentText)
        {
            if (string.IsNullOrEmpty(indentText)) return 0f;
            
            if (indentText.Contains("字符"))
            {
                var cleanText = indentText.Replace("字符", "").Trim();
                if (float.TryParse(cleanText, out float charValue))
                {
                    // 1字符约等于16磅（仿宋三号字）
                    return charValue * 16f;
                }
            }
            else if (indentText.Contains("磅"))
            {
                var cleanText = indentText.Replace("磅", "").Trim();
                if (float.TryParse(cleanText, out float pointValue))
                {
                    return pointValue;
                }
            }
            
            return 0f;
        }

        // 获取对齐方式对应的枚举值
        private WdParagraphAlignment GetAlignmentValue()
        {
            var cmbAlign = FindControl<ComboBox>("cmbAlignment");
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
            try
            {
                // 获取当前Word应用程序和文档
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("请先打开一个Word文档！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 应用页面设置（固定为GB/T 9704-2012标准）
                ApplyPageSettings(doc, app);

                // 应用样式设置
                ApplyStyleToDocument("标题 1", Title1Settings, doc);
                ApplyStyleToDocument("标题 2", Title2Settings, doc);
                ApplyStyleToDocument("标题 3", Title3Settings, doc);
                ApplyStyleToDocument("标题 4", Title4Settings, doc);
                ApplyStyleToDocument("标题 5", Title5Settings, doc);
                ApplyStyleToDocument("标题 6", Title6Settings, doc);
                ApplyStyleToDocument("正文", BodyTextSettings, doc);
                
                // 应用自定义样式
                ApplyCustomStyleToDocument("表中文本", TableTextSettings, doc);
                ApplyCustomStyleToDocument("题注", CaptionSettings, doc);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 应用页面设置
        private void ApplyPageSettings(Document doc, Microsoft.Office.Interop.Word.Application app)
        {
            try
            {
                // 应用GB/T 9704-2012标准页面设置
                doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4; // A4纸张
                doc.PageSetup.Orientation = WdOrientation.wdOrientPortrait; // 竖向
                doc.PageSetup.TopMargin = app.InchesToPoints(3.7f / 2.54f); // 上边距37mm
                doc.PageSetup.BottomMargin = app.InchesToPoints(3.5f / 2.54f); // 下边距35mm
                doc.PageSetup.LeftMargin = app.InchesToPoints(2.8f / 2.54f); // 左边距28mm
                doc.PageSetup.RightMargin = app.InchesToPoints(2.6f / 2.54f); // 右边距26mm
                doc.PageSetup.Gutter = 0; // 无装订线
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用页面设置时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 应用单个样式到文档
        private void ApplyStyleToDocument(string styleName, Hashtable settings, Document doc)
        {
            if (settings == null) return;

            try
            {
                var style = doc.Styles[styleName];

                // 设置字体
                style.Font.Name = settings["cnFont"]?.ToString() ?? "仿宋";
                style.Font.NameAscii = settings["enFont"]?.ToString() ?? "仿宋";
                style.Font.Size = float.TryParse(settings["fontSize"]?.ToString(), out float fontSize) ? fontSize : 16f;
                style.Font.Bold = Convert.ToBoolean(settings["isBold"]) ? 1 : 0;

                // 设置段落格式
                style.ParagraphFormat.Alignment = (WdParagraphAlignment)(settings["alignment"] ?? WdParagraphAlignment.wdAlignParagraphLeft);
                style.ParagraphFormat.SpaceBefore = float.TryParse(settings["spaceBefore"]?.ToString(), out float spaceBefore) ? spaceBefore : 0f;
                style.ParagraphFormat.SpaceAfter = float.TryParse(settings["spaceAfter"]?.ToString(), out float spaceAfter) ? spaceAfter : 0f;
                style.ParagraphFormat.LineSpacing = float.TryParse(settings["lineSpacing"]?.ToString(), out float lineSpacing) ? lineSpacing : 28f;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式「{styleName}」时出错：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 应用自定义样式到文档
        private void ApplyCustomStyleToDocument(string styleName, Hashtable settings, Document doc)
        {
            if (settings == null) return;

            try
            {
                // 检查样式是否存在，如果不存在则创建
                Style style;
                try
                {
                    style = doc.Styles[styleName];
                }
                catch
                {
                    // 样式不存在，创建新样式
                    style = doc.Styles.Add(styleName, WdStyleType.wdStyleTypeParagraph);
                }

                // 设置字体
                style.Font.Name = settings["cnFont"]?.ToString() ?? "仿宋";
                style.Font.NameAscii = settings["enFont"]?.ToString() ?? "仿宋";
                style.Font.Size = float.TryParse(settings["fontSize"]?.ToString(), out float fontSize) ? fontSize : 16f;
                style.Font.Bold = Convert.ToBoolean(settings["isBold"]) ? 1 : 0;

                // 设置段落格式
                style.ParagraphFormat.Alignment = (WdParagraphAlignment)(settings["alignment"] ?? WdParagraphAlignment.wdAlignParagraphLeft);
                style.ParagraphFormat.SpaceBefore = float.TryParse(settings["spaceBefore"]?.ToString(), out float spaceBefore) ? spaceBefore : 0f;
                style.ParagraphFormat.SpaceAfter = float.TryParse(settings["spaceAfter"]?.ToString(), out float spaceAfter) ? spaceAfter : 0f;
                style.ParagraphFormat.LineSpacing = float.TryParse(settings["lineSpacing"]?.ToString(), out float lineSpacing) ? lineSpacing : 28f;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用自定义样式「{styleName}」时出错：{ex.Message}", "错误",
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
                case "标题 1":
                    LoadTitleSettings("Title1", Title1Settings);
                    break;
                case "标题 2":
                    LoadTitleSettings("Title2", Title2Settings);
                    break;
                case "标题 3":
                    LoadTitleSettings("Title3", Title3Settings);
                    break;
                case "标题 4":
                    LoadTitleSettings("Title4", Title4Settings);
                    break;
                case "标题 5":
                    LoadTitleSettings("Title5", Title5Settings);
                    break;
                case "标题 6":
                    LoadTitleSettings("Title6", Title6Settings);
                    break;
                case "正文":
                    LoadBodyTextSettings(false);
                    break;
                case "表内文字":
                    LoadTableTextSettings();
                    break;
                case "题注":
                    LoadCaptionSettings();
                    break;
            }
            
            // 更新预览
            _uiDesigner.UpdateStylePreview();
        }

        /// <summary>
        /// 加载样式按钮点击事件
        /// </summary>
        private void LoadStyle_Click(object sender, EventArgs e)
        {
            try
            {
                // 显示打开文件对话框
                using (var openDialog = new OpenFileDialog())
                {
                    openDialog.Filter = "样式文件|*.xml|所有文件|*.*";
                    openDialog.Title = "加载样式设置";
                    openDialog.DefaultExt = "xml";
                    openDialog.CheckFileExists = true;
                    openDialog.Multiselect = false;

                    if (openDialog.ShowDialog() == DialogResult.OK)
                    {
                        // 从XML文件加载样式
                        var styleInfos = StyleSerializationHelper.DeserializeListFromXml<StyleInfo>(openDialog.FileName);
                        
                        // 应用加载的样式设置
                        foreach (var styleInfo in styleInfos)
                        {
                            var settings = styleInfo.ToHashtable();
                            
                            switch (styleInfo.StyleName)
                            {
                                case "正文":
                                    BodyTextSettings = settings;
                                    break;
                                case "标题 1":
                                    Title1Settings = settings;
                                    break;
                                case "标题 2":
                                    Title2Settings = settings;
                                    break;
                                case "标题 3":
                                    Title3Settings = settings;
                                    break;
                                case "标题 4":
                                    Title4Settings = settings;
                                    break;
                                case "标题 5":
                                    Title5Settings = settings;
                                    break;
                                case "标题 6":
                                    Title6Settings = settings;
                                    break;
                                case "表中文本":
                                    TableTextSettings = settings;
                                    break;
                                case "题注":
                                    CaptionSettings = settings;
                                    break;
                            }
                        }

                        // 刷新界面
                        LoadSettings();
                        MessageBox.Show($"已从文件加载样式设置：{openDialog.FileName}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
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

                // 创建样式信息列表
                var styleInfos = new List<StyleInfo>
                {
                    StyleInfo.FromHashtable("正文", BodyTextSettings),
                    StyleInfo.FromHashtable("标题 1", Title1Settings),
                    StyleInfo.FromHashtable("标题 2", Title2Settings),
                    StyleInfo.FromHashtable("标题 3", Title3Settings),
                    StyleInfo.FromHashtable("标题 4", Title4Settings),
                    StyleInfo.FromHashtable("标题 5", Title5Settings),
                    StyleInfo.FromHashtable("标题 6", Title6Settings),
                    StyleInfo.FromHashtable("表中文本", TableTextSettings),
                    StyleInfo.FromHashtable("题注", CaptionSettings)
                };

                // 显示保存文件对话框
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "样式文件|*.xml";
                    saveDialog.Title = "保存样式设置";
                    saveDialog.DefaultExt = "xml";
                    saveDialog.AddExtension = true;
                    saveDialog.FileName = $"样式设置_{DateTime.Now:yyyyMMdd_HHmmss}.xml";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        StyleSerializationHelper.SerializeListToXml(styleInfos, saveDialog.FileName);
                        MessageBox.Show($"样式设置已保存到：{saveDialog.FileName}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
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
                var txtNewStyleName = _uiDesigner.GetControl<TextBox>("txtNewStyleName");
                if (txtNewStyleName == null) return;
                
                string newStyleName = txtNewStyleName.Text.Trim();
                
                // 检查是否为提示文本
                if (newStyleName == "输入添加样式的名称" || string.IsNullOrEmpty(newStyleName))
                {
                    MessageBox.Show("请输入样式名称！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

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

                // 清空输入框并恢复提示文本
                txtNewStyleName.Text = "输入添加样式的名称";
                txtNewStyleName.ForeColor = Color.Gray;

                MessageBox.Show($"样式「{newStyleName}」已添加到列表！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                var result = MessageBox.Show($"确定要从列表中删除样式「{selectedStyle}」吗？", "确认删除", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    // 从列表中移除
                    _styleNames.Remove(selectedStyle);
                    styleList.Items.Remove(selectedStyle);
                    
                    MessageBox.Show($"样式「{selectedStyle}」已从列表中删除！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                // 获取所有内置样式
                var allBuiltInStyles = new List<string>();
                var currentStyles = new List<string>();
                
                foreach (Style style in doc.Styles)
                {
                    if (style.BuiltIn)
                    {
                        allBuiltInStyles.Add(style.NameLocal);
                        // 检查是否已经在当前样式列表中
                        if (_styleNames.Contains(style.NameLocal))
                        {
                            currentStyles.Add(style.NameLocal);
                        }
                    }
                }

                if (allBuiltInStyles.Count == 0)
                {
                    MessageBox.Show("未找到内置样式！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 显示内置样式选择对话框
                var selectedStyles = ShowBuiltInStyleSelector(allBuiltInStyles, currentStyles);
                if (selectedStyles != null && selectedStyles.Count > 0)
                {
                    // 更新样式列表
                    UpdateStyleListFromBuiltIn(selectedStyles);
                    MessageBox.Show($"已添加 {selectedStyles.Count} 个内置样式到列表！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择内置样式时出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 显示内置样式选择对话框
        /// </summary>
        private List<string> ShowBuiltInStyleSelector(List<string> allStyles, List<string> currentStyles)
        {
            var form = new Form
            {
                Text = "选择内置样式",
                Size = new Size(400, 500),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var checkedListBox = new CheckedListBox
            {
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(360, 400),
                CheckOnClick = true
            };

            // 添加所有样式到列表
            foreach (var style in allStyles)
            {
                checkedListBox.Items.Add(style, currentStyles.Contains(style));
            }

            var btnOK = new Button
            {
                Text = "确定",
                Location = new System.Drawing.Point(200, 420),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new System.Drawing.Point(290, 420),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel
            };

            form.Controls.Add(checkedListBox);
            form.Controls.Add(btnOK);
            form.Controls.Add(btnCancel);

            if (form.ShowDialog() == DialogResult.OK)
            {
                var selectedStyles = new List<string>();
                foreach (string item in checkedListBox.CheckedItems)
                {
                    selectedStyles.Add(item);
                }
                return selectedStyles;
            }

            return null;
        }

        /// <summary>
        /// 从内置样式更新样式列表
        /// </summary>
        private void UpdateStyleListFromBuiltIn(List<string> selectedStyles)
        {
            var styleList = _uiDesigner.GetControl<ListBox>("lstStyleList");
            if (styleList == null) return;

            // 清空当前列表
            styleList.Items.Clear();
            _styleNames.Clear();

            // 添加选中的样式
            foreach (var style in selectedStyles)
            {
                _styleNames.Add(style);
                styleList.Items.Add(style);
            }
        }

        /// <summary>
        /// 行距选择变化事件
        /// </summary>
        private void LineSpace_SelectedIndexChanged(object sender, EventArgs e)
        {
            var cmbLineSpace = sender as ComboBox;
            if (cmbLineSpace == null) return;

            var txtLineSpaceValue = _uiDesigner.GetControl<TextBox>("txtLineSpaceValue");
            if (txtLineSpaceValue == null) return;

            // 根据选择显示或隐藏输入框
            string selectedText = cmbLineSpace.SelectedItem?.ToString();
            if (selectedText == "固定值" || selectedText == "最小值")
            {
                txtLineSpaceValue.Visible = true;
                if (selectedText == "固定值")
                    txtLineSpaceValue.Text = "12磅";
                else
                    txtLineSpaceValue.Text = "12磅";
            }
            else
            {
                txtLineSpaceValue.Visible = false;
            }

            // 触发值变化事件
            ControlValueChanged(sender, e);
        }

        /// <summary>
        /// 行距值验证事件
        /// </summary>
        private void LineSpaceValue_Validated(object sender, EventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox == null) return;

            string text = textBox.Text.TrimEnd(' ', '磅', '行');
            try
            {
                float value = float.Parse(text);
                if (textBox.Text.EndsWith("行"))
                {
                    textBox.Text = value.ToString("0.00 行");
                }
                else
                {
                    textBox.Text = value.ToString("0.00 磅");
                }
            }
            catch
            {
                textBox.Text = "12.00 磅";
            }
        }

        /// <summary>
        /// 段落间距验证事件
        /// </summary>
        private void ParagraphSpace_Validated(object sender, EventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox == null) return;

            string text = textBox.Text.TrimEnd(' ', '行', '磅');
            try
            {
                float value = float.Parse(text);
                string unit = "行";
                if (textBox.Text.EndsWith("磅"))
                {
                    unit = "磅";
                }
                textBox.Text = value.ToString("0.00");
                
                // 更新单位标签
                UpdateUnitLabel(textBox, unit);
            }
            catch
            {
                textBox.Text = "0.00";
                UpdateUnitLabel(textBox, "行");
            }
        }

        /// <summary>
        /// 缩进距离验证事件
        /// </summary>
        private void IndentDistance_Validated(object sender, EventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox == null) return;

            string text = textBox.Text.TrimEnd(' ', '磅', '字', '符', '厘', '米');
            try
            {
                float value = float.Parse(text);
                string unit = "字符";
                if (textBox.Text.EndsWith("厘米"))
                {
                    unit = "厘米";
                }
                else if (textBox.Text.EndsWith("磅"))
                {
                    unit = "磅";
                }
                textBox.Text = value.ToString("0.00");
                
                // 更新单位标签
                UpdateUnitLabel(textBox, unit);
            }
            catch
            {
                textBox.Text = "2.00";
                UpdateUnitLabel(textBox, "字符");
            }
        }

        /// <summary>
        /// 更新单位标签
        /// </summary>
        private void UpdateUnitLabel(TextBox textBox, string unit)
        {
            string textBoxName = textBox.Name;
            string unitLabelName = "";
            
            switch (textBoxName)
            {
                case "txtSpaceBefore":
                    unitLabelName = "lblSpaceBeforeUnit";
                    break;
                case "txtSpaceAfter":
                    unitLabelName = "lblSpaceAfterUnit";
                    break;
                case "txtIndentDistance":
                    unitLabelName = "lblIndentUnit";
                    break;
            }
            
            if (!string.IsNullOrEmpty(unitLabelName))
            {
                var unitLabel = _uiDesigner.GetControl<Label>(unitLabelName);
                if (unitLabel != null)
                {
                    unitLabel.Text = unit;
                }
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

                // 绑定格式复选框
                var chkBold = _uiDesigner.GetControl<CheckBox>("chkBold");
                if (chkBold != null)
                    chkBold.CheckedChanged += ControlValueChanged;

                var chkItalic = _uiDesigner.GetControl<CheckBox>("chkItalic");
                if (chkItalic != null)
                    chkItalic.CheckedChanged += ControlValueChanged;

                var chkUnderline = _uiDesigner.GetControl<CheckBox>("chkUnderline");
                if (chkUnderline != null)
                    chkUnderline.CheckedChanged += ControlValueChanged;

                // 绑定对齐方式控件
                var cmbAlign = _uiDesigner.GetControl<ComboBox>("cmbAlignment");
                if (cmbAlign != null)
                    cmbAlign.SelectedIndexChanged += ControlValueChanged;

                // 绑定间距控件
                var txtSpaceBefore = _uiDesigner.GetControl<TextBox>("txtSpaceBefore");
                if (txtSpaceBefore != null)
                {
                    txtSpaceBefore.TextChanged += ControlValueChanged;
                    txtSpaceBefore.Validated += ParagraphSpace_Validated;
                }

                var txtSpaceAfter = _uiDesigner.GetControl<TextBox>("txtSpaceAfter");
                if (txtSpaceAfter != null)
                {
                    txtSpaceAfter.TextChanged += ControlValueChanged;
                    txtSpaceAfter.Validated += ParagraphSpace_Validated;
                }

                var cmbLineSpace = _uiDesigner.GetControl<ComboBox>("cmbLineSpace");
                if (cmbLineSpace != null)
                    cmbLineSpace.SelectedIndexChanged += LineSpace_SelectedIndexChanged;

                var txtLineSpaceValue = _uiDesigner.GetControl<TextBox>("txtLineSpaceValue");
                if (txtLineSpaceValue != null)
                {
                    txtLineSpaceValue.TextChanged += ControlValueChanged;
                    txtLineSpaceValue.Validated += LineSpaceValue_Validated;
                }

                // 绑定缩进控件
                var txtFirstIndent = _uiDesigner.GetControl<TextBox>("txtFirstIndent");
                if (txtFirstIndent != null)
                    txtFirstIndent.TextChanged += ControlValueChanged;

                var txtIndentDistance = _uiDesigner.GetControl<TextBox>("txtIndentDistance");
                if (txtIndentDistance != null)
                {
                    txtIndentDistance.TextChanged += ControlValueChanged;
                    txtIndentDistance.Validated += IndentDistance_Validated;
                }
                    
                // 绑定段前分页复选框
                var chkPageBreakBefore = _uiDesigner.GetControl<CheckBox>("chkPageBreakBefore");
                if (chkPageBreakBefore != null)
                    chkPageBreakBefore.CheckedChanged += ControlValueChanged;

                // 绑定大纲级别和缩进方式
                var cmbOutlineLevel = _uiDesigner.GetControl<ComboBox>("cmbOutlineLevel");
                if (cmbOutlineLevel != null)
                    cmbOutlineLevel.SelectedIndexChanged += ControlValueChanged;

                var cmbIndentType = _uiDesigner.GetControl<ComboBox>("cmbIndentType");
                if (cmbIndentType != null)
                    cmbIndentType.SelectedIndexChanged += ControlValueChanged;
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
                // 调用UI设计器的预览更新方法
                _uiDesigner?.UpdateStylePreview();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"更新样式预览出错：{ex.Message}");
            }
        }

        #endregion
    }
}