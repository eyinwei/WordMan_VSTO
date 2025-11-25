namespace WordMan
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.WordMan = this.Factory.CreateRibbonTab();
            this.文本处理 = this.Factory.CreateRibbonGroup();
            this.去除断行 = this.Factory.CreateRibbonButton();
            this.去除空格 = this.Factory.CreateRibbonButton();
            this.去除空行 = this.Factory.CreateRibbonButton();
            this.英标转中标 = this.Factory.CreateRibbonButton();
            this.中标转英标 = this.Factory.CreateRibbonButton();
            this.自动加空格 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.缩进2字符 = this.Factory.CreateRibbonButton();
            this.去除缩进 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.希腊字母 = this.Factory.CreateRibbonButton();
            this.常用符号 = this.Factory.CreateRibbonButton();
            this.表格处理 = this.Factory.CreateRibbonGroup();
            this.创建三线表 = this.Factory.CreateRibbonButton();
            this.设为三线表 = this.Factory.CreateRibbonButton();
            this.插入N行 = this.Factory.CreateRibbonButton();
            this.插入N列 = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.表格编号 = this.Factory.CreateRibbonButton();
            this.表注样式1 = this.Factory.CreateRibbonToggleButton();
            this.表注样式2 = this.Factory.CreateRibbonToggleButton();
            this.表注样式3 = this.Factory.CreateRibbonToggleButton();
            this.公式处理 = this.Factory.CreateRibbonGroup();
            this.公式编号 = this.Factory.CreateRibbonButton();
            this.公式样式1 = this.Factory.CreateRibbonToggleButton();
            this.公式样式2 = this.Factory.CreateRibbonToggleButton();
            this.公式样式3 = this.Factory.CreateRibbonToggleButton();
            this.图片处理 = this.Factory.CreateRibbonGroup();
            this.图片编号 = this.Factory.CreateRibbonButton();
            this.图注样式1 = this.Factory.CreateRibbonToggleButton();
            this.图注样式2 = this.Factory.CreateRibbonToggleButton();
            this.图注样式3 = this.Factory.CreateRibbonToggleButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.宽度刷 = this.Factory.CreateRibbonToggleButton();
            this.高度刷 = this.Factory.CreateRibbonToggleButton();
            this.位图化 = this.Factory.CreateRibbonButton();
            this.全文处理 = this.Factory.CreateRibbonGroup();
            this.排版工具 = this.Factory.CreateRibbonButton();
            this.样式设置 = this.Factory.CreateRibbonButton();
            this.多级列表 = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.域名高亮 = this.Factory.CreateRibbonButton();
            this.取消高亮 = this.Factory.CreateRibbonButton();
            this.编号设置 = this.Factory.CreateRibbonMenu();
            this.上标 = this.Factory.CreateRibbonButton();
            this.正常 = this.Factory.CreateRibbonButton();
            this.一键排版 = this.Factory.CreateRibbonButton();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.另存PDF = this.Factory.CreateRibbonSplitButton();
            this.版本 = this.Factory.CreateRibbonButton();
            this.文档操作 = this.Factory.CreateRibbonMenu();
            this.文档合并 = this.Factory.CreateRibbonButton();
            this.文档拆分 = this.Factory.CreateRibbonButton();
            this.快速密级 = this.Factory.CreateRibbonMenu();
            this.公开 = this.Factory.CreateRibbonButton();
            this.内部 = this.Factory.CreateRibbonButton();
            this.移除密级 = this.Factory.CreateRibbonButton();
            this.WordMan.SuspendLayout();
            this.文本处理.SuspendLayout();
            this.表格处理.SuspendLayout();
            this.公式处理.SuspendLayout();
            this.图片处理.SuspendLayout();
            this.全文处理.SuspendLayout();
            this.SuspendLayout();
            // 
            // WordMan
            // 
            this.WordMan.Groups.Add(this.文本处理);
            this.WordMan.Groups.Add(this.表格处理);
            this.WordMan.Groups.Add(this.公式处理);
            this.WordMan.Groups.Add(this.图片处理);
            this.WordMan.Groups.Add(this.全文处理);
            this.WordMan.Label = "WordMan";
            this.WordMan.Name = "WordMan";
            // 
            // 文本处理
            // 
            this.文本处理.Items.Add(this.去除断行);
            this.文本处理.Items.Add(this.去除空格);
            this.文本处理.Items.Add(this.去除空行);
            this.文本处理.Items.Add(this.英标转中标);
            this.文本处理.Items.Add(this.中标转英标);
            this.文本处理.Items.Add(this.自动加空格);
            this.文本处理.Items.Add(this.separator1);
            this.文本处理.Items.Add(this.缩进2字符);
            this.文本处理.Items.Add(this.去除缩进);
            this.文本处理.Items.Add(this.separator2);
            this.文本处理.Items.Add(this.希腊字母);
            this.文本处理.Items.Add(this.常用符号);
            this.文本处理.Label = "文本处理";
            this.文本处理.Name = "文本处理";
            // 
            // 去除断行
            // 
            this.去除断行.Image = global::WordMan.Properties.Resources.GroupTaskOutcomesActions;
            this.去除断行.Label = "去除断行";
            this.去除断行.Name = "去除断行";
            this.去除断行.ScreenTip = "去除多余空行";
            this.去除断行.ShowImage = true;
            this.去除断行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除断行_Click);
            // 
            // 去除空格
            // 
            this.去除空格.Image = global::WordMan.Properties.Resources.Delete;
            this.去除空格.Label = "去除空格";
            this.去除空格.Name = "去除空格";
            this.去除空格.OfficeImageId = "Delete";
            this.去除空格.ScreenTip = "去除多余空格";
            this.去除空格.ShowImage = true;
            this.去除空格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除空格_Click);
            // 
            // 去除空行
            // 
            this.去除空行.Image = global::WordMan.Properties.Resources.DeleteCells;
            this.去除空行.Label = "去除空行";
            this.去除空行.Name = "去除空行";
            this.去除空行.OfficeImageId = "Delete";
            this.去除空行.ScreenTip = "去除段落之间的空行";
            this.去除空行.ShowImage = true;
            this.去除空行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除空行_Click);
            // 
            // 英标转中标
            // 
            this.英标转中标.Image = global::WordMan.Properties.Resources.CommaSign;
            this.英标转中标.Label = "英标转中标";
            this.英标转中标.Name = "英标转中标";
            this.英标转中标.OfficeImageId = "CommaSign";
            this.英标转中标.ScreenTip = "英文标点转中文标点";
            this.英标转中标.ShowImage = true;
            this.英标转中标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.英标转中标_Click);
            // 
            // 中标转英标
            // 
            this.中标转英标.Image = global::WordMan.Properties.Resources.DollarSign;
            this.中标转英标.Label = "中标转英标";
            this.中标转英标.Name = "中标转英标";
            this.中标转英标.OfficeImageId = "DollarSign";
            this.中标转英标.ScreenTip = "中文标点转英文标点";
            this.中标转英标.ShowImage = true;
            this.中标转英标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.中标转英标_Click);
            // 
            // 自动加空格
            // 
            this.自动加空格.Image = global::WordMan.Properties.Resources.GroupTableCellFormat;
            this.自动加空格.Label = "自动加空格";
            this.自动加空格.Name = "自动加空格";
            this.自动加空格.OfficeImageId = "TextAlignLeft";
            this.自动加空格.ScreenTip = "数字和单位间加入空格";
            this.自动加空格.ShowImage = true;
            this.自动加空格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.自动加空格_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // 缩进2字符
            // 
            this.缩进2字符.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.缩进2字符.Image = global::WordMan.Properties.Resources.WordIndent;
            this.缩进2字符.Label = "缩进2字符";
            this.缩进2字符.Name = "缩进2字符";
            this.缩进2字符.OfficeImageId = "TextAlignRight";
            this.缩进2字符.ShowImage = true;
            this.缩进2字符.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.缩进2字符_Click);
            // 
            // 去除缩进
            // 
            this.去除缩进.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.去除缩进.Image = global::WordMan.Properties.Resources.WordOutdent;
            this.去除缩进.Label = "去除缩进";
            this.去除缩进.Name = "去除缩进";
            this.去除缩进.OfficeImageId = "TextAlignLeft";
            this.去除缩进.ShowImage = true;
            this.去除缩进.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除缩进_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // 希腊字母
            // 
            this.希腊字母.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.希腊字母.Image = global::WordMan.Properties.Resources.EquationEdit;
            this.希腊字母.Label = "希腊字母";
            this.希腊字母.Name = "希腊字母";
            this.希腊字母.OfficeImageId = "EquationEdit";
            this.希腊字母.ShowImage = true;
            this.希腊字母.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.希腊字母_Click);
            // 
            // 常用符号
            // 
            this.常用符号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.常用符号.Image = global::WordMan.Properties.Resources.EquationChangeLimitLocation;
            this.常用符号.Label = "常用符号";
            this.常用符号.Name = "常用符号";
            this.常用符号.OfficeImageId = "EquationOperatorGallery";
            this.常用符号.ShowImage = true;
            this.常用符号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.常用符号_Click);
            // 
            // 表格处理
            // 
            this.表格处理.Items.Add(this.创建三线表);
            this.表格处理.Items.Add(this.设为三线表);
            this.表格处理.Items.Add(this.插入N行);
            this.表格处理.Items.Add(this.插入N列);
            this.表格处理.Items.Add(this.separator3);
            this.表格处理.Items.Add(this.表格编号);
            this.表格处理.Items.Add(this.表注样式1);
            this.表格处理.Items.Add(this.表注样式2);
            this.表格处理.Items.Add(this.表注样式3);
            this.表格处理.Label = "表格处理";
            this.表格处理.Name = "表格处理";
            // 
            // 创建三线表
            // 
            this.创建三线表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.创建三线表.Image = global::WordMan.Properties.Resources.TableInsert;
            this.创建三线表.Label = "创建三线表";
            this.创建三线表.Name = "创建三线表";
            this.创建三线表.OfficeImageId = "AccessFormModalDialog";
            this.创建三线表.ShowImage = true;
            this.创建三线表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.创建三线表_Click);
            // 
            // 设为三线表
            // 
            this.设为三线表.Image = global::WordMan.Properties.Resources.TableProperties3;
            this.设为三线表.Label = "设为三线表";
            this.设为三线表.Name = "设为三线表";
            this.设为三线表.OfficeImageId = "TableProperties";
            this.设为三线表.ScreenTip = "设为三线表";
            this.设为三线表.ShowImage = true;
            this.设为三线表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.设为三线表_Click);
            // 
            // 插入N行
            // 
            this.插入N行.Image = global::WordMan.Properties.Resources.TableRowsInsertBelowWord;
            this.插入N行.Label = "插入N行";
            this.插入N行.Name = "插入N行";
            this.插入N行.OfficeImageId = "EquationMatrixInsertRowAfter";
            this.插入N行.ShowImage = true;
            this.插入N行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入N行_Click);
            // 
            // 插入N列
            // 
            this.插入N列.Image = global::WordMan.Properties.Resources.TableColumnsInsertRight;
            this.插入N列.Label = "插入N列";
            this.插入N列.Name = "插入N列";
            this.插入N列.OfficeImageId = "EquationMatrixInsertColumnAfter";
            this.插入N列.ShowImage = true;
            this.插入N列.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入N列_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // 表格编号
            // 
            this.表格编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.表格编号.Image = global::WordMan.Properties.Resources.CaptionProperties2;
            this.表格编号.Label = "表格编号";
            this.表格编号.Name = "表格编号";
            this.表格编号.OfficeImageId = "TableDesign";
            this.表格编号.ScreenTip = "所在表格上方插入表题";
            this.表格编号.ShowImage = true;
            this.表格编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表格编号_Click);
            // 
            // 表注样式1
            // 
            this.表注样式1.Checked = true;
            this.表注样式1.Image = global::WordMan.Properties.Resources.TableAutoFormat;
            this.表注样式1.Label = "表 1  ";
            this.表注样式1.Name = "表注样式1";
            this.表注样式1.OfficeImageId = "AdpDiagramNewTable";
            this.表注样式1.ShowImage = true;
            this.表注样式1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表注样式1_Click);
            // 
            // 表注样式2
            // 
            this.表注样式2.Image = global::WordMan.Properties.Resources.TableAutoFormat;
            this.表注样式2.Label = "表 1-1";
            this.表注样式2.Name = "表注样式2";
            this.表注样式2.OfficeImageId = "AdpDiagramNewTable";
            this.表注样式2.ShowImage = true;
            this.表注样式2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表注样式2_Click);
            // 
            // 表注样式3
            // 
            this.表注样式3.Image = global::WordMan.Properties.Resources.TableAutoFormat;
            this.表注样式3.Label = "表 1.1";
            this.表注样式3.Name = "表注样式3";
            this.表注样式3.OfficeImageId = "AdpDiagramNewTable";
            this.表注样式3.ShowImage = true;
            this.表注样式3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表注样式3_Click);
            // 
            // 公式处理
            // 
            this.公式处理.Items.Add(this.公式编号);
            this.公式处理.Items.Add(this.公式样式1);
            this.公式处理.Items.Add(this.公式样式2);
            this.公式处理.Items.Add(this.公式样式3);
            this.公式处理.Label = "公式处理";
            this.公式处理.Name = "公式处理";
            // 
            // 公式编号
            // 
            this.公式编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.公式编号.Image = global::WordMan.Properties.Resources.Formula;
            this.公式编号.Label = "公式编号";
            this.公式编号.Name = "公式编号";
            this.公式编号.OfficeImageId = "FormulaEvaluate";
            this.公式编号.ScreenTip = "公式所在行进行编号(表格法)";
            this.公式编号.ShowImage = true;
            this.公式编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式编号_Click);
            // 
            // 公式样式1
            // 
            this.公式样式1.Checked = true;
            this.公式样式1.Image = global::WordMan.Properties.Resources.FormulaEvaluate;
            this.公式样式1.Label = "（ 1 ）";
            this.公式样式1.Name = "公式样式1";
            this.公式样式1.OfficeImageId = "Numbering";
            this.公式样式1.ShowImage = true;
            this.公式样式1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式样式1_Click);
            // 
            // 公式样式2
            // 
            this.公式样式2.Image = global::WordMan.Properties.Resources.FormulaEvaluate;
            this.公式样式2.Label = "（1-1）";
            this.公式样式2.Name = "公式样式2";
            this.公式样式2.OfficeImageId = "Numbering";
            this.公式样式2.ScreenTip = "第一个数字来源于一级标题编号";
            this.公式样式2.ShowImage = true;
            this.公式样式2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式样式2_Click);
            // 
            // 公式样式3
            // 
            this.公式样式3.Image = global::WordMan.Properties.Resources.FormulaEvaluate;
            this.公式样式3.Label = "（1.1）";
            this.公式样式3.Name = "公式样式3";
            this.公式样式3.OfficeImageId = "Numbering";
            this.公式样式3.ScreenTip = "第一个数字来源于一级标题编号";
            this.公式样式3.ShowImage = true;
            this.公式样式3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式样式3_Click);
            // 
            // 图片处理
            // 
            this.图片处理.Items.Add(this.图片编号);
            this.图片处理.Items.Add(this.图注样式1);
            this.图片处理.Items.Add(this.图注样式2);
            this.图片处理.Items.Add(this.图注样式3);
            this.图片处理.Items.Add(this.separator4);
            this.图片处理.Items.Add(this.宽度刷);
            this.图片处理.Items.Add(this.高度刷);
            this.图片处理.Items.Add(this.位图化);
            this.图片处理.Label = "图片处理";
            this.图片处理.Name = "图片处理";
            // 
            // 图片编号
            // 
            this.图片编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.图片编号.Image = global::WordMan.Properties.Resources.PictureStylesGallery;
            this.图片编号.Label = "图片编号";
            this.图片编号.Name = "图片编号";
            this.图片编号.OfficeImageId = "ContentControlPicture";
            this.图片编号.ScreenTip = "图片下方插入图题";
            this.图片编号.ShowImage = true;
            this.图片编号.SuperTip = "选中图标或将光标放于图片后";
            this.图片编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图片编号_Click);
            // 
            // 图注样式1
            // 
            this.图注样式1.Checked = true;
            this.图注样式1.Image = global::WordMan.Properties.Resources.GroupOrganizationChartStyleClassic;
            this.图注样式1.Label = "图 1  ";
            this.图注样式1.Name = "图注样式1";
            this.图注样式1.OfficeImageId = "GroupOrganizationChartStyleClassic";
            this.图注样式1.ShowImage = true;
            this.图注样式1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图注样式1_Click);
            // 
            // 图注样式2
            // 
            this.图注样式2.Image = global::WordMan.Properties.Resources.GroupOrganizationChartStyleClassic;
            this.图注样式2.Label = "图 1-1";
            this.图注样式2.Name = "图注样式2";
            this.图注样式2.OfficeImageId = "GroupOrganizationChartStyleClassic";
            this.图注样式2.ScreenTip = "第一个数字来源于一级标题编号";
            this.图注样式2.ShowImage = true;
            this.图注样式2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图注样式2_Click);
            // 
            // 图注样式3
            // 
            this.图注样式3.Image = global::WordMan.Properties.Resources.GroupOrganizationChartStyleClassic;
            this.图注样式3.Label = "图 1.1";
            this.图注样式3.Name = "图注样式3";
            this.图注样式3.OfficeImageId = "GroupOrganizationChartStyleClassic";
            this.图注样式3.ScreenTip = "第一个数字来源于一级标题编号";
            this.图注样式3.ShowImage = true;
            this.图注样式3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图注样式3_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // 宽度刷
            // 
            this.宽度刷.Image = global::WordMan.Properties.Resources.PictureReflectionGalleryItem;
            this.宽度刷.Label = "宽度刷";
            this.宽度刷.Name = "宽度刷";
            this.宽度刷.OfficeImageId = "FormatPainter";
            this.宽度刷.ShowImage = true;
            this.宽度刷.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.宽度刷_Click);
            // 
            // 高度刷
            // 
            this.高度刷.Image = global::WordMan.Properties.Resources.PictureColorTempertatureGallery;
            this.高度刷.Label = "高度刷";
            this.高度刷.Name = "高度刷";
            this.高度刷.OfficeImageId = "FormatPainter";
            this.高度刷.ShowImage = true;
            this.高度刷.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.高度刷_Click);
            // 
            // 位图化
            // 
            this.位图化.Image = global::WordMan.Properties.Resources.PasteAsPicture;
            this.位图化.Label = "位图化";
            this.位图化.Name = "位图化";
            this.位图化.OfficeImageId = "PasteAsPicture";
            this.位图化.ShowImage = true;
            this.位图化.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.位图化_Click);
            // 
            // 全文处理
            // 
            this.全文处理.Items.Add(this.排版工具);
            this.全文处理.Items.Add(this.样式设置);
            this.全文处理.Items.Add(this.多级列表);
            this.全文处理.Items.Add(this.separator5);
            this.全文处理.Items.Add(this.域名高亮);
            this.全文处理.Items.Add(this.取消高亮);
            this.全文处理.Items.Add(this.编号设置);
            this.全文处理.Items.Add(this.一键排版);
            this.全文处理.Items.Add(this.separator6);
            this.全文处理.Items.Add(this.另存PDF);
            this.全文处理.Items.Add(this.文档操作);
            this.全文处理.Items.Add(this.快速密级);
            this.全文处理.Label = "全文处理";
            this.全文处理.Name = "全文处理";
            // 
            // 排版工具
            // 
            this.排版工具.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.排版工具.Image = global::WordMan.Properties.Resources.FormatPainter;
            this.排版工具.Label = "排版工具";
            this.排版工具.Name = "排版工具";
            this.排版工具.OfficeImageId = "FormatPainter";
            this.排版工具.ShowImage = true;
            this.排版工具.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TypesettingButton_Click);
            // 
            // 样式设置
            // 
            this.样式设置.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.样式设置.Image = global::WordMan.Properties.Resources.GroupPensOneNote;
            this.样式设置.Label = "样式设置";
            this.样式设置.Name = "样式设置";
            this.样式设置.OfficeImageId = "CaptionInsert";
            this.样式设置.ShowImage = true;
            this.样式设置.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.样式设置_Click);
            // 
            // 多级列表
            // 
            this.多级列表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.多级列表.Image = global::WordMan.Properties.Resources.NumberedListControl;
            this.多级列表.Label = "多级列表";
            this.多级列表.Name = "多级列表";
            this.多级列表.OfficeImageId = "Numbering";
            this.多级列表.ShowImage = true;
            this.多级列表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.多级列表_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // 域名高亮
            // 
            this.域名高亮.Image = global::WordMan.Properties.Resources.FormatBar;
            this.域名高亮.Label = "域名高亮";
            this.域名高亮.Name = "域名高亮";
            this.域名高亮.OfficeImageId = "TextHighlightColorPicker";
            this.域名高亮.ShowImage = true;
            this.域名高亮.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.域名高亮_Click);
            // 
            // 取消高亮
            // 
            this.取消高亮.Image = global::WordMan.Properties.Resources.FormatComment;
            this.取消高亮.Label = "取消高亮";
            this.取消高亮.Name = "取消高亮";
            this.取消高亮.OfficeImageId = "FormatPlaceholder";
            this.取消高亮.ShowImage = true;
            this.取消高亮.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.取消高亮_Click);
            // 
            // 编号设置
            // 
            this.编号设置.Image = global::WordMan.Properties.Resources.ControlWizards;
            this.编号设置.Items.Add(this.上标);
            this.编号设置.Items.Add(this.正常);
            this.编号设置.Label = "文献编号";
            this.编号设置.Name = "编号设置";
            this.编号设置.OfficeImageId = "ControlWizards";
            this.编号设置.ShowImage = true;
            // 
            // 上标
            // 
            this.上标.Label = "上标";
            this.上标.Name = "上标";
            this.上标.ShowImage = true;
            this.上标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.上标_Click);
            // 
            // 正常
            // 
            this.正常.Label = "正常";
            this.正常.Name = "正常";
            this.正常.ShowImage = true;
            this.正常.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.正常_Click);
            // 
            // 一键排版
            // 
            this.一键排版.Label = "";
            this.一键排版.Name = "一键排版";
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // 另存PDF
            // 
            this.另存PDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.另存PDF.Image = global::WordMan.Properties.Resources.SaveAsPub;
            this.另存PDF.Items.Add(this.版本);
            this.另存PDF.Label = "另存PDF";
            this.另存PDF.Name = "另存PDF";
            this.另存PDF.OfficeImageId = "FileSaveAs";
            this.另存PDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.另存PDF_Click);
            // 
            // 版本
            // 
            this.版本.Label = "版本V2.3";
            this.版本.Name = "版本";
            this.版本.OfficeImageId = "Info";
            this.版本.ShowImage = true;
            this.版本.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.版本_Click);
            // 
            // 文档操作
            // 
            this.文档操作.Items.Add(this.文档合并);
            this.文档操作.Items.Add(this.文档拆分);
            this.文档操作.Label = "文档操作";
            this.文档操作.Name = "文档操作";
            // 
            // 文档合并
            // 
            this.文档合并.Label = "文档合并";
            this.文档合并.Name = "文档合并";
            this.文档合并.ShowImage = true;
            this.文档合并.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文档合并_Click);
            // 
            // 文档拆分
            // 
            this.文档拆分.Label = "文档拆分";
            this.文档拆分.Name = "文档拆分";
            this.文档拆分.ShowImage = true;
            this.文档拆分.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文档拆分_Click);
            // 
            // 快速密级
            // 
            this.快速密级.Items.Add(this.公开);
            this.快速密级.Items.Add(this.内部);
            this.快速密级.Items.Add(this.移除密级);
            this.快速密级.Label = "快速密级";
            this.快速密级.Name = "快速密级";
            // 
            // 公开
            // 
            this.公开.Label = "公开";
            this.公开.Name = "公开";
            this.公开.ShowImage = true;
            this.公开.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公开_Click);
            // 
            // 内部
            // 
            this.内部.Label = "内部";
            this.内部.Name = "内部";
            this.内部.ShowImage = true;
            this.内部.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.内部_Click);
            // 
            // 移除密级
            // 
            this.移除密级.Label = "移除";
            this.移除密级.Name = "移除密级";
            this.移除密级.ShowImage = true;
            this.移除密级.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.移除密级_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.WordMan);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.WordMan.ResumeLayout(false);
            this.WordMan.PerformLayout();
            this.文本处理.ResumeLayout(false);
            this.文本处理.PerformLayout();
            this.表格处理.ResumeLayout(false);
            this.表格处理.PerformLayout();
            this.公式处理.ResumeLayout(false);
            this.公式处理.PerformLayout();
            this.图片处理.ResumeLayout(false);
            this.图片处理.PerformLayout();
            this.全文处理.ResumeLayout(false);
            this.全文处理.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab WordMan;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 文本处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除断行;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除空格;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除空行;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除缩进;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 缩进2字符;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 希腊字母;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 常用符号;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 全文处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 另存PDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 公式处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 公式编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 中标转英标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 英标转中标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 自动加空格;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 公式样式2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 公式样式3;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 公式样式1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 表格处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 设为三线表;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 插入N行;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 插入N列;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 域名高亮;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 取消高亮;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 版本;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 图片编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 图注样式1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 图注样式2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 图注样式3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 表格编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 表注样式1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 表注样式2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 表注样式3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 一键排版;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 宽度刷;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 高度刷;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 创建三线表;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 排版工具;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 位图化;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 文档操作;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 文档合并;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 文档拆分;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 快速密级;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 公开;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 内部;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 移除密级;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 多级列表;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 样式设置;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 编号设置;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 上标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 正常;
    }

    partial class ThisRibbonCollection
    {

    }
}
