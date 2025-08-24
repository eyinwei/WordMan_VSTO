namespace WordMan_VSTO
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
            this.去除缩进 = this.Factory.CreateRibbonButton();
            this.缩进2字符 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.希腊字母 = this.Factory.CreateRibbonButton();
            this.常用符号 = this.Factory.CreateRibbonButton();
            this.表格处理 = this.Factory.CreateRibbonGroup();
            this.三线表 = this.Factory.CreateRibbonButton();
            this.插入N行 = this.Factory.CreateRibbonButton();
            this.插入N列 = this.Factory.CreateRibbonButton();
            this.公式编号 = this.Factory.CreateRibbonGroup();
            this.编号 = this.Factory.CreateRibbonButton();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.toggleButton2 = this.Factory.CreateRibbonToggleButton();
            this.toggleButton3 = this.Factory.CreateRibbonToggleButton();
            this.图片处理 = this.Factory.CreateRibbonGroup();
            this.其他 = this.Factory.CreateRibbonGroup();
            this.域名高亮 = this.Factory.CreateRibbonButton();
            this.取消高亮 = this.Factory.CreateRibbonButton();
            this.另存PDF = this.Factory.CreateRibbonButton();
            this.版本 = this.Factory.CreateRibbonButton();
            this.WordMan.SuspendLayout();
            this.文本处理.SuspendLayout();
            this.表格处理.SuspendLayout();
            this.公式编号.SuspendLayout();
            this.其他.SuspendLayout();
            this.SuspendLayout();
            // 
            // WordMan
            // 
            this.WordMan.Groups.Add(this.文本处理);
            this.WordMan.Groups.Add(this.表格处理);
            this.WordMan.Groups.Add(this.公式编号);
            this.WordMan.Groups.Add(this.图片处理);
            this.WordMan.Groups.Add(this.其他);
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
            this.文本处理.Items.Add(this.去除缩进);
            this.文本处理.Items.Add(this.缩进2字符);
            this.文本处理.Items.Add(this.separator2);
            this.文本处理.Items.Add(this.希腊字母);
            this.文本处理.Items.Add(this.常用符号);
            this.文本处理.Label = "文本处理";
            this.文本处理.Name = "文本处理";
            // 
            // 去除断行
            // 
            this.去除断行.Label = "去除断行";
            this.去除断行.Name = "去除断行";
            this.去除断行.OfficeImageId = "InsertParagraphHtmlTag";
            this.去除断行.ShowImage = true;
            this.去除断行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除断行_Click);
            // 
            // 去除空格
            // 
            this.去除空格.Label = "去除空格";
            this.去除空格.Name = "去除空格";
            this.去除空格.OfficeImageId = "ContextMenuContactCardOverflowDropdown";
            this.去除空格.ShowImage = true;
            this.去除空格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除空格_Click);
            // 
            // 去除空行
            // 
            this.去除空行.Label = "去除空行";
            this.去除空行.Name = "去除空行";
            this.去除空行.OfficeImageId = "ObjectNudgeUp";
            this.去除空行.ShowImage = true;
            this.去除空行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除空行_Click);
            // 
            // 英标转中标
            // 
            this.英标转中标.Label = "英标转中标";
            this.英标转中标.Name = "英标转中标";
            this.英标转中标.OfficeImageId = "CommaSign";
            this.英标转中标.ShowImage = true;
            this.英标转中标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.英标转中标_Click);
            // 
            // 中标转英标
            // 
            this.中标转英标.Label = "中标转英标";
            this.中标转英标.Name = "中标转英标";
            this.中标转英标.OfficeImageId = "DollarSign";
            this.中标转英标.ShowImage = true;
            this.中标转英标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.中标转英标_Click);
            // 
            // 自动加空格
            // 
            this.自动加空格.Label = "自动加空格";
            this.自动加空格.Name = "自动加空格";
            this.自动加空格.OfficeImageId = "GroupShapeSheetFormulaTracing";
            this.自动加空格.ShowImage = true;
            this.自动加空格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.自动加空格_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // 去除缩进
            // 
            this.去除缩进.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.去除缩进.Label = "去除缩进";
            this.去除缩进.Name = "去除缩进";
            this.去除缩进.OfficeImageId = "WordOutdent";
            this.去除缩进.ShowImage = true;
            this.去除缩进.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除缩进_Click);
            // 
            // 缩进2字符
            // 
            this.缩进2字符.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.缩进2字符.Label = "缩进2字符";
            this.缩进2字符.Name = "缩进2字符";
            this.缩进2字符.OfficeImageId = "WordIndent";
            this.缩进2字符.ShowImage = true;
            this.缩进2字符.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.缩进2字符_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // 希腊字母
            // 
            this.希腊字母.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.希腊字母.Label = "希腊字母";
            this.希腊字母.Name = "希腊字母";
            this.希腊字母.OfficeImageId = "EquationEdit";
            this.希腊字母.ShowImage = true;
            this.希腊字母.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.希腊字母_Click);
            // 
            // 常用符号
            // 
            this.常用符号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.常用符号.Label = "常用符号";
            this.常用符号.Name = "常用符号";
            this.常用符号.OfficeImageId = "EquationOperatorGallery";
            this.常用符号.ShowImage = true;
            this.常用符号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.常用符号_Click);
            // 
            // 表格处理
            // 
            this.表格处理.Items.Add(this.三线表);
            this.表格处理.Items.Add(this.插入N行);
            this.表格处理.Items.Add(this.插入N列);
            this.表格处理.Label = "表格处理";
            this.表格处理.Name = "表格处理";
            // 
            // 三线表
            // 
            this.三线表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.三线表.Label = "三线表";
            this.三线表.Name = "三线表";
            this.三线表.OfficeImageId = "TableCellAlignMiddleCenter";
            this.三线表.ShowImage = true;
            this.三线表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.三线表_Click);
            // 
            // 插入N行
            // 
            this.插入N行.Label = "插入N行";
            this.插入N行.Name = "插入N行";
            this.插入N行.OfficeImageId = "EquationMatrixInsertRowAfter";
            this.插入N行.ShowImage = true;
            this.插入N行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入N行_Click);
            // 
            // 插入N列
            // 
            this.插入N列.Label = "插入N列";
            this.插入N列.Name = "插入N列";
            this.插入N列.OfficeImageId = "EquationMatrixInsertColumnAfter";
            this.插入N列.ShowImage = true;
            this.插入N列.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入N列_Click);
            // 
            // 公式编号
            // 
            this.公式编号.Items.Add(this.编号);
            this.公式编号.Items.Add(this.toggleButton1);
            this.公式编号.Items.Add(this.toggleButton2);
            this.公式编号.Items.Add(this.toggleButton3);
            this.公式编号.Label = "公式编号";
            this.公式编号.Name = "公式编号";
            // 
            // 编号
            // 
            this.编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.编号.Label = "编号";
            this.编号.Name = "编号";
            this.编号.OfficeImageId = "DataTypeCalculatedColumn";
            this.编号.ScreenTip = "公式编号";
            this.编号.ShowImage = true;
            this.编号.SuperTip = "公式编号右对齐（用于单独一条公式）";
            this.编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.编号_Click);
            // 
            // toggleButton1
            // 
            this.toggleButton1.Checked = true;
            this.toggleButton1.Label = "（ 1 ）";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.OfficeImageId = "LineNumbersMenu";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // toggleButton2
            // 
            this.toggleButton2.Label = "（1-1）";
            this.toggleButton2.Name = "toggleButton2";
            this.toggleButton2.OfficeImageId = "LineNumbersMenu";
            this.toggleButton2.ShowImage = true;
            this.toggleButton2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton2_Click);
            // 
            // toggleButton3
            // 
            this.toggleButton3.Label = "（1.1）";
            this.toggleButton3.Name = "toggleButton3";
            this.toggleButton3.OfficeImageId = "LineNumbersMenu";
            this.toggleButton3.ShowImage = true;
            this.toggleButton3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton3_Click);
            // 
            // 图片处理
            // 
            this.图片处理.Label = "图片处理";
            this.图片处理.Name = "图片处理";
            // 
            // 其他
            // 
            this.其他.Items.Add(this.域名高亮);
            this.其他.Items.Add(this.取消高亮);
            this.其他.Items.Add(this.另存PDF);
            this.其他.Items.Add(this.版本);
            this.其他.Label = "其他";
            this.其他.Name = "其他";
            // 
            // 域名高亮
            // 
            this.域名高亮.Label = "域名高亮";
            this.域名高亮.Name = "域名高亮";
            this.域名高亮.OfficeImageId = "SparklineFirstPointMoreColors";
            this.域名高亮.ShowImage = true;
            this.域名高亮.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.域名高亮_Click);
            // 
            // 取消高亮
            // 
            this.取消高亮.Label = "取消高亮";
            this.取消高亮.Name = "取消高亮";
            this.取消高亮.OfficeImageId = "FormatPlaceholder";
            this.取消高亮.ShowImage = true;
            this.取消高亮.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.取消高亮_Click);
            // 
            // 另存PDF
            // 
            this.另存PDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.另存PDF.Label = "另存PDF";
            this.另存PDF.Name = "另存PDF";
            this.另存PDF.OfficeImageId = "P";
            this.另存PDF.ShowImage = true;
            this.另存PDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.另存PDF_Click);
            // 
            // 版本
            // 
            this.版本.Label = "版本 V1.1";
            this.版本.Name = "版本";
            this.版本.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.版本_Click);
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
            this.公式编号.ResumeLayout(false);
            this.公式编号.PerformLayout();
            this.其他.ResumeLayout(false);
            this.其他.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 其他;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 另存PDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 公式编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 中标转英标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 英标转中标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 自动加空格;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton3;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 表格处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 三线表;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 插入N行;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 插入N列;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 域名高亮;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 取消高亮;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 版本;
    }

    partial class ThisRibbonCollection
    {

    }
}
