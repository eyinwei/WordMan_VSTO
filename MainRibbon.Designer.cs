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
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.去除缩进 = this.Factory.CreateRibbonButton();
            this.缩进2字符 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.希腊字母 = this.Factory.CreateRibbonButton();
            this.常用符号 = this.Factory.CreateRibbonButton();
            this.其他 = this.Factory.CreateRibbonGroup();
            this.版本 = this.Factory.CreateRibbonLabel();
            this.WordMan.SuspendLayout();
            this.文本处理.SuspendLayout();
            this.其他.SuspendLayout();
            this.SuspendLayout();
            // 
            // WordMan
            // 
            this.WordMan.Groups.Add(this.文本处理);
            this.WordMan.Groups.Add(this.其他);
            this.WordMan.Label = "WordMan";
            this.WordMan.Name = "WordMan";
            // 
            // 文本处理
            // 
            this.文本处理.Items.Add(this.去除断行);
            this.文本处理.Items.Add(this.去除空格);
            this.文本处理.Items.Add(this.去除空行);
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
            // 其他
            // 
            this.其他.Items.Add(this.版本);
            this.其他.Label = "其他";
            this.其他.Name = "其他";
            // 
            // 版本
            // 
            this.版本.Label = "版本 V0.1";
            this.版本.Name = "版本";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel 版本;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
