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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainRibbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            this.WordMan = this.Factory.CreateRibbonTab();
            this.文本处理 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.表格处理 = this.Factory.CreateRibbonGroup();
            this.题注与引用 = this.Factory.CreateRibbonGroup();
            this.图片处理 = this.Factory.CreateRibbonGroup();
            this.全文处理 = this.Factory.CreateRibbonGroup();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.清除格式 = this.Factory.CreateRibbonButton();
            this.格式刷 = this.Factory.CreateRibbonToggleButton();
            this.只留文本 = this.Factory.CreateRibbonButton();
            this.去除断行 = this.Factory.CreateRibbonButton();
            this.去除空格 = this.Factory.CreateRibbonButton();
            this.去除空行 = this.Factory.CreateRibbonButton();
            this.英标转中标 = this.Factory.CreateRibbonButton();
            this.中标转英标 = this.Factory.CreateRibbonButton();
            this.自动加空格 = this.Factory.CreateRibbonButton();
            this.缩进2字符 = this.Factory.CreateRibbonButton();
            this.去除缩进 = this.Factory.CreateRibbonButton();
            this.希腊字母 = this.Factory.CreateRibbonButton();
            this.常用符号 = this.Factory.CreateRibbonButton();
            this.字体替换 = this.Factory.CreateRibbonMenu();
            this.仿宋替换 = this.Factory.CreateRibbonButton();
            this.楷体替换 = this.Factory.CreateRibbonButton();
            this.方正小标宋替换 = this.Factory.CreateRibbonButton();
            this.数字替换 = this.Factory.CreateRibbonButton();
            this.创建表格 = this.Factory.CreateRibbonGallery();
            this.设置表格 = this.Factory.CreateRibbonGallery();
            this.插入N行 = this.Factory.CreateRibbonButton();
            this.插入N列 = this.Factory.CreateRibbonButton();
            this.重复标题行 = this.Factory.CreateRibbonToggleButton();
            this.图编号 = this.Factory.CreateRibbonSplitButton();
            this.图注样式1 = this.Factory.CreateRibbonToggleButton();
            this.图注样式2 = this.Factory.CreateRibbonToggleButton();
            this.图注样式3 = this.Factory.CreateRibbonToggleButton();
            this.表编号 = this.Factory.CreateRibbonSplitButton();
            this.表注样式1 = this.Factory.CreateRibbonToggleButton();
            this.表注样式2 = this.Factory.CreateRibbonToggleButton();
            this.表注样式3 = this.Factory.CreateRibbonToggleButton();
            this.式编号 = this.Factory.CreateRibbonSplitButton();
            this.公式样式1 = this.Factory.CreateRibbonToggleButton();
            this.公式样式2 = this.Factory.CreateRibbonToggleButton();
            this.公式样式3 = this.Factory.CreateRibbonToggleButton();
            this.交叉引用 = this.Factory.CreateRibbonToggleButton();
            this.宽度刷 = this.Factory.CreateRibbonToggleButton();
            this.高度刷 = this.Factory.CreateRibbonToggleButton();
            this.位图化 = this.Factory.CreateRibbonButton();
            this.导出图片 = this.Factory.CreateRibbonButton();
            this.排版工具 = this.Factory.CreateRibbonButton();
            this.样式设置 = this.Factory.CreateRibbonButton();
            this.多级列表 = this.Factory.CreateRibbonButton();
            this.域名高亮 = this.Factory.CreateRibbonSplitButton();
            this.取消高亮 = this.Factory.CreateRibbonButton();
            this.另存PDF = this.Factory.CreateRibbonSplitButton();
            this.版本 = this.Factory.CreateRibbonButton();
            this.编号设置 = this.Factory.CreateRibbonMenu();
            this.上标 = this.Factory.CreateRibbonButton();
            this.正常 = this.Factory.CreateRibbonButton();
            this.快速密级 = this.Factory.CreateRibbonMenu();
            this.公开 = this.Factory.CreateRibbonButton();
            this.内部 = this.Factory.CreateRibbonButton();
            this.移除密级 = this.Factory.CreateRibbonButton();
            this.文档操作 = this.Factory.CreateRibbonMenu();
            this.文档合并 = this.Factory.CreateRibbonButton();
            this.文档拆分 = this.Factory.CreateRibbonButton();
            this.字母替换 = this.Factory.CreateRibbonButton();
            this.WordMan.SuspendLayout();
            this.文本处理.SuspendLayout();
            this.表格处理.SuspendLayout();
            this.题注与引用.SuspendLayout();
            this.图片处理.SuspendLayout();
            this.全文处理.SuspendLayout();
            this.SuspendLayout();
            // 
            // WordMan
            // 
            this.WordMan.Groups.Add(this.文本处理);
            this.WordMan.Groups.Add(this.表格处理);
            this.WordMan.Groups.Add(this.题注与引用);
            this.WordMan.Groups.Add(this.图片处理);
            this.WordMan.Groups.Add(this.全文处理);
            this.WordMan.Label = "WordMan";
            this.WordMan.Name = "WordMan";
            // 
            // 文本处理
            // 
            this.文本处理.Items.Add(this.清除格式);
            this.文本处理.Items.Add(this.格式刷);
            this.文本处理.Items.Add(this.只留文本);
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
            this.文本处理.Items.Add(this.separator3);
            this.文本处理.Items.Add(this.字体替换);
            this.文本处理.Label = "文本处理";
            this.文本处理.Name = "文本处理";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // 表格处理
            // 
            this.表格处理.Items.Add(this.创建表格);
            this.表格处理.Items.Add(this.设置表格);
            this.表格处理.Items.Add(this.插入N行);
            this.表格处理.Items.Add(this.插入N列);
            this.表格处理.Items.Add(this.重复标题行);
            this.表格处理.Label = "表格处理";
            this.表格处理.Name = "表格处理";
            // 
            // 题注与引用
            // 
            this.题注与引用.Items.Add(this.图编号);
            this.题注与引用.Items.Add(this.表编号);
            this.题注与引用.Items.Add(this.式编号);
            this.题注与引用.Items.Add(this.交叉引用);
            this.题注与引用.Label = "题注与引用";
            this.题注与引用.Name = "题注与引用";
            // 
            // 图片处理
            // 
            this.图片处理.Items.Add(this.宽度刷);
            this.图片处理.Items.Add(this.高度刷);
            this.图片处理.Items.Add(this.位图化);
            this.图片处理.Items.Add(this.导出图片);
            this.图片处理.Label = "图片处理";
            this.图片处理.Name = "图片处理";
            // 
            // 全文处理
            // 
            this.全文处理.Items.Add(this.排版工具);
            this.全文处理.Items.Add(this.样式设置);
            this.全文处理.Items.Add(this.多级列表);
            this.全文处理.Items.Add(this.域名高亮);
            this.全文处理.Items.Add(this.separator6);
            this.全文处理.Items.Add(this.另存PDF);
            this.全文处理.Items.Add(this.编号设置);
            this.全文处理.Items.Add(this.快速密级);
            this.全文处理.Items.Add(this.文档操作);
            this.全文处理.Label = "全文处理";
            this.全文处理.Name = "全文处理";
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // 清除格式
            // 
            this.清除格式.Label = "清除格式";
            this.清除格式.Name = "清除格式";
            this.清除格式.OfficeImageId = "ClearFormatting";
            this.清除格式.ScreenTip = "清除选中文本的所有格式";
            this.清除格式.ShowImage = true;
            this.清除格式.SuperTip = "清除选中文本或当前段落的所有格式，包括字体、字号、颜色、加粗、斜体等";
            this.清除格式.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.清除格式_Click);
            // 
            // 格式刷
            // 
            this.格式刷.Label = "格式刷";
            this.格式刷.Name = "格式刷";
            this.格式刷.OfficeImageId = "FormatPainter";
            this.格式刷.ScreenTip = "复制格式并应用到其他文本";
            this.格式刷.ShowImage = true;
            this.格式刷.SuperTip = "点击后进入格式刷模式，选择目标文本即可应用格式。再次点击可退出格式刷模式";
            this.格式刷.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.格式刷_Click);
            // 
            // 只留文本
            // 
            this.只留文本.Label = "只留文本";
            this.只留文本.Name = "只留文本";
            this.只留文本.OfficeImageId = "PasteTextOnly";
            this.只留文本.ScreenTip = "清除格式，只保留纯文本";
            this.只留文本.ShowImage = true;
            this.只留文本.SuperTip = "从剪贴板粘贴文本时，自动清除所有格式，只保留纯文本内容";
            this.只留文本.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.只留文本_Click);
            // 
            // 去除断行
            // 
            this.去除断行.Image = ((System.Drawing.Image)(resources.GetObject("去除断行.Image")));
            this.去除断行.Label = "去除断行";
            this.去除断行.Name = "去除断行";
            this.去除断行.ScreenTip = "去除段落内的手动换行符";
            this.去除断行.ShowImage = true;
            this.去除断行.SuperTip = "去除选中文本或当前段落中的手动换行符（Shift+Enter），将其转换为空格";
            this.去除断行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除断行_Click);
            // 
            // 去除空格
            // 
            this.去除空格.Image = ((System.Drawing.Image)(resources.GetObject("去除空格.Image")));
            this.去除空格.Label = "去除空格";
            this.去除空格.Name = "去除空格";
            this.去除空格.OfficeImageId = "Delete";
            this.去除空格.ScreenTip = "去除选中文本中的多余空格";
            this.去除空格.ShowImage = true;
            this.去除空格.SuperTip = "去除选中文本或当前段落中的多余空格，保留每个单词之间的单个空格";
            this.去除空格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除空格_Click);
            // 
            // 去除空行
            // 
            this.去除空行.Image = ((System.Drawing.Image)(resources.GetObject("去除空行.Image")));
            this.去除空行.Label = "去除空行";
            this.去除空行.Name = "去除空行";
            this.去除空行.OfficeImageId = "Delete";
            this.去除空行.ScreenTip = "去除文档中的空白段落";
            this.去除空行.ShowImage = true;
            this.去除空行.SuperTip = "去除选中文本或全文中的空行（空白段落），保留段落之间的正常间距";
            this.去除空行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除空行_Click);
            // 
            // 英标转中标
            // 
            this.英标转中标.Image = ((System.Drawing.Image)(resources.GetObject("英标转中标.Image")));
            this.英标转中标.Label = "英标转中标";
            this.英标转中标.Name = "英标转中标";
            this.英标转中标.OfficeImageId = "CommaSign";
            this.英标转中标.ScreenTip = "英文标点转中文标点";
            this.英标转中标.ShowImage = true;
            this.英标转中标.SuperTip = "将选中文本或当前段落中的英文标点符号转换为对应的中文标点符号";
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
            this.中标转英标.SuperTip = "将选中文本或当前段落中的中文标点符号转换为对应的英文标点符号";
            this.中标转英标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.中标转英标_Click);
            // 
            // 自动加空格
            // 
            this.自动加空格.Image = ((System.Drawing.Image)(resources.GetObject("自动加空格.Image")));
            this.自动加空格.Label = "自动加空格";
            this.自动加空格.Name = "自动加空格";
            this.自动加空格.OfficeImageId = "TextAlignLeft";
            this.自动加空格.ScreenTip = "在数字和单位之间自动添加空格";
            this.自动加空格.ShowImage = true;
            this.自动加空格.SuperTip = "在选中文本或当前段落中，自动在数字和单位（如cm、kg、°C等）之间添加空格，但不处理百分比和角度单位";
            this.自动加空格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.自动加空格_Click);
            // 
            // 缩进2字符
            // 
            this.缩进2字符.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.缩进2字符.Image = global::WordMan.Properties.Resources.增加缩进;
            this.缩进2字符.Label = "缩进2字符";
            this.缩进2字符.Name = "缩进2字符";
            this.缩进2字符.OfficeImageId = "TextAlignRight";
            this.缩进2字符.ScreenTip = "设置段落首行缩进2个字符";
            this.缩进2字符.ShowImage = true;
            this.缩进2字符.SuperTip = "为选中段落或当前段落设置首行缩进2个字符，适用于中文文档的标准段落格式";
            this.缩进2字符.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.缩进2字符_Click);
            // 
            // 去除缩进
            // 
            this.去除缩进.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.去除缩进.Image = global::WordMan.Properties.Resources.减少缩进;
            this.去除缩进.Label = "去除缩进";
            this.去除缩进.Name = "去除缩进";
            this.去除缩进.OfficeImageId = "TextAlignLeft";
            this.去除缩进.ScreenTip = "清除段落的首行缩进和左右缩进";
            this.去除缩进.ShowImage = true;
            this.去除缩进.SuperTip = "清除选中段落或当前段落的首行缩进和左右缩进，恢复为无缩进状态";
            this.去除缩进.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除缩进_Click);
            // 
            // 希腊字母
            // 
            this.希腊字母.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.希腊字母.Image = global::WordMan.Properties.Resources.希腊;
            this.希腊字母.Label = "希腊字母";
            this.希腊字母.Name = "希腊字母";
            this.希腊字母.OfficeImageId = "EquationEdit";
            this.希腊字母.ScreenTip = "插入常用希腊字母";
            this.希腊字母.ShowImage = true;
            this.希腊字母.SuperTip = "打开希腊字母选择窗口，可以快速插入常用的希腊字母符号";
            this.希腊字母.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.希腊字母_Click);
            // 
            // 常用符号
            // 
            this.常用符号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.常用符号.Image = global::WordMan.Properties.Resources.符号;
            this.常用符号.Label = "常用符号";
            this.常用符号.Name = "常用符号";
            this.常用符号.OfficeImageId = "EquationOperatorGallery";
            this.常用符号.ScreenTip = "插入常用特殊符号";
            this.常用符号.ShowImage = true;
            this.常用符号.SuperTip = "打开常用符号选择窗口，可以快速插入数学、物理等常用特殊符号";
            this.常用符号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.常用符号_Click);
            // 
            // 字体替换
            // 
            this.字体替换.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.字体替换.Image = global::WordMan.Properties.Resources.字体替换;
            this.字体替换.Items.Add(this.仿宋替换);
            this.字体替换.Items.Add(this.楷体替换);
            this.字体替换.Items.Add(this.方正小标宋替换);
            this.字体替换.Items.Add(this.数字替换);
            this.字体替换.Label = "字体替换";
            this.字体替换.Name = "字体替换";
            this.字体替换.ScreenTip = "批量替换文档中的字体";
            this.字体替换.ShowImage = true;
            this.字体替换.SuperTip = "批量替换选中文本或全文中的指定字体，支持多种常用字体替换选项";
            // 
            // 仿宋替换
            // 
            this.仿宋替换.Label = "仿宋GB2312→仿宋";
            this.仿宋替换.Name = "仿宋替换";
            this.仿宋替换.ScreenTip = "将选中文本或全文中的仿宋GB2312字体替换为仿宋字体";
            this.仿宋替换.ShowImage = true;
            this.仿宋替换.SuperTip = "批量将选中文本或全文中的仿宋GB2312字体替换为仿宋字体，用于字体标准化";
            this.仿宋替换.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.仿宋替换_Click);
            // 
            // 楷体替换
            // 
            this.楷体替换.Label = "楷体GB2312→楷体";
            this.楷体替换.Name = "楷体替换";
            this.楷体替换.ScreenTip = "将选中文本或全文中的楷体GB2312字体替换为楷体字体";
            this.楷体替换.ShowImage = true;
            this.楷体替换.SuperTip = "批量将选中文本或全文中的楷体GB2312字体替换为楷体字体，用于字体标准化";
            this.楷体替换.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.楷体替换_Click);
            // 
            // 方正小标宋替换
            // 
            this.方正小标宋替换.Label = "方正小标宋→黑体";
            this.方正小标宋替换.Name = "方正小标宋替换";
            this.方正小标宋替换.ScreenTip = "将选中文本或全文中的方正小标宋字体替换为黑体字体";
            this.方正小标宋替换.ShowImage = true;
            this.方正小标宋替换.SuperTip = "批量将选中文本或全文中的方正小标宋字体替换为黑体字体，用于字体标准化";
            this.方正小标宋替换.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.方正小标宋替换_Click);
            // 
            // 数字替换
            // 
            this.数字替换.Label = "数字字母→Times New Roman";
            this.数字替换.Name = "数字替换";
            this.数字替换.ScreenTip = "将选中文本或全文中的数字、英文及希腊字母替换为Times New Roman字体";
            this.数字替换.ShowImage = true;
            this.数字替换.SuperTip = "批量将选中文本或全文中的数字、英文及希腊字母替换为Times New Roman字体，用于统一英文字体";
            this.数字替换.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.数字替换_Click);
            // 
            // 创建表格
            // 
            this.创建表格.ColumnCount = 1;
            this.创建表格.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.创建表格.Image = global::WordMan.Properties.Resources.三线图;
            this.创建表格.ItemImageSize = new System.Drawing.Size(450, 150);
            ribbonDropDownItemImpl1.Image = ((System.Drawing.Image)(resources.GetObject("ribbonDropDownItemImpl1.Image")));
            ribbonDropDownItemImpl1.ScreenTip = "三线表样式";
            ribbonDropDownItemImpl1.SuperTip = "创建表格时使用三线表样式，包含顶线、底线和表头下线";
            ribbonDropDownItemImpl2.Image = global::WordMan.Properties.Resources.国标表示意图;
            ribbonDropDownItemImpl2.ScreenTip = "国标表样式";
            ribbonDropDownItemImpl2.SuperTip = "创建表格时使用国标表格样式，外边框1.5磅，标题栏下边框1.5磅，内部框线0.75磅";
            ribbonDropDownItemImpl3.Image = global::WordMan.Properties.Resources.无线表示意图;
            ribbonDropDownItemImpl3.ScreenTip = "无线表样式";
            ribbonDropDownItemImpl3.SuperTip = "创建2行2列无框线表格，全部居中，关闭尺寸重调，单元格左右边距为0";
            this.创建表格.Items.Add(ribbonDropDownItemImpl1);
            this.创建表格.Items.Add(ribbonDropDownItemImpl2);
            this.创建表格.Items.Add(ribbonDropDownItemImpl3);
            this.创建表格.Label = "创建表格";
            this.创建表格.Name = "创建表格";
            this.创建表格.OfficeImageId = "AccessFormModalDialog";
            this.创建表格.RowCount = 3;
            this.创建表格.ScreenTip = "创建一个新的表格";
            this.创建表格.ShowImage = true;
            this.创建表格.SuperTip = "在当前位置创建一个新的表格，选择样式";
            this.创建表格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.创建表格_Click);
            // 
            // 设置表格
            // 
            this.设置表格.ColumnCount = 1;
            this.设置表格.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.设置表格.Image = global::WordMan.Properties.Resources.表前插入行;
            this.设置表格.ItemImageSize = new System.Drawing.Size(450, 150);
            ribbonDropDownItemImpl4.Image = ((System.Drawing.Image)(resources.GetObject("ribbonDropDownItemImpl4.Image")));
            ribbonDropDownItemImpl4.ScreenTip = "三线表样式";
            ribbonDropDownItemImpl4.SuperTip = "设置表格时使用三线表样式，包含顶线、底线和表头下线";
            ribbonDropDownItemImpl5.Image = global::WordMan.Properties.Resources.国标表示意图;
            ribbonDropDownItemImpl5.ScreenTip = "国标表样式";
            ribbonDropDownItemImpl5.SuperTip = "设置表格时使用国标表格样式，外边框1.5磅，标题栏下边框1.5磅，内部框线0.75磅";
            ribbonDropDownItemImpl6.Image = global::WordMan.Properties.Resources.无线表示意图;
            ribbonDropDownItemImpl6.ScreenTip = "无线表样式";
            ribbonDropDownItemImpl6.SuperTip = "设置表格为无框线样式，全部居中，关闭尺寸重调，单元格左右边距为0";
            this.设置表格.Items.Add(ribbonDropDownItemImpl4);
            this.设置表格.Items.Add(ribbonDropDownItemImpl5);
            this.设置表格.Items.Add(ribbonDropDownItemImpl6);
            this.设置表格.Label = "设置表格";
            this.设置表格.Name = "设置表格";
            this.设置表格.OfficeImageId = "TableProperties";
            this.设置表格.RowCount = 3;
            this.设置表格.ScreenTip = "设置选中表格的格式";
            this.设置表格.ShowImage = true;
            this.设置表格.SuperTip = "将当前选中的表格设置为指定样式";
            this.设置表格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.设置表格_Click);
            // 
            // 插入N行
            // 
            this.插入N行.Image = ((System.Drawing.Image)(resources.GetObject("插入N行.Image")));
            this.插入N行.Label = "插入N行";
            this.插入N行.Name = "插入N行";
            this.插入N行.OfficeImageId = "EquationMatrixInsertRowAfter";
            this.插入N行.ScreenTip = "在表格中插入指定数量的行";
            this.插入N行.ShowImage = true;
            this.插入N行.SuperTip = "在表格中插入指定数量的行，可以选择在上方或下方插入";
            this.插入N行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入N行_Click);
            // 
            // 插入N列
            // 
            this.插入N列.Image = ((System.Drawing.Image)(resources.GetObject("插入N列.Image")));
            this.插入N列.Label = "插入N列";
            this.插入N列.Name = "插入N列";
            this.插入N列.OfficeImageId = "EquationMatrixInsertColumnAfter";
            this.插入N列.ScreenTip = "在表格中插入指定数量的列";
            this.插入N列.ShowImage = true;
            this.插入N列.SuperTip = "在表格中插入指定数量的列，可以选择在左侧或右侧插入。插入后会自动调整列宽以适应页面";
            this.插入N列.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入N列_Click);
            // 
            // 重复标题行
            // 
            this.重复标题行.Label = "重复标题行";
            this.重复标题行.Name = "重复标题行";
            this.重复标题行.OfficeImageId = "TableRepeatHeaderRows";
            this.重复标题行.ScreenTip = "在表格的每一页顶部重复标题行";
            this.重复标题行.ShowImage = true;
            this.重复标题行.SuperTip = "设置表格在跨页时，在每一页的顶部自动重复显示标题行";
            this.重复标题行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.重复标题行_Click);
            // 
            // 图编号
            // 
            this.图编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.图编号.Image = global::WordMan.Properties.Resources.图片;
            this.图编号.Items.Add(this.图注样式1);
            this.图编号.Items.Add(this.图注样式2);
            this.图编号.Items.Add(this.图注样式3);
            this.图编号.Label = "图编号";
            this.图编号.Name = "图编号";
            this.图编号.OfficeImageId = "ContentControlPicture";
            this.图编号.ScreenTip = "在选中图片下方插入带编号的图题";
            this.图编号.SuperTip = "在选中图片下方插入带编号的图题。\n使用方法：选中图片或将光标放于图片后，点击按钮即可插入图编号";
            this.图编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图编号_Click);
            // 
            // 图注样式1
            // 
            this.图注样式1.Checked = true;
            this.图注样式1.Image = global::WordMan.Properties.Resources.照片管理;
            this.图注样式1.Label = "图 1  ";
            this.图注样式1.Name = "图注样式1";
            this.图注样式1.OfficeImageId = "GroupOrganizationChartStyleClassic";
            this.图注样式1.ScreenTip = "设置图编号格式为\'图 1\'样式";
            this.图注样式1.ShowImage = true;
            this.图注样式1.SuperTip = "设置图编号格式为\'图 1\'样式，使用简单的阿拉伯数字编号";
            this.图注样式1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图注样式1_Click);
            // 
            // 图注样式2
            // 
            this.图注样式2.Image = global::WordMan.Properties.Resources.照片管理;
            this.图注样式2.Label = "图 1-1";
            this.图注样式2.Name = "图注样式2";
            this.图注样式2.OfficeImageId = "GroupOrganizationChartStyleClassic";
            this.图注样式2.ScreenTip = "设置图编号格式为\'图 1-1\'样式，第一个数字来源于一级标题编号";
            this.图注样式2.ShowImage = true;
            this.图注样式2.SuperTip = "设置图编号格式为\'图 1-1\'样式，第一个数字来源于一级标题编号，使用连字符分隔";
            this.图注样式2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图注样式2_Click);
            // 
            // 图注样式3
            // 
            this.图注样式3.Image = global::WordMan.Properties.Resources.照片管理;
            this.图注样式3.Label = "图 1.1";
            this.图注样式3.Name = "图注样式3";
            this.图注样式3.OfficeImageId = "GroupOrganizationChartStyleClassic";
            this.图注样式3.ScreenTip = "设置图编号格式为\'图 1.1\'样式，第一个数字来源于一级标题编号";
            this.图注样式3.ShowImage = true;
            this.图注样式3.SuperTip = "设置图编号格式为\'图 1.1\'样式，第一个数字来源于一级标题编号，使用点号分隔";
            this.图注样式3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图注样式3_Click);
            // 
            // 表编号
            // 
            this.表编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.表编号.Image = global::WordMan.Properties.Resources.表前插入行;
            this.表编号.Items.Add(this.表注样式1);
            this.表编号.Items.Add(this.表注样式2);
            this.表编号.Items.Add(this.表注样式3);
            this.表编号.Label = "表编号";
            this.表编号.Name = "表编号";
            this.表编号.OfficeImageId = "TableDesign";
            this.表编号.ScreenTip = "在选中表格上方插入带编号的表题";
            this.表编号.SuperTip = "在选中表格上方插入带编号的表题。\n使用方法：将光标放在表格中，点击按钮即可插入表编号";
            this.表编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表编号_Click);
            // 
            // 表注样式1
            // 
            this.表注样式1.Checked = true;
            this.表注样式1.Image = global::WordMan.Properties.Resources.表编号;
            this.表注样式1.Label = "表 1  ";
            this.表注样式1.Name = "表注样式1";
            this.表注样式1.OfficeImageId = "AdpDiagramNewTable";
            this.表注样式1.ScreenTip = "设置表编号格式为\'表 1\'样式";
            this.表注样式1.ShowImage = true;
            this.表注样式1.SuperTip = "设置表编号格式为\'表 1\'样式，使用简单的阿拉伯数字编号";
            this.表注样式1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表注样式1_Click);
            // 
            // 表注样式2
            // 
            this.表注样式2.Image = global::WordMan.Properties.Resources.表编号;
            this.表注样式2.Label = "表 1-1";
            this.表注样式2.Name = "表注样式2";
            this.表注样式2.OfficeImageId = "AdpDiagramNewTable";
            this.表注样式2.ScreenTip = "设置表编号格式为\'表 1-1\'样式";
            this.表注样式2.ShowImage = true;
            this.表注样式2.SuperTip = "设置表编号格式为\'表 1-1\'样式，第一个数字来源于一级标题编号，使用连字符分隔";
            this.表注样式2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表注样式2_Click);
            // 
            // 表注样式3
            // 
            this.表注样式3.Image = global::WordMan.Properties.Resources.表编号;
            this.表注样式3.Label = "表 1.1";
            this.表注样式3.Name = "表注样式3";
            this.表注样式3.OfficeImageId = "AdpDiagramNewTable";
            this.表注样式3.ScreenTip = "设置表编号格式为\'表 1.1\'样式";
            this.表注样式3.ShowImage = true;
            this.表注样式3.SuperTip = "设置表编号格式为\'表 1.1\'样式，第一个数字来源于一级标题编号，使用点号分隔";
            this.表注样式3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.表注样式3_Click);
            // 
            // 式编号
            // 
            this.式编号.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.式编号.Image = global::WordMan.Properties.Resources.公式;
            this.式编号.Items.Add(this.公式样式1);
            this.式编号.Items.Add(this.公式样式2);
            this.式编号.Items.Add(this.公式样式3);
            this.式编号.Label = "式编号";
            this.式编号.Name = "式编号";
            this.式编号.OfficeImageId = "FormulaEvaluate";
            this.式编号.ScreenTip = "对公式所在行进行编号（使用表格法）";
            this.式编号.SuperTip = "对公式所在行进行编号，使用表格法实现公式居中对齐、编号右对齐的效果。\n使用方法：将光标放在包含公式的段落中，点击按钮即可插入式编号";
            this.式编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.式编号_Click);
            // 
            // 公式样式1
            // 
            this.公式样式1.Checked = true;
            this.公式样式1.Label = "（ 1 ）";
            this.公式样式1.Name = "公式样式1";
            this.公式样式1.OfficeImageId = "Numbering";
            this.公式样式1.ScreenTip = "设置式编号格式为\'(1)\'样式";
            this.公式样式1.ShowImage = true;
            this.公式样式1.SuperTip = "设置式编号格式为\'(1)\'样式，使用简单的阿拉伯数字编号，带括号";
            this.公式样式1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式样式1_Click);
            // 
            // 公式样式2
            // 
            this.公式样式2.Label = "（1-1）";
            this.公式样式2.Name = "公式样式2";
            this.公式样式2.OfficeImageId = "Numbering";
            this.公式样式2.ScreenTip = "设置式编号格式为\'(1-1)\'样式，第一个数字来源于一级标题编号";
            this.公式样式2.ShowImage = true;
            this.公式样式2.SuperTip = "设置式编号格式为\'(1-1)\'样式，第一个数字来源于一级标题编号，使用连字符分隔";
            this.公式样式2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式样式2_Click);
            // 
            // 公式样式3
            // 
            this.公式样式3.Label = "（1.1）";
            this.公式样式3.Name = "公式样式3";
            this.公式样式3.OfficeImageId = "Numbering";
            this.公式样式3.ScreenTip = "设置式编号格式为\'(1.1)\'样式，第一个数字来源于一级标题编号";
            this.公式样式3.ShowImage = true;
            this.公式样式3.SuperTip = "设置式编号格式为\'(1.1)\'样式，第一个数字来源于一级标题编号，使用点号分隔";
            this.公式样式3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公式样式3_Click);
            // 
            // 交叉引用
            // 
            this.交叉引用.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.交叉引用.Image = global::WordMan.Properties.Resources.引用;
            this.交叉引用.Label = "交叉引用";
            this.交叉引用.Name = "交叉引用";
            this.交叉引用.ScreenTip = "插入对图表或公式的交叉引用";
            this.交叉引用.ShowImage = true;
            this.交叉引用.SuperTip = "进入交叉引用模式，可以快速插入对图、表或公式的交叉引用。点击后再次点击图表或公式即可插入引用";
            this.交叉引用.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交叉引用_Click);
            // 
            // 宽度刷
            // 
            this.宽度刷.Image = ((System.Drawing.Image)(resources.GetObject("宽度刷.Image")));
            this.宽度刷.Label = "宽度刷";
            this.宽度刷.Name = "宽度刷";
            this.宽度刷.OfficeImageId = "FormatPainter";
            this.宽度刷.ScreenTip = "统一设置图片宽度";
            this.宽度刷.ShowImage = true;
            this.宽度刷.SuperTip = "进入宽度刷模式，先点击一个图片作为宽度参考，然后点击其他图片即可统一宽度";
            this.宽度刷.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.宽度刷_Click);
            // 
            // 高度刷
            // 
            this.高度刷.Image = ((System.Drawing.Image)(resources.GetObject("高度刷.Image")));
            this.高度刷.Label = "高度刷";
            this.高度刷.Name = "高度刷";
            this.高度刷.OfficeImageId = "FormatPainter";
            this.高度刷.ScreenTip = "统一设置图片高度";
            this.高度刷.ShowImage = true;
            this.高度刷.SuperTip = "进入高度刷模式，先点击一个图片作为高度参考，然后点击其他图片即可统一高度";
            this.高度刷.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.高度刷_Click);
            // 
            // 位图化
            // 
            this.位图化.Image = global::WordMan.Properties.Resources.照片管理;
            this.位图化.Label = "位图化";
            this.位图化.Name = "位图化";
            this.位图化.OfficeImageId = "PasteAsPicture";
            this.位图化.ScreenTip = "将选中图形转换为位图";
            this.位图化.ShowImage = true;
            this.位图化.SuperTip = "将选中的矢量图形或图表转换为位图格式，可以防止图形在其他电脑上显示异常";
            this.位图化.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.位图化_Click);
            // 
            // 导出图片
            // 
            this.导出图片.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.导出图片.Image = global::WordMan.Properties.Resources.照片管理;
            this.导出图片.Label = "导出图片";
            this.导出图片.Name = "导出图片";
            this.导出图片.OfficeImageId = "FileSaveAs";
            this.导出图片.ScreenTip = "将选中的图片高清导出到指定位置";
            this.导出图片.ShowImage = true;
            this.导出图片.SuperTip = "将文档中选中的图片以高清质量导出到指定位置，支持多种图片格式";
            this.导出图片.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.导出图片_Click);
            // 
            // 排版工具
            // 
            this.排版工具.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.排版工具.Image = global::WordMan.Properties.Resources.格式刷;
            this.排版工具.Label = "排版工具";
            this.排版工具.Name = "排版工具";
            this.排版工具.OfficeImageId = "FormatPainter";
            this.排版工具.ScreenTip = "打开排版工具任务窗格";
            this.排版工具.ShowImage = true;
            this.排版工具.SuperTip = "打开或关闭排版工具任务窗格，提供便捷的排版工具集";
            this.排版工具.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TypesettingButton_Click);
            // 
            // 样式设置
            // 
            this.样式设置.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.样式设置.Image = global::WordMan.Properties.Resources.笔筒;
            this.样式设置.Label = "样式设置";
            this.样式设置.Name = "样式设置";
            this.样式设置.OfficeImageId = "CaptionInsert";
            this.样式设置.ScreenTip = "设置文档样式和格式";
            this.样式设置.ShowImage = true;
            this.样式设置.SuperTip = "打开样式设置窗口，可以批量修改、导入、导出文档样式和格式";
            this.样式设置.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.样式设置_Click);
            // 
            // 多级列表
            // 
            this.多级列表.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.多级列表.Image = global::WordMan.Properties.Resources.列表模式;
            this.多级列表.Label = "多级列表";
            this.多级列表.Name = "多级列表";
            this.多级列表.OfficeImageId = "Numbering";
            this.多级列表.ScreenTip = "设置多级列表格式";
            this.多级列表.ShowImage = true;
            this.多级列表.SuperTip = "打开多级列表设置窗口，可以配置和修改文档的多级列表格式";
            this.多级列表.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.多级列表_Click);
            // 
            // 域名高亮
            // 
            this.域名高亮.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.域名高亮.Image = global::WordMan.Properties.Resources.彩色;
            this.域名高亮.Items.Add(this.取消高亮);
            this.域名高亮.Label = "域名高亮";
            this.域名高亮.Name = "域名高亮";
            this.域名高亮.OfficeImageId = "TextHighlightColorPicker";
            this.域名高亮.ScreenTip = "高亮显示交叉引用和文献引用";
            this.域名高亮.SuperTip = "高亮显示文档中的所有交叉引用和文献引用字段，方便查看和定位";
            this.域名高亮.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.域名高亮_Click);
            // 
            // 取消高亮
            // 
            this.取消高亮.Image = global::WordMan.Properties.Resources.黑色;
            this.取消高亮.Label = "取消高亮";
            this.取消高亮.Name = "取消高亮";
            this.取消高亮.OfficeImageId = "FormatPlaceholder";
            this.取消高亮.ScreenTip = "取消交叉引用和文献引用的高亮显示";
            this.取消高亮.ShowImage = true;
            this.取消高亮.SuperTip = "取消文档中交叉引用和文献引用的高亮显示，恢复正常显示状态";
            this.取消高亮.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.取消高亮_Click);
            // 
            // 另存PDF
            // 
            this.另存PDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.另存PDF.Image = global::WordMan.Properties.Resources.pdf;
            this.另存PDF.Items.Add(this.版本);
            this.另存PDF.Label = "另存PDF";
            this.另存PDF.Name = "另存PDF";
            this.另存PDF.OfficeImageId = "FileSaveAs";
            this.另存PDF.ScreenTip = "将文档另存为PDF格式";
            this.另存PDF.SuperTip = "将当前文档另存为PDF格式文件，方便文档分享和打印";
            this.另存PDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.另存PDF_Click);
            // 
            // 版本
            // 
            this.版本.Label = "版本V3.1";
            this.版本.Name = "版本";
            this.版本.OfficeImageId = "Info";
            this.版本.ScreenTip = "查看WordMan插件版本信息";
            this.版本.ShowImage = true;
            this.版本.SuperTip = "查看WordMan插件的版本信息和相关说明";
            this.版本.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.版本_Click);
            // 
            // 编号设置
            // 
            this.编号设置.Image = global::WordMan.Properties.Resources.魔术棒;
            this.编号设置.Items.Add(this.上标);
            this.编号设置.Items.Add(this.正常);
            this.编号设置.Label = "文献编号";
            this.编号设置.Name = "编号设置";
            this.编号设置.OfficeImageId = "ControlWizards";
            this.编号设置.ScreenTip = "设置文献编号格式";
            this.编号设置.SuperTip = "设置文献引用编号的显示格式，可以选择上标或正常格式";
            // 
            // 上标
            // 
            this.上标.Label = "上标";
            this.上标.Name = "上标";
            this.上标.ScreenTip = "将文献引用设置为上标格式";
            this.上标.ShowImage = true;
            this.上标.SuperTip = "将文档中的所有文献引用字段设置为上标格式";
            this.上标.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.上标_Click);
            // 
            // 正常
            // 
            this.正常.Label = "正常";
            this.正常.Name = "正常";
            this.正常.ScreenTip = "将文献引用设置为正常格式";
            this.正常.ShowImage = true;
            this.正常.SuperTip = "将文档中的所有文献引用字段设置为正常格式（非上标）";
            this.正常.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.正常_Click);
            // 
            // 快速密级
            // 
            this.快速密级.Items.Add(this.公开);
            this.快速密级.Items.Add(this.内部);
            this.快速密级.Items.Add(this.移除密级);
            this.快速密级.Label = "快速密级";
            this.快速密级.Name = "快速密级";
            this.快速密级.ScreenTip = "快速密级";
            this.快速密级.SuperTip = "快速添加或移除文档的密级标签";
            // 
            // 公开
            // 
            this.公开.Label = "公开";
            this.公开.Name = "公开";
            this.公开.ScreenTip = "添加公开密级标签";
            this.公开.ShowImage = true;
            this.公开.SuperTip = "在文档当前页添加公开密级标签";
            this.公开.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.公开_Click);
            // 
            // 内部
            // 
            this.内部.Label = "内部";
            this.内部.Name = "内部";
            this.内部.ScreenTip = "添加内部密级标签";
            this.内部.ShowImage = true;
            this.内部.SuperTip = "在文档当前页添加内部密级标签";
            this.内部.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.内部_Click);
            // 
            // 移除密级
            // 
            this.移除密级.Label = "移除";
            this.移除密级.Name = "移除密级";
            this.移除密级.ScreenTip = "移除当前页的密级标签";
            this.移除密级.ShowImage = true;
            this.移除密级.SuperTip = "移除文档当前页的密级标签";
            this.移除密级.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.移除密级_Click);
            // 
            // 文档操作
            // 
            this.文档操作.Items.Add(this.文档合并);
            this.文档操作.Items.Add(this.文档拆分);
            this.文档操作.Label = "文档操作";
            this.文档操作.Name = "文档操作";
            this.文档操作.ScreenTip = "文档操作";
            this.文档操作.SuperTip = "提供文档合并和拆分功能";
            // 
            // 文档合并
            // 
            this.文档合并.Label = "文档合并";
            this.文档合并.Name = "文档合并";
            this.文档合并.ScreenTip = "将多个文档合并为一个文档";
            this.文档合并.ShowImage = true;
            this.文档合并.SuperTip = "将多个Word文档按顺序合并为一个文档，方便文档整理和归档";
            this.文档合并.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文档合并_Click);
            // 
            // 文档拆分
            // 
            this.文档拆分.Label = "文档拆分";
            this.文档拆分.Name = "文档拆分";
            this.文档拆分.ScreenTip = "将文档按指定规则拆分为多个文档";
            this.文档拆分.ShowImage = true;
            this.文档拆分.SuperTip = "将当前文档按照指定规则（如按页拆分）拆分为多个独立的文档";
            this.文档拆分.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文档拆分_Click);
            // 
            // 字母替换
            // 
            this.字母替换.Label = "";
            this.字母替换.Name = "字母替换";
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.WordMan);
            this.WordMan.ResumeLayout(false);
            this.WordMan.PerformLayout();
            this.文本处理.ResumeLayout(false);
            this.文本处理.PerformLayout();
            this.表格处理.ResumeLayout(false);
            this.表格处理.PerformLayout();
            this.题注与引用.ResumeLayout(false);
            this.题注与引用.PerformLayout();
            this.图片处理.ResumeLayout(false);
            this.图片处理.PerformLayout();
            this.全文处理.ResumeLayout(false);
            this.全文处理.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab WordMan;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 文本处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 清除格式;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 格式刷;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 只留文本;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除断行;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除空格;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除空行;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 英标转中标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 中标转英标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 自动加空格;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 缩进2字符;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 去除缩进;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 希腊字母;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 常用符号;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 表格处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery 创建表格;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery 设置表格;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 插入N行;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 插入N列;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 重复标题行;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 表编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 表注样式1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 表注样式2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 表注样式3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 题注与引用;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 式编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 公式样式1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 公式样式2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 公式样式3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 图编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 图注样式1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 图注样式2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 图注样式3;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 宽度刷;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 高度刷;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 位图化;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 导出图片;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 全文处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton 交叉引用;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 排版工具;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 样式设置;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 多级列表;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 域名高亮;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 取消高亮;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 编号设置;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 上标;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 正常;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 另存PDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 版本;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 文档操作;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 文档合并;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 文档拆分;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 快速密级;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 公开;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 内部;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 移除密级;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu 字体替换;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 仿宋替换;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 楷体替换;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 方正小标宋替换;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 数字替换;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 字母替换;
    }

    partial class ThisRibbonCollection
    {

    }
}
