using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    public partial class CommonSymbolForm : Form
    {
        // 定义符号结构体
        public class SymbolItem
        {
            public string Symbol { get; set; }
            public string Name { get; set; }
        }

        // 分类符号数据
        private static readonly Dictionary<string, SymbolItem[]> CategorySymbols = new Dictionary<string, SymbolItem[]>
        {
            // 符号类
            ["符号"] = new SymbolItem[]
            {
                    new SymbolItem { Symbol = "℃", Name = "摄氏度" },
                    new SymbolItem { Symbol = "℉", Name = "华氏度" },
                    new SymbolItem { Symbol = "°", Name = "度" },
                    new SymbolItem { Symbol = "‰", Name = "千分号" },
                    new SymbolItem { Symbol = "‱", Name = "万分号" },
                    new SymbolItem { Symbol = "µ", Name = "微" },
                    new SymbolItem { Symbol = "Ω", Name = "欧姆" },
                    new SymbolItem { Symbol = "Å", Name = "埃" },
                    new SymbolItem { Symbol = "㏑", Name = "自然对数ln" },
                    new SymbolItem { Symbol = "㏒", Name = "常用对数log" },
                    new SymbolItem { Symbol = "→", Name = "右箭头" },
                    new SymbolItem { Symbol = "←", Name = "左箭头" },
                    new SymbolItem { Symbol = "↑", Name = "上箭头" },
                    new SymbolItem { Symbol = "↓", Name = "下箭头" },
                    new SymbolItem { Symbol = "↔", Name = "双向箭头" },
                    new SymbolItem { Symbol = "⇒", Name = "蕴含/推导" },
                    new SymbolItem { Symbol = "⇔", Name = "充要/等价" },
                    new SymbolItem { Symbol = "⤴", Name = "右上弯箭头" },
                    new SymbolItem { Symbol = "⤵", Name = "右下弯箭头" },
                    new SymbolItem { Symbol = "□", Name = "空心方框" },
                    new SymbolItem { Symbol = "■", Name = "实心方框" },
                    new SymbolItem { Symbol = "▫", Name = "小空心方块" },
                    new SymbolItem { Symbol = "▪", Name = "小黑方块" },
                    new SymbolItem { Symbol = "▣", Name = "勾选方框" },
                    new SymbolItem { Symbol = "☐", Name = "选择框" },
                    new SymbolItem { Symbol = "☑", Name = "对勾框" },
                    new SymbolItem { Symbol = "☒", Name = "叉选框" },
                    new SymbolItem { Symbol = "○", Name = "空心圆" },
                    new SymbolItem { Symbol = "●", Name = "实心圆" },
                    new SymbolItem { Symbol = "◦", Name = "小空心圆" },
                    new SymbolItem { Symbol = "●", Name = "小黑圆" },
                    new SymbolItem { Symbol = "◎", Name = "圆环" },
                    new SymbolItem { Symbol = "★", Name = "黑星" },
                    new SymbolItem { Symbol = "☆", Name = "空心星" },
                    new SymbolItem { Symbol = "▲", Name = "黑三角" },
                    new SymbolItem { Symbol = "△", Name = "空心三角" },
                    new SymbolItem { Symbol = "◆", Name = "黑方菱形" },
                    new SymbolItem { Symbol = "◇", Name = "空心方菱形" },
                    new SymbolItem { Symbol = "※", Name = "参考标记" },
                    new SymbolItem { Symbol = "§", Name = "节号" }
            },
            // 数学类
            ["数学"] = new SymbolItem[]
            {
                new SymbolItem { Symbol = "·", Name = "乘积点/化学点" },
                new SymbolItem { Symbol = "×", Name = "乘号" },
                new SymbolItem { Symbol = "÷", Name = "除号" },
                new SymbolItem{ Symbol="∑", Name="求和" },
                new SymbolItem{ Symbol="∏", Name="连乘积" },
                new SymbolItem { Symbol = "±", Name = "正负号" },
                new SymbolItem { Symbol = "∓", Name = "负正号" },
                new SymbolItem { Symbol = "≈", Name = "约等于" },
                new SymbolItem { Symbol = "≠", Name = "不等于" },
                new SymbolItem { Symbol = "≡", Name = "恒等于" },
                new SymbolItem { Symbol = "≥", Name = "大于等于" },
                new SymbolItem { Symbol = "≤", Name = "小于等于" },
                new SymbolItem { Symbol = "≪", Name = "远小于" },
                new SymbolItem { Symbol = "≫", Name = "远大于" },
                new SymbolItem { Symbol = "≅", Name = "全等于" },
                new SymbolItem { Symbol = "∝", Name = "成正比" },
                new SymbolItem { Symbol = "∥", Name = "平行于" },
                new SymbolItem { Symbol = "∠", Name = "角" },
                new SymbolItem { Symbol = "∟", Name = "直角" },
                new SymbolItem { Symbol = "∩", Name = "交集" },
                new SymbolItem { Symbol = "∪", Name = "并集" },
                new SymbolItem { Symbol = "∈", Name = "属于" },
                new SymbolItem { Symbol = "∉", Name = "不属于" },
                new SymbolItem { Symbol = "∅", Name = "空集" },
                new SymbolItem { Symbol = "⊂", Name = "真子集" },
                new SymbolItem { Symbol = "⊃", Name = "真超集" },
                new SymbolItem { Symbol = "⊆", Name = "子集" },
                new SymbolItem { Symbol = "⊇", Name = "超集" },
                new SymbolItem { Symbol = "⊥", Name = "垂直于" },
                new SymbolItem { Symbol = "∵", Name = "因为" },
                new SymbolItem { Symbol = "∴", Name = "所以" },
                new SymbolItem { Symbol = "∫", Name = "积分" },
                new SymbolItem { Symbol = "∬", Name = "二重积分" },
                new SymbolItem { Symbol = "∭", Name = "三重积分" },
                new SymbolItem { Symbol = "∮", Name = "曲线积分" },
                new SymbolItem { Symbol = "√", Name = "平方根" },
                new SymbolItem { Symbol = "∛", Name = "立方根" },
                new SymbolItem { Symbol = "∜", Name = "四次根" }

            },
            // 序号类
            ["序号"] = new SymbolItem[]
            {
                // 序号符号（可用于条列、分节、项目编号等）
                new SymbolItem { Symbol = "①", Name = "带圈数字1" },
                new SymbolItem { Symbol = "②", Name = "带圈数字2" },
                new SymbolItem { Symbol = "③", Name = "带圈数字3" },
                new SymbolItem { Symbol = "④", Name = "带圈数字4" },
                new SymbolItem { Symbol = "⑤", Name = "带圈数字5" },
                new SymbolItem { Symbol = "⑥", Name = "带圈数字6" },
                new SymbolItem { Symbol = "⑦", Name = "带圈数字7" },
                new SymbolItem { Symbol = "⑧", Name = "带圈数字8" },
                new SymbolItem { Symbol = "⑨", Name = "带圈数字9" },
                new SymbolItem { Symbol = "⑩", Name = "带圈数字10" },
                new SymbolItem { Symbol = "⑴", Name = "带圈序号1" },
                new SymbolItem { Symbol = "⑵", Name = "带圈序号2" },
                new SymbolItem { Symbol = "⑶", Name = "带圈序号3" },
                new SymbolItem { Symbol = "⑷", Name = "带圈序号4" },
                new SymbolItem { Symbol = "⑸", Name = "带圈序号5" },
                new SymbolItem { Symbol = "⑹", Name = "带圈序号6" },
                new SymbolItem { Symbol = "⑺", Name = "带圈序号7" },
                new SymbolItem { Symbol = "⑻", Name = "带圈序号8" },
                new SymbolItem { Symbol = "⑼", Name = "带圈序号9" },
                new SymbolItem { Symbol = "⑽", Name = "带圈序号10" },
                new SymbolItem { Symbol = "㈠", Name = "带圈汉字一" },
                new SymbolItem { Symbol = "㈡", Name = "带圈汉字二" },
                new SymbolItem { Symbol = "㈢", Name = "带圈汉字三" },
                new SymbolItem { Symbol = "㈣", Name = "带圈汉字四" },
                new SymbolItem { Symbol = "㈤", Name = "带圈汉字五" },
                new SymbolItem { Symbol = "㈥", Name = "带圈汉字六" },
                new SymbolItem { Symbol = "㈦", Name = "带圈汉字七" },
                new SymbolItem { Symbol = "㈧", Name = "带圈汉字八" },
                new SymbolItem { Symbol = "㈨", Name = "带圈汉字九" },
                new SymbolItem { Symbol = "㈩", Name = "带圈汉字十" },
                new SymbolItem { Symbol = "Ⅰ", Name = "罗马数字一" },
                new SymbolItem { Symbol = "Ⅱ", Name = "罗马数字二" },
                new SymbolItem { Symbol = "Ⅲ", Name = "罗马数字三" },
                new SymbolItem { Symbol = "Ⅳ", Name = "罗马数字四" },
                new SymbolItem { Symbol = "Ⅴ", Name = "罗马数字五" },
                new SymbolItem { Symbol = "Ⅵ", Name = "罗马数字六" },
                new SymbolItem { Symbol = "Ⅶ", Name = "罗马数字七" },
                new SymbolItem { Symbol = "Ⅷ", Name = "罗马数字八" },
                new SymbolItem { Symbol = "Ⅸ", Name = "罗马数字九" },
                new SymbolItem { Symbol = "Ⅹ", Name = "罗马数字十" },
                new SymbolItem { Symbol = "(1)", Name = "圆括号数字1" },
                new SymbolItem { Symbol = "(2)", Name = "圆括号数字2" },
                new SymbolItem { Symbol = "(3)", Name = "圆括号数字3" },
                new SymbolItem { Symbol = "(4)", Name = "圆括号数字4" },
                new SymbolItem { Symbol = "(5)", Name = "圆括号数字5" },
                new SymbolItem { Symbol = "(6)", Name = "圆括号数字6" },
                new SymbolItem { Symbol = "(7)", Name = "圆括号数字7" },
                new SymbolItem { Symbol = "(8)", Name = "圆括号数字8" },
                new SymbolItem { Symbol = "(9)", Name = "圆括号数字9" },
                new SymbolItem { Symbol = "(10)", Name = "圆括号数字10" }

            },
            ["扩展符号"] = new SymbolItem[]
            {
                // 逻辑与集合
                new SymbolItem { Symbol = "⊕", Name = "直和/异或" },
                new SymbolItem { Symbol = "⊖", Name = "对称差" },
                new SymbolItem { Symbol = "⊗", Name = "直积/张量积" },
                new SymbolItem { Symbol = "⊙", Name = "点积/哈达玛积" },
                new SymbolItem { Symbol = "⊘", Name = "带斜杠圆" },
                new SymbolItem { Symbol = "⊚", Name = "双圆积" },
                new SymbolItem { Symbol = "⊢", Name = "可推出" },
                new SymbolItem { Symbol = "⊨", Name = "语义蕴涵" },
                new SymbolItem { Symbol = "⊩", Name = "强语义蕴涵" },
                new SymbolItem { Symbol = "≃", Name = "同构于" },
                new SymbolItem { Symbol = "≅", Name = "全等于" },
                new SymbolItem { Symbol = "≌", Name = "约等于" },
                new SymbolItem { Symbol = "≐", Name = "定义为" },
                new SymbolItem { Symbol = "≔", Name = "赋值为" },
                new SymbolItem { Symbol = "≜", Name = "定义等号" },
                new SymbolItem { Symbol = "≞", Name = "量词等价" },
                new SymbolItem { Symbol = "∘", Name = "复合运算符" },
                new SymbolItem { Symbol = "∖", Name = "集合差" },
                new SymbolItem { Symbol = "∮", Name = "闭合曲线积分" },

                // 箭头与映射
                new SymbolItem { Symbol = "↕", Name = "上下箭头" },
                new SymbolItem { Symbol = "↖", Name = "左上箭头" },
                new SymbolItem { Symbol = "↗", Name = "右上箭头" },
                new SymbolItem { Symbol = "↘", Name = "右下箭头" },
                new SymbolItem { Symbol = "↙", Name = "左下箭头" },
                new SymbolItem { Symbol = "↩", Name = "左回转箭头" },
                new SymbolItem { Symbol = "↪", Name = "右回转箭头" },
                new SymbolItem { Symbol = "↻", Name = "顺时针圆箭头" },
                new SymbolItem { Symbol = "↺", Name = "逆时针圆箭头" },
                new SymbolItem { Symbol = "⇑", Name = "双线上箭头" },
                new SymbolItem { Symbol = "⇓", Name = "双线下箭头" },
                new SymbolItem { Symbol = "⇗", Name = "双线右上箭头" },
                new SymbolItem { Symbol = "⇘", Name = "双线右下箭头" },

                // 单位与科学
                new SymbolItem { Symbol = "㎏", Name = "千克" },
                new SymbolItem { Symbol = "㎎", Name = "毫克" },
                new SymbolItem { Symbol = "㎜", Name = "毫米" },
                new SymbolItem { Symbol = "㎝", Name = "厘米" },
                new SymbolItem { Symbol = "㎠", Name = "平方厘米" },
                new SymbolItem { Symbol = "㎢", Name = "平方千米" },
                new SymbolItem { Symbol = "㎖", Name = "毫升" },
                new SymbolItem { Symbol = "㎗", Name = "分升" },
                new SymbolItem { Symbol = "㎞", Name = "千米" },
                new SymbolItem { Symbol = "㎰", Name = "皮克" },
                new SymbolItem { Symbol = "㎲", Name = "微秒" },
                new SymbolItem { Symbol = "㎳", Name = "毫秒" },

                // 括号与引号
                new SymbolItem { Symbol = "『", Name = "左书名号" },
                new SymbolItem { Symbol = "』", Name = "右书名号" },
                new SymbolItem { Symbol = "〔", Name = "左中括号" },
                new SymbolItem { Symbol = "〕", Name = "右中括号" },
                new SymbolItem { Symbol = "〚", Name = "左上角括号" },
                new SymbolItem { Symbol = "〛", Name = "右上角括号" },

                // 特殊标记与其他
                new SymbolItem { Symbol = "†", Name = "匕首标记" },
                new SymbolItem { Symbol = "‡", Name = "双匕首标记" },
                new SymbolItem { Symbol = "¶", Name = "段落符" },
                new SymbolItem { Symbol = "‖", Name = "平行线符号" },
                new SymbolItem { Symbol = "※", Name = "注释标记" },
                new SymbolItem { Symbol = "♯", Name = "升号/井号" },
                new SymbolItem { Symbol = "♭", Name = "降号" },
                new SymbolItem { Symbol = "♮", Name = "还原号" },
                new SymbolItem { Symbol = "♣", Name = "梅花" },
                new SymbolItem { Symbol = "♥", Name = "红心" },
                new SymbolItem { Symbol = "♦", Name = "方片" },
                new SymbolItem { Symbol = "♠", Name = "黑桃" },
                new SymbolItem { Symbol = "♪", Name = "音符" },
                new SymbolItem { Symbol = "♬", Name = "双音符" },
                new SymbolItem { Symbol = "♩", Name = "四分音符" },
                new SymbolItem { Symbol = "☀", Name = "太阳" },
                new SymbolItem { Symbol = "☁", Name = "云" },
                new SymbolItem { Symbol = "☂", Name = "雨伞" },
                new SymbolItem { Symbol = "☃", Name = "雪人" },
                new SymbolItem { Symbol = "☽", Name = "月亮" },
                new SymbolItem { Symbol = "☼", Name = "太阳" },
                new SymbolItem { Symbol = "☯", Name = "太极" },
                new SymbolItem { Symbol = "☭", Name = "锤镰" },
                new SymbolItem { Symbol = "☮", Name = "和平" },
                new SymbolItem { Symbol = "☹", Name = "不高兴" },
                new SymbolItem { Symbol = "☺", Name = "微笑" },
                new SymbolItem { Symbol = "☻", Name = "黑脸笑" }
            }

        };

        public CommonSymbolForm()
        {
            InitializeComponent();
        }

        private void CommonSymbolForm_Load(object sender, EventArgs e)
        {
            // 分别为三个Tab动态生成TableLayoutPanel和符号按钮
            InitTabButtons(tabPageSymbols, CategorySymbols["符号"]);
            InitTabButtons(tabPageMath, CategorySymbols["数学"]);
            InitTabButtons(tabPageNumbers, CategorySymbols["序号"]);
            InitTabButtons(tabPageExtend, CategorySymbols["扩展符号"]);
        }

        /// <summary>
        /// 初始化TabPage上的按钮
        /// </summary>
        private void InitTabButtons(TabPage tab, SymbolItem[] items)
        {
            int colCount = 8; // 一行8个
            int rowCount = (int)Math.Ceiling(items.Length / (float)colCount);

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = colCount,
                RowCount = rowCount,
                Padding = new Padding(5),
                AutoSize = true
            };
            // 设置列宽度均分
            for (int i = 0; i < colCount; i++)
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / colCount));
            for (int i = 0; i < rowCount; i++)
                table.RowStyles.Add(new RowStyle(SizeType.Percent, 100f / rowCount));

            table.GrowStyle = TableLayoutPanelGrowStyle.FixedSize;
            tab.Controls.Clear();
            tab.Controls.Add(table);

            for (int i = 0; i < items.Length; i++)
            {
                var btn = new Button
                {
                    Text = items[i].Symbol,
                    Tag = items[i],
                    Font = new Font("Segoe UI Symbol", 20, FontStyle.Regular),
                    Dock = DockStyle.Fill,
                    Margin = new Padding(3),
                    Height = 48
                };
                btn.Click += SymbolButton_Click;
                toolTip1.SetToolTip(btn, items[i].Name);
                table.Controls.Add(btn, i % colCount, i / colCount);
            }
        }

        /// <summary>
        /// 按钮点击插入符号到Word
        /// </summary>
        private void SymbolButton_Click(object sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is SymbolItem item)
            {
                try
                {
                    var selection = Globals.ThisAddIn.Application.Selection;
                    selection.TypeText(item.Symbol);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("插入符号失败：" + ex.Message);
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
