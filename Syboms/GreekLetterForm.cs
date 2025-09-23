using Word=Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WordMan
{
    public partial class GreekLetterForm : Form
    {
        public GreekLetterForm()
        {
            InitializeComponent();
            // 绑定事件
            chkBold.CheckedChanged += CheckBox_StyleChanged;
            chkItalic.CheckedChanged += CheckBox_StyleChanged;
            chkUppercase.CheckedChanged += CheckBox_StyleChanged;
        }

        // 常规希腊小写字母
        public class GreekLetterInfo
        {
            public string Lower { get; set; }
            public string Upper { get; set; }
            public string Name { get; set; }
        }

        private readonly GreekLetterInfo[] GreekLetters =
        {
        new GreekLetterInfo { Lower = "α", Upper = "Α", Name = "Alpha" },
        new GreekLetterInfo { Lower = "β", Upper = "Β", Name = "Beta" },
        new GreekLetterInfo { Lower = "γ", Upper = "Γ", Name = "Gamma" },
        new GreekLetterInfo { Lower = "δ", Upper = "Δ", Name = "Delta" },
        new GreekLetterInfo { Lower = "ε", Upper = "Ε", Name = "Epsilon" },
        new GreekLetterInfo { Lower = "ζ", Upper = "Ζ", Name = "Zeta" },
        new GreekLetterInfo { Lower = "η", Upper = "Η", Name = "Eta" },
        new GreekLetterInfo { Lower = "θ", Upper = "Θ", Name = "Theta" },
        new GreekLetterInfo { Lower = "ι", Upper = "Ι", Name = "Iota" },
        new GreekLetterInfo { Lower = "κ", Upper = "Κ", Name = "Kappa" },
        new GreekLetterInfo { Lower = "λ", Upper = "Λ", Name = "Lambda" },
        new GreekLetterInfo { Lower = "μ", Upper = "Μ", Name = "Mu" },
        new GreekLetterInfo { Lower = "ν", Upper = "Ν", Name = "Nu" },
        new GreekLetterInfo { Lower = "ξ", Upper = "Ξ", Name = "Xi" },
        new GreekLetterInfo { Lower = "ο", Upper = "Ο", Name = "Omicron" },
        new GreekLetterInfo { Lower = "π", Upper = "Π", Name = "Pi" },
        new GreekLetterInfo { Lower = "ρ", Upper = "Ρ", Name = "Rho" },
        new GreekLetterInfo { Lower = "σ", Upper = "Σ", Name = "Sigma" },
        new GreekLetterInfo { Lower = "τ", Upper = "Τ", Name = "Tau" },
        new GreekLetterInfo { Lower = "υ", Upper = "Υ", Name = "Upsilon" },
        new GreekLetterInfo { Lower = "φ", Upper = "Φ", Name = "Phi" },
        new GreekLetterInfo { Lower = "χ", Upper = "Χ", Name = "Chi" },
        new GreekLetterInfo { Lower = "ψ", Upper = "Ψ", Name = "Psi" },
        new GreekLetterInfo { Lower = "ω", Upper = "Ω", Name = "Omega" }
    };

        private void GreekLetterForm_Load(object sender, EventArgs e)
        {
            tableLetters.Controls.Clear();
            for (int i = 0; i < GreekLetters.Length; i++)
            {
                var info = GreekLetters[i];
                Button btn = new Button
                {
                    Text = info.Lower,
                    Tag = info,
                    Dock = DockStyle.Fill,
                    Font = new Font("Times New Roman", 24, FontStyle.Regular),
                };
                btn.Click += LetterButton_Click;
                tableLetters.Controls.Add(btn);
            }
            UpdateLetterButtonsStyle(); // 刷新样式
        }

        // 按钮点击事件：插入字母
        private void LetterButton_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;

            var info = btn.Tag as GreekLetterInfo;
            if (info == null) return;

            bool isItalic = chkItalic.Checked;
            bool isUpper = chkUppercase.Checked;
            bool isBold = chkBold.Checked;

            string toInsert = isUpper ? info.Upper : info.Lower;

            var selection = Globals.ThisAddIn.Application.Selection;
            selection.TypeText(toInsert);

            // 选中刚插入的字母并设置样式
            selection.MoveLeft(Word.WdUnits.wdCharacter, 1, Word.WdMovementType.wdExtend);
            selection.Font.Name = "Times New Roman";
            selection.Font.Italic = isItalic ? 1 : 0;
            selection.Font.Bold = isBold ? 1 : 0;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }


        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CheckBox_StyleChanged(object sender, EventArgs e)
        {
            UpdateLetterButtonsStyle();
        }

        private void UpdateLetterButtonsStyle()
        {
            bool isItalic = chkItalic.Checked;
            bool isUpper = chkUppercase.Checked;
            bool isBold = chkBold.Checked;

            foreach (Control ctrl in tableLetters.Controls)
            {
                Button btn = ctrl as Button;
                if (btn != null && btn.Tag is GreekLetterInfo info)
                {
                    // 字符切换大小写
                    btn.Text = isUpper ? info.Upper : info.Lower;

                    // 字体样式设置
                    FontStyle style = FontStyle.Regular;
                    if (isBold) style |= FontStyle.Bold;
                    if (isItalic) style |= FontStyle.Italic;
                    btn.Font = new Font(btn.Font.FontFamily, btn.Font.Size, style);
                }
            }
        }
    }
}
