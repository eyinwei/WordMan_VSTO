using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordMan_VSTO.StylePane
{
    public class NumericUpDownWithUnit : NumericUpDown
    {
        private const int EM_SETMARGINS = 211;
        private const int EC_RIGHTMARGIN = 2;

        private readonly Label label;

        public string Label
        {
            get
            {
                return label.Text;
            }
            set
            {
                label.Text = value;
                if (base.IsHandleCreated)
                {
                    SetMargin();
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hwnd, int msg, int wParam, int lParam);

        public NumericUpDownWithUnit()
        {
            Control control = base.Controls[1];
            label = new Label
            {
                Text = "cm",
                Dock = DockStyle.Right,
                AutoSize = true,
                BackColor = Color.Transparent
            };
            control.Controls.Add(label);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            SetMargin();
        }

        private void SetMargin()
        {
            SendMessage(base.Controls[1].Handle, 211, 2, label.Width << 16);
        }
    }
}
