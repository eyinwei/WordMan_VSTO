using System;
using System.Drawing;
using System.Windows.Forms;
using Color = System.Drawing.Color;

namespace WordMan_VSTO.MultiLevel
{
    public class ToggleButton : Button
    {
        private bool pressed;

        public bool Pressed
        {
            get
            {
                return pressed;
            }
            set
            {
                pressed = value;
                BackColor = (pressed ? Color.DarkGray : Color.AliceBlue);
                OnPressedChanged(EventArgs.Empty);
            }
        }

        public event EventHandler PressedChanged;

        public ToggleButton()
        {
            pressed = false;
            BackColor = Color.AliceBlue;
        }

        protected virtual void OnPressedChanged(EventArgs e)
        {
            this.PressedChanged?.Invoke(this, e);
        }

        protected override void OnClick(EventArgs e)
        {
            pressed = !pressed;
            BackColor = (pressed ? Color.DarkGray : Color.AliceBlue);
            OnPressedChanged(EventArgs.Empty);
            base.OnClick(e);
        }
    }
}
