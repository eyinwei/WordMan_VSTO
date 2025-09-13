using System.Drawing;
using System.Windows.Forms;

namespace WordFormatHelper{

public class WaterMarkTextBoxEx : TextBox
{
	private const int WM_PAINT = 15;

	private string waterMark;

	private Color waterTextColor;

	public string WaterMark
	{
		get
		{
			return waterMark;
		}
		set
		{
			waterMark = value;
		}
	}

	public Color WaterTextColor
	{
		get
		{
			return waterTextColor;
		}
		set
		{
			waterTextColor = value;
		}
	}

	protected override void WndProc(ref Message m)
	{
		base.WndProc(ref m);
		if (m.Msg == 15)
		{
			WmPaint();
		}
	}

	private void WmPaint()
	{
		using Graphics dc = Graphics.FromHwnd(base.Handle);
		if (Text.Length == 0 && !string.IsNullOrEmpty(waterMark) && !Focused)
		{
			TextFormatFlags textFormatFlags = TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter;
			if (RightToLeft == RightToLeft.Yes)
			{
				textFormatFlags |= TextFormatFlags.Right | TextFormatFlags.RightToLeft;
			}
			TextRenderer.DrawText(dc, waterMark, Font, base.ClientRectangle, waterTextColor, textFormatFlags);
		}
	}
}
}