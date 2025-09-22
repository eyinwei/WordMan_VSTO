using System.Drawing;
using System.Windows.Forms;

namespace WordFormatHelper{

public class TextureSelectComboBox : ComboBox
{
	public TextureSelectComboBox()
	{
		base.DrawMode = DrawMode.OwnerDrawVariable;
		base.DropDownStyle = ComboBoxStyle.DropDownList;
	}

	protected override void OnDrawItem(DrawItemEventArgs e)
	{
		base.OnDrawItem(e);
		if (base.DesignMode || e.Index == -1)
		{
			return;
		}
		Color color = (base.Enabled ? ForeColor : SystemColors.GrayText);
		Rectangle bounds = e.Bounds;
		e.Graphics.FillRectangle(Brushes.CornflowerBlue, e.Bounds);
		if ((e.State & DrawItemState.Focus) == 0)
		{
			e.Graphics.FillRectangle(Brushes.White, e.Bounds);
		}
		using Pen pen = ((e.Index > 5) ? new Pen(Brushes.Black) : new Pen(Color.FromArgb(255, 180, 180, 180)));
		using SolidBrush brush = new SolidBrush(color);
		e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
		switch (e.Index)
		{
		case 0:
		case 6:
		{
			Point pt7 = new Point(bounds.X - bounds.Height - 4, bounds.Y);
			Point pt8 = new Point(bounds.X - 4, bounds.Y + bounds.Height);
			for (int l = 0; l < (e.Bounds.Width + e.Bounds.Height) / 4 + 1; l++)
			{
				pt7.X += 4;
				pt8.X += 4;
				e.Graphics.DrawLine(pen, pt7, pt8);
			}
			pt7 = new Point(bounds.X - 4, bounds.Y);
			pt8 = new Point(bounds.X - bounds.Height - 4, bounds.Y + bounds.Height);
			for (int m = 0; m < (e.Bounds.Width + e.Bounds.Height) / 4 + 1; m++)
			{
				pt7.X += 4;
				pt8.X += 4;
				e.Graphics.DrawLine(pen, pt7, pt8);
			}
			break;
		}
		case 1:
		case 7:
		{
			Point pt11 = new Point(bounds.X - 4, bounds.Y);
			Point pt12 = new Point(bounds.X - 4, bounds.Y + bounds.Height);
			for (int num = 0; num < e.Bounds.Width / 4 + 1; num++)
			{
				pt11.X += 4;
				pt12.X += 4;
				e.Graphics.DrawLine(pen, pt11, pt12);
			}
			pt11 = new Point(bounds.X, bounds.Y - 4);
			pt12 = new Point(bounds.X + e.Bounds.Width, bounds.Y - 4);
			for (int num2 = 0; num2 < e.Bounds.Height / 4 + 1; num2++)
			{
				pt11.Y += 4;
				pt12.Y += 4;
				e.Graphics.DrawLine(pen, pt11, pt12);
			}
			break;
		}
		case 2:
		case 8:
		{
			Point pt3 = new Point(bounds.X - bounds.Height - 4, bounds.Y);
			Point pt4 = new Point(bounds.X - 4, bounds.Y + bounds.Height);
			for (int j = 0; j < (e.Bounds.Width + e.Bounds.Height) / 4 + 1; j++)
			{
				pt3.X += 4;
				pt4.X += 4;
				e.Graphics.DrawLine(pen, pt3, pt4);
			}
			break;
		}
		case 3:
		case 9:
		{
			Point pt9 = new Point(bounds.X - 4, bounds.Y);
			Point pt10 = new Point(bounds.X - bounds.Height - 4, bounds.Y + bounds.Height);
			for (int n = 0; n < (e.Bounds.Width + e.Bounds.Height) / 4 + 1; n++)
			{
				pt9.X += 4;
				pt10.X += 4;
				e.Graphics.DrawLine(pen, pt9, pt10);
			}
			break;
		}
		case 4:
		case 10:
		{
			Point pt5 = new Point(bounds.X - 4, bounds.Y);
			Point pt6 = new Point(bounds.X - 4, bounds.Y + bounds.Height);
			for (int k = 0; k < e.Bounds.Width / 4 + 1; k++)
			{
				pt5.X += 4;
				pt6.X += 4;
				e.Graphics.DrawLine(pen, pt5, pt6);
			}
			break;
		}
		case 5:
		case 11:
		{
			Point pt = new Point(bounds.X, bounds.Y - 4);
			Point pt2 = new Point(bounds.X + e.Bounds.Width, bounds.Y - 4);
			for (int i = 0; i < e.Bounds.Height / 4 + 1; i++)
			{
				pt.Y += 4;
				pt2.Y += 4;
				e.Graphics.DrawLine(pen, pt, pt2);
			}
			break;
		}
		}
	}
}
}