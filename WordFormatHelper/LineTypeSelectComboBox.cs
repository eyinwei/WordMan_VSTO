using System.Drawing;
using System.Windows.Forms;

namespace WordFormatHelper{

public class LineTypeSelectComboBox : ComboBox
{
	public LineTypeSelectComboBox()
	{
		base.DrawMode = DrawMode.OwnerDrawVariable;
		base.DropDownStyle = ComboBoxStyle.DropDownList;
	}

	protected override void OnDrawItem(DrawItemEventArgs e)
	{
		base.OnDrawItem(e);
		if (base.DesignMode)
		{
			return;
		}
		int num = 3;
		Color color = (base.Enabled ? ForeColor : SystemColors.GrayText);
		Rectangle bounds = e.Bounds;
		e.Graphics.FillRectangle(Brushes.CornflowerBlue, e.Bounds);
		if ((e.State & DrawItemState.Focus) == 0)
		{
			e.Graphics.FillRectangle(Brushes.White, e.Bounds);
		}
		float num2 = e.Graphics.MeasureString("双细实线", Font).Width + (float)num;
		using Pen pen = new Pen(color);
		using SolidBrush brush = new SolidBrush(color);
		if (base.Items.Count == 0)
		{
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			return;
		}
		switch (e.Index)
		{
		case 0:
		{
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			pen.Width = 2f;
			Point pt15 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2);
			Point pt16 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2);
			e.Graphics.DrawLine(pen, pt15, pt16);
			break;
		}
		case 1:
		{
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			Point pt11 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2 - 2);
			Point pt12 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2 - 2);
			Point pt13 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2 + 2);
			Point pt14 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2 + 2);
			pen.Width = 2f;
			e.Graphics.DrawLine(pen, pt11, pt12);
			e.Graphics.DrawLine(pen, pt13, pt14);
			break;
		}
		case 2:
		{
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			Point pt7 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2 - 2);
			Point pt8 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2 - 2);
			Point pt9 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2 + 2);
			Point pt10 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2 + 2);
			pen.Width = 2f;
			e.Graphics.DrawLine(pen, pt7, pt8);
			pen.Width = 4f;
			e.Graphics.DrawLine(pen, pt9, pt10);
			break;
		}
		case 3:
		{
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			Point pt3 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2 - 2);
			Point pt4 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2 - 2);
			Point pt5 = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2 + 2);
			Point pt6 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2 + 2);
			pen.Width = 4f;
			e.Graphics.DrawLine(pen, pt3, pt4);
			pen.Width = 2f;
			e.Graphics.DrawLine(pen, pt5, pt6);
			break;
		}
		case 4:
		{
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			pen.Width = 4f;
			Point pt = new Point(bounds.X + (int)num2, bounds.Y + bounds.Height / 2);
			Point pt2 = new Point(bounds.X + bounds.Width - num, bounds.Y + bounds.Height / 2);
			e.Graphics.DrawLine(pen, pt, pt2);
			break;
		}
		case 5:
			e.Graphics.DrawString(base.Items[e.Index].ToString(), Font, brush, bounds);
			break;
		}
	}
}
}