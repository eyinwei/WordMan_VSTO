using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordFormatHelper{

public class TableGirdShow : UserControl
{
	public struct Frame
	{
		public bool Visible;

		public int LineStyle;

		public int Width;

		public Brush LineColor;

		public Frame([Optional] bool visible, [Optional] int lineStyle, [Optional] int lineWidth, [Optional] Brush lineColor)
		{
			Visible = visible;
			LineStyle = ((lineStyle == 0) ? 1 : lineStyle);
			Width = ((lineWidth == 0) ? 1 : lineWidth);
			LineColor = lineColor ?? Brushes.Black;
		}
	}

	private readonly Frame[] rowFrame;

	private readonly Frame[] columnFrame;

	private Frame diagonal;

	private int currentLineStyle;

	private int currentLineWidth;

	[Category("RowFrames")]
	public Frame FirstRowTopFrame
	{
		get
		{
			return rowFrame[0];
		}
		set
		{
			rowFrame[0] = value;
		}
	}

	[Category("RowFrames")]
	public Frame FirstRowBottomFrame
	{
		get
		{
			return rowFrame[1];
		}
		set
		{
			rowFrame[1] = value;
		}
	}

	[Category("RowFrames")]
	public Frame InnerRowFrame
	{
		get
		{
			return rowFrame[2];
		}
		set
		{
			rowFrame[2] = value;
		}
	}

	[Category("RowFrames")]
	public Frame LastRowTopFrame
	{
		get
		{
			return rowFrame[3];
		}
		set
		{
			rowFrame[3] = value;
		}
	}

	[Category("RowFrames")]
	public Frame LastRowBottomFrame
	{
		get
		{
			return rowFrame[4];
		}
		set
		{
			rowFrame[4] = value;
		}
	}

	[Category("ColumnFrames")]
	public Frame FirstColumnLeftFrame
	{
		get
		{
			return columnFrame[0];
		}
		set
		{
			columnFrame[0] = value;
		}
	}

	[Category("ColumnFrames")]
	public Frame FirstColumnRightFrame
	{
		get
		{
			return columnFrame[1];
		}
		set
		{
			columnFrame[1] = value;
		}
	}

	[Category("ColumnFrames")]
	public Frame InnerColumnFrame
	{
		get
		{
			return columnFrame[2];
		}
		set
		{
			columnFrame[2] = value;
		}
	}

	[Category("ColumnFrames")]
	public Frame LastColumnLeftFrame
	{
		get
		{
			return columnFrame[3];
		}
		set
		{
			columnFrame[3] = value;
		}
	}

	[Category("ColumnFrames")]
	public Frame LastColumnRightFrame
	{
		get
		{
			return columnFrame[4];
		}
		set
		{
			columnFrame[4] = value;
		}
	}

	public Frame Diagonal
	{
		get
		{
			return diagonal;
		}
		set
		{
			diagonal = value;
		}
	}

	public int CurrentLineStyle
	{
		get
		{
			return currentLineStyle;
		}
		set
		{
			currentLineStyle = value;
		}
	}

	public int CurrentLineWidth
	{
		get
		{
			return currentLineWidth;
		}
		set
		{
			currentLineWidth = value;
		}
	}

	public TableGirdShow()
	{
		SetStyle(ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, value: true);
		rowFrame = new Frame[5]
		{
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black)
		};
		columnFrame = new Frame[5]
		{
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black),
			new Frame(visible: false, 0, 0, Brushes.Black)
		};
		diagonal = new Frame(visible: false, 0, 0, Brushes.Black);
		currentLineStyle = 0;
		currentLineWidth = 1;
	}

	protected override void OnPaint(PaintEventArgs e)
	{
		base.OnPaint(e);
		int num = base.Width;
		int num2 = base.Height;
		int num3 = (num2 - 10) / 4;
		int num4 = (num - 10) / 4;
		Pen pen = new Pen(Brushes.LightGray, 2f)
		{
			DashStyle = DashStyle.Dot,
			Alignment = PenAlignment.Center
		};
		Point pt = new Point(5, 5);
		Point pt2 = new Point(num4 + 5, num3 + 5);
		if (diagonal.Visible)
		{
			Pen pen2 = new Pen(diagonal.LineColor, diagonal.Width);
			switch (diagonal.LineStyle)
			{
			case 0:
				e.Graphics.DrawLine(pen2, pt, pt2);
				break;
			case 1:
				pen2.Width = 2f;
				pen2.Alignment = PenAlignment.Center;
				e.Graphics.DrawLine(pen2, pt.X + 3, pt.Y, pt2.X, pt2.Y - 3);
				e.Graphics.DrawLine(pen2, pt.X, pt.Y + 3, pt2.X - 3, pt2.Y);
				break;
			case 2:
				pen2.Alignment = PenAlignment.Center;
				pen2.Width = 1f;
				e.Graphics.DrawLine(pen2, pt.X + 3, pt.Y, pt2.X, pt2.Y - 3);
				pen2.Width = 3f;
				e.Graphics.DrawLine(pen2, pt.X, pt.Y + 3, pt2.X - 3, pt2.Y);
				break;
			case 3:
				pen2.Alignment = PenAlignment.Center;
				pen2.Width = 3f;
				e.Graphics.DrawLine(pen2, pt.X + 3, pt.Y, pt2.X, pt2.Y - 3);
				pen2.Width = 1f;
				e.Graphics.DrawLine(pen2, pt.X, pt.Y + 3, pt2.X - 3, pt2.Y);
				break;
			}
			e.Graphics.DrawLine(pen2, pt, pt2);
		}
		else
		{
			e.Graphics.DrawLine(pen, pt, pt2);
		}
		for (int i = 0; i < 5; i++)
		{
			pt = new Point(5, 5 + i * num3);
			pt2 = new Point(num - 5, 5 + i * num3);
			if (rowFrame[i].Visible)
			{
				Pen pen3 = new Pen(rowFrame[i].LineColor, rowFrame[i].Width)
				{
					Alignment = PenAlignment.Center
				};
				switch (rowFrame[i].LineStyle)
				{
				case 0:
					e.Graphics.DrawLine(pen3, pt, pt2);
					break;
				case 1:
					pen3.Width = 2f;
					e.Graphics.DrawLine(pen3, pt, pt2);
					switch (i)
					{
					case 0:
						e.Graphics.DrawLine(pen3, pt.X, pt.Y + 2, pt2.X, pt2.Y + 2);
						break;
					case 4:
						e.Graphics.DrawLine(pen3, pt.X, pt.Y + 2, pt2.X, pt2.Y + 2);
						break;
					default:
						e.Graphics.DrawLine(pen3, pt.X, pt.Y + 2, pt2.X, pt2.Y + 2);
						break;
					}
					break;
				case 2:
					pen3.Width = 1f;
					e.Graphics.DrawLine(pen3, pt, pt2);
					pen3.Width = 3f;
					e.Graphics.DrawLine(pen3, pt.X, pt.Y + 2, pt2.X, pt2.Y + 2);
					break;
				case 3:
					pen3.Width = 3f;
					e.Graphics.DrawLine(pen3, pt, pt2);
					pen3.Width = 1f;
					e.Graphics.DrawLine(pen3, pt.X, pt.Y + 2, pt2.X, pt2.Y + 2);
					break;
				}
			}
			else
			{
				e.Graphics.DrawLine(pen, pt, pt2);
			}
		}
		for (int j = 0; j < 5; j++)
		{
			pt = new Point(5 + j * num4, 5);
			pt2 = new Point(5 + j * num4, num2 - 5);
			if (columnFrame[j].Visible)
			{
				Pen pen4 = new Pen(columnFrame[j].LineColor, columnFrame[j].Width);
				switch (columnFrame[j].LineStyle)
				{
				case 0:
					e.Graphics.DrawLine(pen4, pt, pt2);
					break;
				case 1:
					pen4.Width = 2f;
					pen4.Alignment = PenAlignment.Center;
					e.Graphics.DrawLine(pen4, pt, pt2);
					e.Graphics.DrawLine(pen4, pt.X + 1, pt.Y, pt2.X + 1, pt2.Y);
					break;
				case 2:
					pen4.Alignment = PenAlignment.Center;
					pen4.Width = 1f;
					e.Graphics.DrawLine(pen4, pt, pt2);
					pen4.Width = 3f;
					e.Graphics.DrawLine(pen4, pt, pt2);
					break;
				case 3:
					pen4.Alignment = PenAlignment.Center;
					pen4.Width = 3f;
					e.Graphics.DrawLine(pen4, pt, pt2);
					pen4.Width = 1f;
					e.Graphics.DrawLine(pen4, pt.X + 2, pt.Y, pt2.X + 2, pt2.Y);
					break;
				}
				e.Graphics.DrawLine(new Pen(columnFrame[j].LineColor, columnFrame[j].Width), pt, pt2);
			}
			else
			{
				e.Graphics.DrawLine(pen, pt, pt2);
			}
		}
	}

	protected override void OnClick(EventArgs e)
	{
		base.OnClick(e);
		int num = (e as MouseEventArgs).X;
		int num2 = (e as MouseEventArgs).Y;
		int num3 = base.Width;
		int num4 = (base.Height - 10) / 4;
		int num5 = (num3 - 10) / 4;
		int num6 = (int)Math.Round((double)(num2 - 5) / (double)num4, MidpointRounding.AwayFromZero);
		if (num6 * num4 < num2 && num2 < num6 * num4 + 10)
		{
			rowFrame[num6].Visible = !rowFrame[num6].Visible;
			rowFrame[num6].Width = currentLineWidth;
			rowFrame[num6].LineStyle = currentLineStyle;
		}
		num6 = (int)Math.Round((double)(num - 5) / (double)num5, MidpointRounding.AwayFromZero);
		if (num6 * num5 < num && num < num6 * num5 + 10)
		{
			columnFrame[num6].Visible = !columnFrame[num6].Visible;
			columnFrame[num6].Width = currentLineWidth;
			columnFrame[num6].LineStyle = currentLineStyle;
		}
		if (num > 10 && num < num5 && Math.Abs((double)(num2 - 5) / (double)(num - 5) - (double)num4 / (double)num5) < 0.1)
		{
			diagonal.Visible = !diagonal.Visible;
			diagonal.Width = currentLineWidth;
			diagonal.LineStyle = currentLineStyle;
		}
		Refresh();
	}
}
}