using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordFormatHelper{

public class WaterMarkTextBox : TextBox
{
	private const uint EM_SETCUEBANNER = 5377u;

	private string watermarktext;

	public string WatermarkText
	{
		get
		{
			return watermarktext;
		}
		set
		{
			watermarktext = value;
			SetWatermark(watermarktext);
		}
	}

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

	private void SetWatermark(string watermarktext)
	{
		SendMessage(base.Handle, 5377u, 0u, watermarktext);
	}
}
}