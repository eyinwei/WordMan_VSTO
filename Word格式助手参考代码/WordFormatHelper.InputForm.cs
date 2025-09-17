// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.InputForm
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

internal class InputForm : Form
{
	private string TextInfo = "";

	private bool OKPressed;

	public string InputText
	{
		get
		{
			return TextInfo;
		}
		set
		{
			TextInfo = value;
		}
	}

	public bool OK => OKPressed;

	public InputForm(string Title, [Optional] string defaultText)
	{
		base.MinimizeBox = false;
		base.MaximizeBox = false;
		base.ClientSize = new Size(400, 90);
		base.FormBorderStyle = FormBorderStyle.FixedToolWindow;
		Text = Title;
		base.StartPosition = FormStartPosition.CenterScreen;
		OKPressed = false;
		TextInfo = defaultText;
		TextBox textBox = new TextBox
		{
			Multiline = true,
			Font = new Font(new FontFamily(Font.Name), 10.5f),
			Top = 0,
			Left = 0,
			Width = 400,
			Height = 60,
			Text = defaultText
		};
		Button button = new Button
		{
			Top = 60,
			Left = 300,
			Width = 100,
			Height = 30,
			Text = "确定"
		};
		base.Controls.Add(textBox);
		base.Controls.Add(button);
		button.Click += OKClick;
		textBox.TextChanged += SaveText;
	}

	private void SaveText(object sender, EventArgs e)
	{
		TextInfo = (sender as TextBox).Text;
	}

	private void OKClick(object sender, EventArgs e)
	{
		base.DialogResult = DialogResult.OK;
		OKPressed = true;
		Close();
	}
}
