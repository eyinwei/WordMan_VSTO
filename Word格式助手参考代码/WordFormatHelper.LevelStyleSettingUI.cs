// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.LevelStyleSettingUI
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordFormatHelper;
using WordFormatHelper.Properties;

public class LevelStyleSettingUI : Form
{
	private readonly List<WordStyleInfo> Styles = new List<WordStyleInfo>();

	private readonly List<string> FontNames = new List<string>();

	private bool userChange;

	private IContainer components;

	private DataGridView Dta_StyleList;

	private GroupBox Grp_SetSelectedStyle;

	private ToggleButton Btn_Underline;

	private ToggleButton Btn_Italic;

	private ToggleButton Btn_Bold;

	private ComboBox Cmb_EngFontName;

	private ComboBox Cmb_ChnFontName;

	private Button Btn_SetStyles;

	private ComboBox Cmb_HAlignment;

	private ComboBox Cmb_SpaceAfter;

	private ComboBox Cmb_SpaceBefore;

	private ComboBox Cmb_FontSize;

	private ComboBox Cmb_LineSpace;

	private ToggleButton Btn_BreakBefore;

	private TextBox Txt_RightIndent;

	private TextBox Txt_LeftIndent;

	private Button Btn_FontColor;

	private Label label4;

	private Label label3;

	private Label label2;

	private Label label1;

	private Label label5;

	private Label label14;

	private Label label13;

	private Label label12;

	private Label label11;

	private Label label10;

	private Label label9;

	private Label label8;

	private Label label7;

	private Label label6;

	public LevelStyleSettingUI(int showLevels)
	{
		InitializeComponent();
		base.Icon = Resources.WAIcon;
		InstalledFontCollection installedFontCollection = new InstalledFontCollection();
		userChange = false;
		FontFamily[] families = installedFontCollection.Families;
		foreach (FontFamily fontFamily in families)
		{
			FontNames.Add(fontFamily.Name);
		}
		for (int j = 1; j <= showLevels; j++)
		{
			try
			{
				WdBuiltinStyle wdBuiltinStyle = j switch
				{
					1 => WdBuiltinStyle.wdStyleHeading1, 
					2 => WdBuiltinStyle.wdStyleHeading2, 
					3 => WdBuiltinStyle.wdStyleHeading3, 
					4 => WdBuiltinStyle.wdStyleHeading4, 
					5 => WdBuiltinStyle.wdStyleHeading5, 
					6 => WdBuiltinStyle.wdStyleHeading6, 
					7 => WdBuiltinStyle.wdStyleHeading7, 
					8 => WdBuiltinStyle.wdStyleHeading8, 
					9 => WdBuiltinStyle.wdStyleHeading9, 
					_ => (WdBuiltinStyle)0, 
				};
				Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
				object Index = wdBuiltinStyle;
				Style style = styles[ref Index];
				Styles.Add(new WordStyleInfo(style, wdBuiltinStyle));
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
		Dta_StyleList.Columns.Clear();
		Dta_StyleList.ReadOnly = true;
		Dta_StyleList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
		Dta_StyleList.Columns.AddRange(new DataGridViewTextBoxColumn
		{
			Name = "Col_StyleName",
			DataPropertyName = "StyleName",
			Frozen = true,
			HeaderText = "样式名",
			ReadOnly = true
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_ChnFontName",
			DataPropertyName = "ChnFontName",
			HeaderText = "中文字体"
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_EngFontName",
			DataPropertyName = "EngFontName",
			HeaderText = "西文字体"
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_FontSize",
			DataPropertyName = "FontSize",
			HeaderText = "字体大小",
			Width = 100
		}, new DataGridViewImageColumn
		{
			Name = "Col_FontColor",
			DataPropertyName = "FontColor",
			HeaderText = "颜色",
			ImageLayout = DataGridViewImageCellLayout.Normal,
			Width = 60
		}, new DataGridViewCheckBoxColumn
		{
			Name = "Col_FontBold",
			DataPropertyName = "Bold",
			HeaderText = "粗体",
			FalseValue = false,
			TrueValue = true,
			Width = 80
		}, new DataGridViewCheckBoxColumn
		{
			Name = "Col_FontItalic",
			DataPropertyName = "Italic",
			HeaderText = "斜体",
			FalseValue = false,
			TrueValue = true,
			Width = 80
		}, new DataGridViewCheckBoxColumn
		{
			Name = "Col_FontUnderline",
			DataPropertyName = "Underline",
			HeaderText = "下划线",
			FalseValue = false,
			TrueValue = true,
			Width = 80
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_LeftIndent",
			DataPropertyName = "LeftIndent",
			HeaderText = "左缩进",
			Width = 100
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_RightIndent",
			DataPropertyName = "RightIndent",
			HeaderText = "右缩进",
			Width = 100
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_LineSpace",
			DataPropertyName = "LineSpace",
			HeaderText = "行距",
			Width = 100
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_SpaceBefore",
			DataPropertyName = "SpaceBefore",
			HeaderText = "段前行距",
			Width = 100
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_SpaceAfter",
			DataPropertyName = "SpaceAfter",
			HeaderText = "段后行距",
			Width = 100
		}, new DataGridViewTextBoxColumn
		{
			Name = "Col_HAlignment",
			DataPropertyName = "HAlignment",
			HeaderText = "水平对齐",
			Width = 100
		}, new DataGridViewCheckBoxColumn
		{
			Name = "Col_BreakBefore",
			DataPropertyName = "BreakBefore",
			HeaderText = "段前分页",
			FalseValue = false,
			TrueValue = true,
			Width = 80
		});
		Dta_StyleList.DataSource = Styles;
		Dta_StyleList.CellFormatting += Dta_StyleList_CellFormatting;
		Cmb_ChnFontName.Items.Clear();
		ComboBox.ObjectCollection items = Cmb_ChnFontName.Items;
		List<string> fontNames = FontNames;
		int i = 0;
		object[] array = new object[fontNames.Count];
		foreach (string item in fontNames)
		{
			array[i] = item;
			i++;
		}
		items.AddRange(array);
		Cmb_EngFontName.Items.Clear();
		items = Cmb_EngFontName.Items;
		List<string> fontNames2 = FontNames;
		i = 0;
		array = new object[fontNames2.Count];
		foreach (string item2 in fontNames2)
		{
			array[i] = item2;
			i++;
		}
		items.AddRange(array);
		Cmb_HAlignment.Items.Clear();
		ComboBox.ObjectCollection items2 = Cmb_HAlignment.Items;
		array = WordStyleInfo.HAlignments;
		items2.AddRange(array);
		Cmb_LineSpace.Items.Clear();
		ComboBox.ObjectCollection items3 = Cmb_LineSpace.Items;
		array = WordStyleInfo.LineSpacingValues;
		items3.AddRange(array);
		Cmb_SpaceBefore.Items.Clear();
		ComboBox.ObjectCollection items4 = Cmb_SpaceBefore.Items;
		array = WordStyleInfo.ParagraphSpaceValues;
		items4.AddRange(array);
		Cmb_SpaceAfter.Items.Clear();
		ComboBox.ObjectCollection items5 = Cmb_SpaceAfter.Items;
		array = WordStyleInfo.ParagraphSpaceValues;
		items5.AddRange(array);
		Cmb_FontSize.Items.Clear();
		items = Cmb_FontSize.Items;
		List<string> fontSizeList = WordStyleInfo.FontSizeList;
		i = 0;
		array = new object[fontSizeList.Count];
		foreach (string item3 in fontSizeList)
		{
			array[i] = item3;
			i++;
		}
		items.AddRange(array);
		base.StartPosition = FormStartPosition.CenterScreen;
		userChange = true;
	}

	private void Dta_StyleList_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
	{
		if (e.ColumnIndex != Dta_StyleList.Columns["Col_FontColor"].Index || e.RowIndex < 0 || !(Dta_StyleList.Rows[e.RowIndex].DataBoundItem is WordStyleInfo wordStyleInfo))
		{
			return;
		}
		using Bitmap bitmap = new Bitmap(16, 16);
		using Graphics graphics = Graphics.FromImage(bitmap);
		graphics.Clear(wordStyleInfo.FontColor);
		e.Value = new Bitmap(bitmap);
	}

	private void Dta_StyleList_SelectionChanged(object sender, EventArgs e)
	{
		if (Dta_StyleList.SelectedRows.Count > 0)
		{
			SetValueByStyle(Styles[Dta_StyleList.SelectedRows[0].Index]);
		}
	}

	private void SetValueByStyle(WordStyleInfo style)
	{
		userChange = false;
		int selectedIndex = FontNames.IndexOf(style.ChnFontName);
		Cmb_ChnFontName.SelectedIndex = selectedIndex;
		selectedIndex = FontNames.IndexOf(style.EngFontName);
		Cmb_EngFontName.SelectedIndex = selectedIndex;
		selectedIndex = WordStyleInfo.FontSizeList.IndexOf(style.FontSize);
		if (selectedIndex != -1)
		{
			Cmb_FontSize.SelectedIndex = -1;
			Cmb_FontSize.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_FontSize.Text = style.FontSize;
		}
		Btn_FontColor.BackColor = style.FontColor;
		Btn_Bold.Pressed = style.Bold;
		Btn_Italic.Pressed = style.Italic;
		Btn_Underline.Pressed = style.Underline;
		Txt_LeftIndent.Text = style.LeftIndent;
		Txt_RightIndent.Text = style.RightIndent;
		selectedIndex = WordStyleInfo.LineSpacingValues.ToList().IndexOf(style.LineSpace);
		if (selectedIndex != -1)
		{
			Cmb_LineSpace.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_LineSpace.Text = style.LineSpace;
		}
		selectedIndex = WordStyleInfo.ParagraphSpaceValues.ToList().IndexOf(style.SpaceBefore);
		if (selectedIndex != -1)
		{
			Cmb_SpaceBefore.SelectedIndex = -1;
			Cmb_SpaceBefore.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_SpaceBefore.Text = style.SpaceBefore;
		}
		selectedIndex = WordStyleInfo.ParagraphSpaceValues.ToList().IndexOf(style.SpaceAfter);
		if (selectedIndex != -1)
		{
			Cmb_SpaceAfter.SelectedIndex = -1;
			Cmb_SpaceAfter.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_SpaceAfter.Text = style.SpaceAfter;
		}
		selectedIndex = WordStyleInfo.HAlignments.ToList().IndexOf(style.HAlignment);
		if (selectedIndex != 1)
		{
			Cmb_HAlignment.SelectedIndex = selectedIndex;
		}
		else
		{
			Cmb_HAlignment.SelectedIndex = -1;
		}
		Btn_BreakBefore.Pressed = style.BreakBefore;
		userChange = true;
	}

	private void ToggleButton_PressedChanged(object sender, EventArgs e)
	{
		if (!(sender is ToggleButton toggleButton))
		{
			return;
		}
		toggleButton.Text = (toggleButton.Pressed ? "是" : "否");
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		string columnName = "";
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			switch (toggleButton.Name)
			{
			case "Btn_Bold":
				Styles[selectedRow.Index].Bold = toggleButton.Pressed;
				columnName = "Col_FontBold";
				break;
			case "Btn_Italic":
				Styles[selectedRow.Index].Italic = toggleButton.Pressed;
				columnName = "Col_FontItalic";
				break;
			case "Btn_Underline":
				Styles[selectedRow.Index].Underline = toggleButton.Pressed;
				columnName = "Col_FontUnderline";
				break;
			case "Btn_BreakBefore":
				Styles[selectedRow.Index].BreakBefore = toggleButton.Pressed;
				columnName = "Col_BreakBefore";
				break;
			}
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
		}
	}

	private void Btn_FontColor_Click(object sender, EventArgs e)
	{
		ColorDialog colorDialog = new ColorDialog
		{
			Color = Btn_FontColor.BackColor,
			AnyColor = true,
			SolidColorOnly = true,
			AllowFullOpen = true,
			FullOpen = true
		};
		if (colorDialog.ShowDialog(this) == DialogResult.OK)
		{
			Btn_FontColor.BackColor = colorDialog.Color;
		}
	}

	private void Cmb_FontNameAndHV_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		string columnName = "";
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			switch ((sender as ComboBox).Name)
			{
			case "Cmb_ChnFontName":
				Styles[selectedRow.Index].ChnFontName = (sender as ComboBox).SelectedItem.ToString();
				columnName = "Col_ChnFontName";
				break;
			case "Cmb_EngFontName":
				Styles[selectedRow.Index].EngFontName = (sender as ComboBox).SelectedItem.ToString();
				columnName = "Col_EngFontName";
				break;
			case "Cmb_HAlignment":
				Styles[selectedRow.Index].HAlignment = (sender as ComboBox).SelectedItem.ToString();
				columnName = "Col_HAlignment";
				break;
			}
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
		}
	}

	private void Cmb_FontSize_TextChanged(object sender, EventArgs e)
	{
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			Styles[selectedRow.Index].FontSize = Cmb_FontSize.Text;
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns["Col_FontSize"].Index, selectedRow.Index);
		}
	}

	private void Cmb_FontSize_Validated(object sender, EventArgs e)
	{
		int num = WordStyleInfo.FontSizeList.IndexOf(Cmb_FontSize.Text);
		if (num != -1)
		{
			Cmb_FontSize.SelectedIndex = num;
		}
		else if (Regex.IsMatch(Cmb_FontSize.Text, "^\\d+(?:\\.(?:0|5))?(?:\\s+)?磅?$"))
		{
			string text = Cmb_FontSize.Text.TrimEnd(' ', '磅');
			num = WordStyleInfo.FontSizeValueList.IndexOf(Convert.ToSingle(text));
			if (num != -1)
			{
				Cmb_FontSize.SelectedIndex = num;
			}
			else
			{
				Cmb_FontSize.Text = text + " 磅";
			}
		}
		else
		{
			Cmb_FontSize.SelectedIndex = 10;
		}
	}

	private void Btn_FontColor_BackColorChanged(object sender, EventArgs e)
	{
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			Styles[selectedRow.Index].FontColor = Btn_FontColor.BackColor;
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns["Col_FontColor"].Index, selectedRow.Index);
		}
	}

	private void Txt_Indent_TextChanged(object sender, EventArgs e)
	{
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		string columnName = "";
		TextBox textBox = sender as TextBox;
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			string name = textBox.Name;
			if (!(name == "Txt_LeftIndent"))
			{
				if (name == "Txt_RightIndent")
				{
					Styles[selectedRow.Index].RightIndent = textBox.Text;
					columnName = "Col_RightIndent";
				}
			}
			else
			{
				Styles[selectedRow.Index].LeftIndent = textBox.Text;
				columnName = "Col_LeftIndent";
			}
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
		}
	}

	private void Txt_Indent_Validated(object sender, EventArgs e)
	{
		TextBox textBox = sender as TextBox;
		string s = textBox.Text.TrimEnd(' ', '磅', '厘', '米');
		try
		{
			float num = float.Parse(s);
			if (textBox.Text.EndsWith("厘米"))
			{
				textBox.Text = num.ToString("0.00 厘米");
			}
			else
			{
				textBox.Text = num.ToString("0.00 磅");
			}
		}
		catch
		{
			textBox.Text = "0.00 厘米";
		}
	}

	private void Cmb_LineSpace_Validated(object sender, EventArgs e)
	{
		if (Cmb_LineSpace.SelectedIndex != -1)
		{
			return;
		}
		string s = Cmb_LineSpace.Text.TrimEnd(' ', '磅', '行');
		try
		{
			float num = float.Parse(s);
			if (Cmb_LineSpace.Text.EndsWith("行"))
			{
				Cmb_LineSpace.Text = num.ToString("0.00 行");
			}
			else
			{
				Cmb_LineSpace.Text = num.ToString("0.00 磅");
			}
		}
		catch
		{
			Cmb_LineSpace.SelectedIndex = 1;
		}
	}

	private void Cmb_LineSpace_TextChanged(object sender, EventArgs e)
	{
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			Styles[selectedRow.Index].LineSpace = Cmb_LineSpace.Text;
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns["Col_LineSpace"].Index, selectedRow.Index);
		}
	}

	private void Cmb_SpaceValue_Validated(object sender, EventArgs e)
	{
		ComboBox comboBox = sender as ComboBox;
		if (comboBox.SelectedIndex != -1)
		{
			return;
		}
		string s = comboBox.Text.TrimEnd(' ', '磅', '行');
		try
		{
			float num = float.Parse(s);
			if (comboBox.Text.EndsWith("行"))
			{
				comboBox.Text = num.ToString("0.00 行");
			}
			else
			{
				comboBox.Text = num.ToString("0.00 磅");
			}
		}
		catch
		{
			comboBox.SelectedIndex = 1;
		}
	}

	private void Cmb_SpaceValue_TextChanged(object sender, EventArgs e)
	{
		if (!userChange || Dta_StyleList.SelectedRows.Count <= 0)
		{
			return;
		}
		ComboBox comboBox = sender as ComboBox;
		string columnName = "";
		foreach (DataGridViewRow selectedRow in Dta_StyleList.SelectedRows)
		{
			string name = comboBox.Name;
			if (!(name == "Cmb_SpaceBefore"))
			{
				if (name == "Cmb_SpaceAfter")
				{
					Styles[selectedRow.Index].SpaceAfter = comboBox.Text;
					columnName = "Col_SpaceAfter";
				}
			}
			else
			{
				Styles[selectedRow.Index].SpaceBefore = comboBox.Text;
				columnName = "Col_SpaceBefore";
			}
			Dta_StyleList.UpdateCellValue(Dta_StyleList.Columns[columnName].Index, selectedRow.Index);
		}
	}

	private void Btn_SetStyles_Click(object sender, EventArgs e)
	{
		string text = string.Empty;
		foreach (WordStyleInfo style in Styles)
		{
			if (!style.SetStyle(Globals.ThisAddIn.Application.ActiveDocument))
			{
				text = text + style.StyleName + ";";
			}
		}
		if (!string.IsNullOrEmpty(text))
		{
			MessageBox.Show("样式：" + text.TrimEnd(';') + " 引用设置时出现错误，请检查设置值是否正确！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
		else
		{
			MessageBox.Show("样式设置已应用到文档！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}
		Close();
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	private void InitializeComponent()
	{
		this.Dta_StyleList = new System.Windows.Forms.DataGridView();
		this.Grp_SetSelectedStyle = new System.Windows.Forms.GroupBox();
		this.Txt_RightIndent = new System.Windows.Forms.TextBox();
		this.Cmb_SpaceAfter = new System.Windows.Forms.ComboBox();
		this.label14 = new System.Windows.Forms.Label();
		this.label13 = new System.Windows.Forms.Label();
		this.label12 = new System.Windows.Forms.Label();
		this.label11 = new System.Windows.Forms.Label();
		this.label10 = new System.Windows.Forms.Label();
		this.label9 = new System.Windows.Forms.Label();
		this.label8 = new System.Windows.Forms.Label();
		this.label7 = new System.Windows.Forms.Label();
		this.label6 = new System.Windows.Forms.Label();
		this.label5 = new System.Windows.Forms.Label();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.label1 = new System.Windows.Forms.Label();
		this.Btn_FontColor = new System.Windows.Forms.Button();
		this.Txt_LeftIndent = new System.Windows.Forms.TextBox();
		this.Btn_BreakBefore = new WordFormatHelper.ToggleButton();
		this.Cmb_HAlignment = new System.Windows.Forms.ComboBox();
		this.Cmb_SpaceBefore = new System.Windows.Forms.ComboBox();
		this.Cmb_FontSize = new System.Windows.Forms.ComboBox();
		this.Cmb_LineSpace = new System.Windows.Forms.ComboBox();
		this.Cmb_EngFontName = new System.Windows.Forms.ComboBox();
		this.Cmb_ChnFontName = new System.Windows.Forms.ComboBox();
		this.Btn_Underline = new WordFormatHelper.ToggleButton();
		this.Btn_Italic = new WordFormatHelper.ToggleButton();
		this.Btn_Bold = new WordFormatHelper.ToggleButton();
		this.Btn_SetStyles = new System.Windows.Forms.Button();
		((System.ComponentModel.ISupportInitialize)this.Dta_StyleList).BeginInit();
		this.Grp_SetSelectedStyle.SuspendLayout();
		base.SuspendLayout();
		this.Dta_StyleList.BackgroundColor = System.Drawing.SystemColors.Window;
		this.Dta_StyleList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.Dta_StyleList.Dock = System.Windows.Forms.DockStyle.Top;
		this.Dta_StyleList.Location = new System.Drawing.Point(0, 0);
		this.Dta_StyleList.Name = "Dta_StyleList";
		this.Dta_StyleList.RowTemplate.Height = 23;
		this.Dta_StyleList.Size = new System.Drawing.Size(589, 265);
		this.Dta_StyleList.TabIndex = 0;
		this.Dta_StyleList.SelectionChanged += new System.EventHandler(Dta_StyleList_SelectionChanged);
		this.Grp_SetSelectedStyle.Controls.Add(this.Txt_RightIndent);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_SpaceAfter);
		this.Grp_SetSelectedStyle.Controls.Add(this.label14);
		this.Grp_SetSelectedStyle.Controls.Add(this.label13);
		this.Grp_SetSelectedStyle.Controls.Add(this.label12);
		this.Grp_SetSelectedStyle.Controls.Add(this.label11);
		this.Grp_SetSelectedStyle.Controls.Add(this.label10);
		this.Grp_SetSelectedStyle.Controls.Add(this.label9);
		this.Grp_SetSelectedStyle.Controls.Add(this.label8);
		this.Grp_SetSelectedStyle.Controls.Add(this.label7);
		this.Grp_SetSelectedStyle.Controls.Add(this.label6);
		this.Grp_SetSelectedStyle.Controls.Add(this.label5);
		this.Grp_SetSelectedStyle.Controls.Add(this.label4);
		this.Grp_SetSelectedStyle.Controls.Add(this.label3);
		this.Grp_SetSelectedStyle.Controls.Add(this.label2);
		this.Grp_SetSelectedStyle.Controls.Add(this.label1);
		this.Grp_SetSelectedStyle.Controls.Add(this.Btn_FontColor);
		this.Grp_SetSelectedStyle.Controls.Add(this.Txt_LeftIndent);
		this.Grp_SetSelectedStyle.Controls.Add(this.Btn_BreakBefore);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_HAlignment);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_SpaceBefore);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_FontSize);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_LineSpace);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_EngFontName);
		this.Grp_SetSelectedStyle.Controls.Add(this.Cmb_ChnFontName);
		this.Grp_SetSelectedStyle.Controls.Add(this.Btn_Underline);
		this.Grp_SetSelectedStyle.Controls.Add(this.Btn_Italic);
		this.Grp_SetSelectedStyle.Controls.Add(this.Btn_Bold);
		this.Grp_SetSelectedStyle.Dock = System.Windows.Forms.DockStyle.Top;
		this.Grp_SetSelectedStyle.Location = new System.Drawing.Point(0, 265);
		this.Grp_SetSelectedStyle.Name = "Grp_SetSelectedStyle";
		this.Grp_SetSelectedStyle.Size = new System.Drawing.Size(589, 195);
		this.Grp_SetSelectedStyle.TabIndex = 1;
		this.Grp_SetSelectedStyle.TabStop = false;
		this.Grp_SetSelectedStyle.Text = "为选中样式应用下列设置";
		this.Txt_RightIndent.Location = new System.Drawing.Point(270, 93);
		this.Txt_RightIndent.Name = "Txt_RightIndent";
		this.Txt_RightIndent.Size = new System.Drawing.Size(100, 25);
		this.Txt_RightIndent.TabIndex = 39;
		this.Txt_RightIndent.TextChanged += new System.EventHandler(Txt_Indent_TextChanged);
		this.Txt_RightIndent.Validated += new System.EventHandler(Txt_Indent_Validated);
		this.Cmb_SpaceAfter.FormattingEnabled = true;
		this.Cmb_SpaceAfter.Location = new System.Drawing.Point(270, 126);
		this.Cmb_SpaceAfter.Name = "Cmb_SpaceAfter";
		this.Cmb_SpaceAfter.Size = new System.Drawing.Size(100, 26);
		this.Cmb_SpaceAfter.TabIndex = 27;
		this.Cmb_SpaceAfter.TextChanged += new System.EventHandler(Cmb_SpaceValue_TextChanged);
		this.Cmb_SpaceAfter.Validated += new System.EventHandler(Cmb_SpaceValue_Validated);
		this.label14.AutoSize = true;
		this.label14.Location = new System.Drawing.Point(201, 62);
		this.label14.Name = "label14";
		this.label14.Size = new System.Drawing.Size(64, 18);
		this.label14.TabIndex = 54;
		this.label14.Text = "字体颜色";
		this.label13.AutoSize = true;
		this.label13.Location = new System.Drawing.Point(12, 164);
		this.label13.Name = "label13";
		this.label13.Size = new System.Drawing.Size(64, 18);
		this.label13.TabIndex = 53;
		this.label13.Text = "段落对齐";
		this.label12.AutoSize = true;
		this.label12.Location = new System.Drawing.Point(406, 130);
		this.label12.Name = "label12";
		this.label12.Size = new System.Drawing.Size(64, 18);
		this.label12.TabIndex = 52;
		this.label12.Text = "段前分页";
		this.label11.AutoSize = true;
		this.label11.Location = new System.Drawing.Point(201, 130);
		this.label11.Name = "label11";
		this.label11.Size = new System.Drawing.Size(64, 18);
		this.label11.TabIndex = 51;
		this.label11.Text = "段后间距";
		this.label10.AutoSize = true;
		this.label10.Location = new System.Drawing.Point(12, 130);
		this.label10.Name = "label10";
		this.label10.Size = new System.Drawing.Size(64, 18);
		this.label10.TabIndex = 50;
		this.label10.Text = "段前间距";
		this.label9.AutoSize = true;
		this.label9.Location = new System.Drawing.Point(406, 96);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(64, 18);
		this.label9.TabIndex = 49;
		this.label9.Text = "段落行距";
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(201, 96);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(50, 18);
		this.label8.TabIndex = 48;
		this.label8.Text = "右缩进";
		this.label7.AutoSize = true;
		this.label7.Location = new System.Drawing.Point(12, 96);
		this.label7.Name = "label7";
		this.label7.Size = new System.Drawing.Size(50, 18);
		this.label7.TabIndex = 47;
		this.label7.Text = "左缩进";
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(491, 62);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(50, 18);
		this.label6.TabIndex = 46;
		this.label6.Text = "下划线";
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(410, 62);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(36, 18);
		this.label5.TabIndex = 45;
		this.label5.Text = "斜体";
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(322, 62);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(36, 18);
		this.label4.TabIndex = 44;
		this.label4.Text = "粗体";
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(12, 62);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(64, 18);
		this.label3.TabIndex = 43;
		this.label3.Text = "字体大小";
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(312, 28);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(64, 18);
		this.label2.TabIndex = 42;
		this.label2.Text = "西文字体";
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(12, 28);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(64, 18);
		this.label1.TabIndex = 3;
		this.label1.Text = "中文字体";
		this.Btn_FontColor.Location = new System.Drawing.Point(270, 56);
		this.Btn_FontColor.Name = "Btn_FontColor";
		this.Btn_FontColor.Size = new System.Drawing.Size(40, 30);
		this.Btn_FontColor.TabIndex = 41;
		this.Btn_FontColor.UseVisualStyleBackColor = true;
		this.Btn_FontColor.BackColorChanged += new System.EventHandler(Btn_FontColor_BackColorChanged);
		this.Btn_FontColor.Click += new System.EventHandler(Btn_FontColor_Click);
		this.Txt_LeftIndent.Location = new System.Drawing.Point(82, 93);
		this.Txt_LeftIndent.Name = "Txt_LeftIndent";
		this.Txt_LeftIndent.Size = new System.Drawing.Size(100, 25);
		this.Txt_LeftIndent.TabIndex = 38;
		this.Txt_LeftIndent.TextChanged += new System.EventHandler(Txt_Indent_TextChanged);
		this.Txt_LeftIndent.Validated += new System.EventHandler(Txt_Indent_Validated);
		this.Btn_BreakBefore.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_BreakBefore.Location = new System.Drawing.Point(482, 124);
		this.Btn_BreakBefore.Name = "Btn_BreakBefore";
		this.Btn_BreakBefore.Pressed = false;
		this.Btn_BreakBefore.Size = new System.Drawing.Size(40, 30);
		this.Btn_BreakBefore.TabIndex = 34;
		this.Btn_BreakBefore.Text = "否";
		this.Btn_BreakBefore.UseVisualStyleBackColor = false;
		this.Btn_BreakBefore.PressedChanged += new System.EventHandler(ToggleButton_PressedChanged);
		this.Cmb_HAlignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_HAlignment.FormattingEnabled = true;
		this.Cmb_HAlignment.Location = new System.Drawing.Point(82, 160);
		this.Cmb_HAlignment.Name = "Cmb_HAlignment";
		this.Cmb_HAlignment.Size = new System.Drawing.Size(100, 26);
		this.Cmb_HAlignment.TabIndex = 30;
		this.Cmb_HAlignment.SelectedIndexChanged += new System.EventHandler(Cmb_FontNameAndHV_SelectedIndexChanged);
		this.Cmb_SpaceBefore.FormattingEnabled = true;
		this.Cmb_SpaceBefore.Location = new System.Drawing.Point(82, 126);
		this.Cmb_SpaceBefore.Name = "Cmb_SpaceBefore";
		this.Cmb_SpaceBefore.Size = new System.Drawing.Size(100, 26);
		this.Cmb_SpaceBefore.TabIndex = 26;
		this.Cmb_SpaceBefore.TextChanged += new System.EventHandler(Cmb_SpaceValue_TextChanged);
		this.Cmb_SpaceBefore.Validated += new System.EventHandler(Cmb_SpaceValue_Validated);
		this.Cmb_FontSize.FormattingEnabled = true;
		this.Cmb_FontSize.Location = new System.Drawing.Point(82, 58);
		this.Cmb_FontSize.Name = "Cmb_FontSize";
		this.Cmb_FontSize.Size = new System.Drawing.Size(100, 26);
		this.Cmb_FontSize.TabIndex = 24;
		this.Cmb_FontSize.TextChanged += new System.EventHandler(Cmb_FontSize_TextChanged);
		this.Cmb_FontSize.Validated += new System.EventHandler(Cmb_FontSize_Validated);
		this.Cmb_LineSpace.FormattingEnabled = true;
		this.Cmb_LineSpace.Location = new System.Drawing.Point(482, 92);
		this.Cmb_LineSpace.Name = "Cmb_LineSpace";
		this.Cmb_LineSpace.Size = new System.Drawing.Size(100, 26);
		this.Cmb_LineSpace.TabIndex = 23;
		this.Cmb_LineSpace.TextChanged += new System.EventHandler(Cmb_LineSpace_TextChanged);
		this.Cmb_LineSpace.Validated += new System.EventHandler(Cmb_LineSpace_Validated);
		this.Cmb_EngFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_EngFontName.FormattingEnabled = true;
		this.Cmb_EngFontName.Location = new System.Drawing.Point(382, 24);
		this.Cmb_EngFontName.Name = "Cmb_EngFontName";
		this.Cmb_EngFontName.Size = new System.Drawing.Size(200, 26);
		this.Cmb_EngFontName.TabIndex = 15;
		this.Cmb_EngFontName.SelectedIndexChanged += new System.EventHandler(Cmb_FontNameAndHV_SelectedIndexChanged);
		this.Cmb_ChnFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_ChnFontName.FormattingEnabled = true;
		this.Cmb_ChnFontName.Location = new System.Drawing.Point(82, 24);
		this.Cmb_ChnFontName.Name = "Cmb_ChnFontName";
		this.Cmb_ChnFontName.Size = new System.Drawing.Size(200, 26);
		this.Cmb_ChnFontName.TabIndex = 14;
		this.Cmb_ChnFontName.SelectedIndexChanged += new System.EventHandler(Cmb_FontNameAndHV_SelectedIndexChanged);
		this.Btn_Underline.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_Underline.Location = new System.Drawing.Point(542, 56);
		this.Btn_Underline.Name = "Btn_Underline";
		this.Btn_Underline.Pressed = false;
		this.Btn_Underline.Size = new System.Drawing.Size(40, 30);
		this.Btn_Underline.TabIndex = 5;
		this.Btn_Underline.Text = "否";
		this.Btn_Underline.UseVisualStyleBackColor = false;
		this.Btn_Underline.PressedChanged += new System.EventHandler(ToggleButton_PressedChanged);
		this.Btn_Italic.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_Italic.Location = new System.Drawing.Point(447, 56);
		this.Btn_Italic.Name = "Btn_Italic";
		this.Btn_Italic.Pressed = false;
		this.Btn_Italic.Size = new System.Drawing.Size(40, 30);
		this.Btn_Italic.TabIndex = 4;
		this.Btn_Italic.Text = "否";
		this.Btn_Italic.UseVisualStyleBackColor = false;
		this.Btn_Italic.PressedChanged += new System.EventHandler(ToggleButton_PressedChanged);
		this.Btn_Bold.BackColor = System.Drawing.Color.AliceBlue;
		this.Btn_Bold.Location = new System.Drawing.Point(359, 56);
		this.Btn_Bold.Name = "Btn_Bold";
		this.Btn_Bold.Pressed = false;
		this.Btn_Bold.Size = new System.Drawing.Size(40, 30);
		this.Btn_Bold.TabIndex = 3;
		this.Btn_Bold.Text = "否";
		this.Btn_Bold.UseVisualStyleBackColor = false;
		this.Btn_Bold.PressedChanged += new System.EventHandler(ToggleButton_PressedChanged);
		this.Btn_SetStyles.Location = new System.Drawing.Point(400, 463);
		this.Btn_SetStyles.Name = "Btn_SetStyles";
		this.Btn_SetStyles.Size = new System.Drawing.Size(186, 30);
		this.Btn_SetStyles.TabIndex = 2;
		this.Btn_SetStyles.Text = "确定";
		this.Btn_SetStyles.UseVisualStyleBackColor = true;
		this.Btn_SetStyles.Click += new System.EventHandler(Btn_SetStyles_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(96f, 96f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.ClientSize = new System.Drawing.Size(589, 496);
		base.Controls.Add(this.Btn_SetStyles);
		base.Controls.Add(this.Grp_SetSelectedStyle);
		base.Controls.Add(this.Dta_StyleList);
		this.DoubleBuffered = true;
		this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "LevelStyleSettingUI";
		this.Text = "多级段落样式设置";
		((System.ComponentModel.ISupportInitialize)this.Dta_StyleList).EndInit();
		this.Grp_SetSelectedStyle.ResumeLayout(false);
		this.Grp_SetSelectedStyle.PerformLayout();
		base.ResumeLayout(false);
	}
}
