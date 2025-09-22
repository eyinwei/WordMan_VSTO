using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class PageSetUI : UserControl
{
	private readonly List<string> FontSizeCha = new List<string>(16)
	{
		"八号", "七号", "小六", "六号", "小五", "五号", "小四", "四号", "小三", "三号",
		"小二", "二号", "小一", "一号", "小初", "初号"
	};

	private readonly List<float> FontSizePoint = new List<float>(16)
	{
		5f, 5.5f, 6.5f, 7.5f, 9f, 10.5f, 12f, 14f, 15f, 16f,
		18f, 22f, 24f, 26f, 36f, 42f
	};

	private static HeaderFooterFontInfo fontInfo = new HeaderFooterFontInfo();

	private static HeaderFooterTextInfo textInfo = new HeaderFooterTextInfo();

	private bool NotUserChanged;

	private static readonly List<string> historyInput = new List<string>();

	private IContainer components;

	private GroupBox groupBox3;

	private LineTypeSelectComboBox CmbFooterlineTypeSelect;

	private LineTypeSelectComboBox CmbHeaderlineTypeSelect;

	private ComboBox CmbLOGOSize;

	private Label LabAddLogo;

	private Label label12;

	private ComboBox CmbHeaderFooterFontSize;

	private Label label11;

	private NumericUpDown NumUpDownPageStart;

	private NumericUpDownWithUnit NumUpDownFooterHeight;

	private NumericUpDownWithUnit NumUpDownHeaderHeight;

	private Label label9;

	private Label label8;

	private ComboBox CmbFooterRightText;

	private ComboBox CmbHeaderRightText;

	private ComboBox CmbFooterSelect;

	private ComboBox CmbFooterLeftText;

	private ComboBox CmbHeaderSelect;

	private ComboBox CmbHeaderLeftText;

	private CheckBox ChkOddEvenPageDiffrent;

	private CheckBox ChkFirstPageDiffrent;

	private ComboBox CmbFooterMiddleText;

	private ComboBox CmbHeaderMiddleText;

	private Button BtnSetHeaderFooter;

	private PictureBox PicBoxLOGOFile;

	private Label label18;

	private Label label17;

	private Label label16;

	private Label label15;

	private Label label14;

	private GroupBox groupBox2;

	private ComboBox CmbPageMarginType;

	private CheckBox ChkSetBookbinding;

	private CheckBox ChkSetPageMargin;

	private CheckBox ChkApplySectionMargin;

	private Button BtnApplyPageMargin;

	private ComboBox CmbBookbinding;

	private NumericUpDownWithUnit Num_TopMargin;

	private NumericUpDownWithUnit NumUpDownBookbinding;

	private OpenFileDialog DlgOpenFiles;

	private CheckBox ChkRestartAtSection;

	private CheckBox ChkSameHeaderFooterHeight;

	private RadioButton RdoApplyToSection;

	private RadioButton RdoApplyToSectionEnd;

	private RadioButton RdoApplyToDocument;

	internal Label LabCurrentSectionNo;

	private ToggleButton TogItalic;

	private ToggleButton TogBold;

	private ComboBox CmbHeaderFooterFont;

	private ComboBox CmbHeaderFooterFontName;

	private ComboBox CmbLogoIndex;

	private Label label1;

	private CheckBox ChkClearCurrent;

	private Label label5;

	private NumericUpDownWithUnit Num_RightMargin;

	private Label label4;

	private NumericUpDownWithUnit Num_LeftMargin;

	private Label label3;

	private NumericUpDownWithUnit Num_BottomMargin;

	private Label label2;

	protected override CreateParams CreateParams
	{
		get
		{
			CreateParams obj = base.CreateParams;
			obj.ExStyle |= 33554432;
			return obj;
		}
	}

	public PageSetUI()
	{
		InitializeComponent();
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		Section first = application.Selection.Sections.First;
		float num = application.PointsToCentimeters(first.PageSetup.PageHeight);
		NotUserChanged = true;
		CmbBookbinding.SelectedIndex = 0;
		if (first.PageSetup.MirrorMargins == -1)
		{
			CmbBookbinding.SelectedIndex = 2;
		}
		else if (first.PageSetup.GutterPos == WdGutterStyle.wdGutterPosLeft)
		{
			CmbBookbinding.SelectedIndex = 0;
		}
		else if (first.PageSetup.GutterPos == WdGutterStyle.wdGutterPosTop)
		{
			CmbBookbinding.SelectedIndex = 1;
		}
		else
		{
			CmbBookbinding.SelectedIndex = -1;
		}
		CmbHeaderSelect.SelectedIndex = 0;
		CmbFooterSelect.SelectedIndex = 0;
		Num_TopMargin.Maximum = (decimal)num;
		Num_BottomMargin.Maximum = (decimal)num;
		Num_LeftMargin.Maximum = (decimal)num;
		Num_RightMargin.Maximum = (decimal)num;
		Num_TopMargin.Value = (decimal)application.PointsToCentimeters(first.PageSetup.TopMargin);
		Num_BottomMargin.Value = (decimal)application.PointsToCentimeters(first.PageSetup.BottomMargin);
		Num_LeftMargin.Value = (decimal)application.PointsToCentimeters(first.PageSetup.LeftMargin);
		Num_RightMargin.Value = (decimal)application.PointsToCentimeters(first.PageSetup.RightMargin);
		CmbPageMarginType.SelectedIndex = 0;
		NumUpDownBookbinding.Maximum = (decimal)num;
		NumUpDownBookbinding.Value = (decimal)application.PointsToCentimeters(first.PageSetup.Gutter);
		NumUpDownHeaderHeight.Maximum = (decimal)num;
		NumUpDownFooterHeight.Maximum = (decimal)num;
		NumUpDownHeaderHeight.Value = (decimal)application.PointsToCentimeters(first.PageSetup.HeaderDistance);
		NumUpDownFooterHeight.Value = (decimal)application.PointsToCentimeters(first.PageSetup.FooterDistance);
		ComboBox.ObjectCollection items = CmbHeaderFooterFontSize.Items;
		List<string> fontSizeCha = FontSizeCha;
		int num2 = 0;
		object[] array = new object[fontSizeCha.Count];
		foreach (string item in fontSizeCha)
		{
			array[num2] = item;
			num2++;
		}
		items.AddRange(array);
		FontFamily[] families = new InstalledFontCollection().Families;
		CmbHeaderFooterFontName.Items.AddRange(((IEnumerable<object>)families.Select((FontFamily item) => item.Name)).ToArray());
		ChkFirstPageDiffrent.Checked = textInfo.FirstPageDiffrent;
		ChkOddEvenPageDiffrent.Checked = textInfo.OddEvenPageDiffrent;
		if (textInfo.ApplyModel == 0)
		{
			RdoApplyToSection.Checked = true;
		}
		else if (textInfo.ApplyModel == 1)
		{
			RdoApplyToDocument.Checked = true;
		}
		else
		{
			RdoApplyToSectionEnd.Checked = true;
		}
		ChkRestartAtSection.Checked = textInfo.PageNumberStartAtSection;
		NumUpDownPageStart.Value = textInfo.StartNumber;
		CmbHeaderSelect.SelectedIndex = 0;
		CmbFooterSelect.SelectedIndex = 0;
		CmbHeaderlineTypeSelect.SelectedIndex = textInfo.HeaderLineType;
		CmbFooterlineTypeSelect.SelectedIndex = textInfo.FooterLineType;
		CmbLogoIndex.SelectedIndex = 0;
		CmbLOGOSize.Text = textInfo.LogoHeight + "倍字高";
		if (File.Exists(textInfo.LogoPath[CmbLogoIndex.SelectedIndex]))
		{
			PicBoxLOGOFile.Image = Image.FromFile(textInfo.LogoPath[CmbLogoIndex.SelectedIndex]);
			LabAddLogo.Visible = false;
		}
		if (fontInfo.HeaderFontName == string.Empty)
		{
			Styles styles = Globals.ThisAddIn.Application.ActiveDocument.Styles;
			object Index = WdBuiltinStyle.wdStyleHeader;
			Style style = styles[ref Index];
			fontInfo.HeaderFontName = style.Font.Name;
			fontInfo.HeaderFontSize = style.Font.Size;
			fontInfo.HeaderFontBold = style.Font.Bold == -1;
			fontInfo.HeaderFontItalic = style.Font.Italic == -1;
		}
		if (fontInfo.FooterFontName == string.Empty)
		{
			Styles styles2 = Globals.ThisAddIn.Application.ActiveDocument.Styles;
			object Index = WdBuiltinStyle.wdStyleFooter;
			Style style2 = styles2[ref Index];
			fontInfo.FooterFontName = style2.Font.Name;
			fontInfo.FooterFontSize = style2.Font.Size;
			fontInfo.FooterFontBold = style2.Font.Bold == -1;
			fontInfo.FooterFontItalic = style2.Font.Italic == -1;
		}
		CmbHeaderFooterFont.SelectedIndex = 0;
		AddHistoryInputItems();
		int index = Globals.ThisAddIn.Application.Selection.Sections.First.Index;
		int count = Globals.ThisAddIn.Application.ActiveDocument.Sections.Count;
		LabCurrentSectionNo.Text = "当前第" + index + "节，全文共" + count + "节。";
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowSelectionChange").AddEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(UpdataSectionNo));
		NotUserChanged = false;
	}

	private void AddHistoryInputItems([Optional] List<string> newItems)
	{
		if (newItems != null)
		{
			if (newItems.Count <= 0)
			{
				return;
			}
			ComboBox.ObjectCollection items = CmbHeaderLeftText.Items;
			int num = 0;
			object[] array = new object[newItems.Count];
			foreach (string newItem in newItems)
			{
				array[num] = newItem;
				num++;
			}
			items.AddRange(array);
			items = CmbHeaderMiddleText.Items;
			num = 0;
			array = new object[newItems.Count];
			foreach (string newItem2 in newItems)
			{
				array[num] = newItem2;
				num++;
			}
			items.AddRange(array);
			items = CmbHeaderRightText.Items;
			num = 0;
			array = new object[newItems.Count];
			foreach (string newItem3 in newItems)
			{
				array[num] = newItem3;
				num++;
			}
			items.AddRange(array);
			items = CmbFooterLeftText.Items;
			num = 0;
			array = new object[newItems.Count];
			foreach (string newItem4 in newItems)
			{
				array[num] = newItem4;
				num++;
			}
			items.AddRange(array);
			items = CmbFooterMiddleText.Items;
			num = 0;
			array = new object[newItems.Count];
			foreach (string newItem5 in newItems)
			{
				array[num] = newItem5;
				num++;
			}
			items.AddRange(array);
			items = CmbFooterRightText.Items;
			num = 0;
			array = new object[newItems.Count];
			foreach (string newItem6 in newItems)
			{
				array[num] = newItem6;
				num++;
			}
			items.AddRange(array);
		}
		else
		{
			if (historyInput.Count <= 0)
			{
				return;
			}
			ComboBox.ObjectCollection items = CmbHeaderLeftText.Items;
			List<string> list = historyInput;
			int num = 0;
			object[] array = new object[list.Count];
			foreach (string item in list)
			{
				array[num] = item;
				num++;
			}
			items.AddRange(array);
			items = CmbHeaderMiddleText.Items;
			List<string> list2 = historyInput;
			num = 0;
			array = new object[list2.Count];
			foreach (string item2 in list2)
			{
				array[num] = item2;
				num++;
			}
			items.AddRange(array);
			items = CmbHeaderRightText.Items;
			List<string> list3 = historyInput;
			num = 0;
			array = new object[list3.Count];
			foreach (string item3 in list3)
			{
				array[num] = item3;
				num++;
			}
			items.AddRange(array);
			items = CmbFooterLeftText.Items;
			List<string> list4 = historyInput;
			num = 0;
			array = new object[list4.Count];
			foreach (string item4 in list4)
			{
				array[num] = item4;
				num++;
			}
			items.AddRange(array);
			items = CmbFooterMiddleText.Items;
			List<string> list5 = historyInput;
			num = 0;
			array = new object[list5.Count];
			foreach (string item5 in list5)
			{
				array[num] = item5;
				num++;
			}
			items.AddRange(array);
			items = CmbFooterRightText.Items;
			List<string> list6 = historyInput;
			num = 0;
			array = new object[list6.Count];
			foreach (string item6 in list6)
			{
				array[num] = item6;
				num++;
			}
			items.AddRange(array);
		}
	}

	private void BtnApplyPageMargin_Click(object sender, EventArgs e)
	{
		float[] pageMargin = new float[4]
		{
			(float)Num_TopMargin.Value,
			(float)Num_BottomMargin.Value,
			(float)Num_LeftMargin.Value,
			(float)Num_RightMargin.Value
		};
		Globals.ThisAddIn.ApplyPageMargin(Globals.ThisAddIn.Application.ActiveDocument, ChkApplySectionMargin.Checked, ChkSetPageMargin.Checked, pageMargin, ChkSetBookbinding.Checked, CmbBookbinding.SelectedIndex, (float)NumUpDownBookbinding.Value);
	}

	private void BtnSetHeaderFooter_Click(object sender, EventArgs e)
	{
		float num = FontSizePoint[FontSizeCha.IndexOf(CmbHeaderFooterFontSize.Text)];
		Convert.ToSingle(CmbLOGOSize.Text.Replace("倍字高", ""));
		foreach (string item in textInfo.PrimaryHeaderText.Concat(textInfo.EvenHeaderText.Concat(textInfo.FirstHeaderText)).Concat(textInfo.PrimaryFooterText.Concat(textInfo.EvenFooterText.Concat(textInfo.FirstFooterText))))
		{
			if (item.Contains("[LOGO1]") && textInfo.LogoPath[0] == "")
			{
				if (MessageBox.Show("未指定Logo1文件的路径！是否选择文件？取消则退出本次设置。", "提醒", MessageBoxButtons.OKCancel) != DialogResult.OK)
				{
					return;
				}
				CmbLogoIndex.SelectedIndex = 0;
				PicBoxLOGOFile_DoubleClick(null, null);
			}
			if (item.Contains("[LOGO2]") && textInfo.LogoPath[1] == "")
			{
				if (MessageBox.Show("未指定Logo2文件的路径！是否选择文件？取消则退出本次设置。", "提醒", MessageBoxButtons.OKCancel) != DialogResult.OK)
				{
					return;
				}
				CmbLogoIndex.SelectedIndex = 1;
				PicBoxLOGOFile_DoubleClick(null, null);
			}
			if (item.Contains("[LOGO3]") && textInfo.LogoPath[2] == "")
			{
				if (MessageBox.Show("未指定Logo3文件的路径！是否选择文件？取消则退出本次设置。", "提醒", MessageBoxButtons.OKCancel) != DialogResult.OK)
				{
					return;
				}
				CmbLogoIndex.SelectedIndex = 2;
				PicBoxLOGOFile_DoubleClick(null, null);
			}
		}
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowSelectionChange").RemoveEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(UpdataSectionNo));
		Globals.ThisAddIn.SetHeaderFooter(textInfo, (float)NumUpDownHeaderHeight.Value, (float)NumUpDownFooterHeight.Value, fontInfo, ChkClearCurrent.Checked, LabCurrentSectionNo);
		List<ComboBox> obj = new List<ComboBox>(6) { CmbHeaderLeftText, CmbHeaderMiddleText, CmbHeaderRightText, CmbFooterLeftText, CmbFooterMiddleText, CmbFooterRightText };
		List<string> list = new List<string>();
		foreach (ComboBox item2 in obj)
		{
			if (item2.SelectedIndex == -1)
			{
				string text = item2.Text.Trim();
				if (!historyInput.Contains(text) && text != "")
				{
					list.Add(text);
					historyInput.Add(text);
				}
			}
		}
		AddHistoryInputItems(list);
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowSelectionChange").AddEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(UpdataSectionNo));
	}

	private void ChkFirstPageDiffrent_CheckedChanged(object sender, EventArgs e)
	{
		textInfo.FirstPageDiffrent = ChkFirstPageDiffrent.Checked;
		if (ChkFirstPageDiffrent.Checked)
		{
			CmbHeaderSelect.Items.Add("首页页眉");
			CmbFooterSelect.Items.Add("首页页脚");
			return;
		}
		CmbHeaderSelect.Items.Remove("首页页眉");
		CmbFooterSelect.Items.Remove("首页页脚");
		CmbHeaderSelect.SelectedIndex = 0;
		CmbFooterSelect.SelectedIndex = 0;
	}

	private void ChkOddEvenPageDiffrent_CheckedChanged(object sender, EventArgs e)
	{
		textInfo.OddEvenPageDiffrent = ChkOddEvenPageDiffrent.Checked;
		if (ChkOddEvenPageDiffrent.Checked)
		{
			CmbHeaderSelect.Items.Remove("页眉");
			CmbHeaderSelect.Items.Add("奇数页眉");
			CmbHeaderSelect.Items.Add("偶数页眉");
			CmbFooterSelect.Items.Remove("页脚");
			CmbFooterSelect.Items.Add("奇数页脚");
			CmbFooterSelect.Items.Add("偶数页脚");
			CmbHeaderSelect.SelectedIndex = 0;
			CmbFooterSelect.SelectedIndex = 0;
		}
		else
		{
			CmbHeaderSelect.Items.Add("页眉");
			CmbHeaderSelect.Items.Remove("奇数页眉");
			CmbHeaderSelect.Items.Remove("偶数页眉");
			CmbFooterSelect.Items.Add("页脚");
			CmbFooterSelect.Items.Remove("奇数页脚");
			CmbFooterSelect.Items.Remove("偶数页脚");
			CmbHeaderSelect.SelectedIndex = 0;
			CmbFooterSelect.SelectedIndex = 0;
		}
	}

	private void CmbHeaderFooter_Validated(object sender, EventArgs e)
	{
		ComboBox comboBox = sender as ComboBox;
		if (Regex.IsMatch(comboBox.Text, "^[ ]{1,}$"))
		{
			comboBox.Text = "";
		}
		string text;
		if (comboBox.Name.Contains("Header"))
		{
			text = CmbHeaderSelect.Text;
			if (!(text == "首页页眉"))
			{
				if (text == "偶数页眉")
				{
					textInfo.EvenHeaderText[0] = CmbHeaderLeftText.Text;
					textInfo.EvenHeaderText[1] = CmbHeaderMiddleText.Text;
					textInfo.EvenHeaderText[2] = CmbHeaderRightText.Text;
				}
				else
				{
					textInfo.PrimaryHeaderText[0] = CmbHeaderLeftText.Text;
					textInfo.PrimaryHeaderText[1] = CmbHeaderMiddleText.Text;
					textInfo.PrimaryHeaderText[2] = CmbHeaderRightText.Text;
				}
			}
			else
			{
				textInfo.FirstHeaderText[0] = CmbHeaderLeftText.Text;
				textInfo.FirstHeaderText[1] = CmbHeaderMiddleText.Text;
				textInfo.FirstHeaderText[2] = CmbHeaderRightText.Text;
			}
			return;
		}
		text = CmbFooterSelect.Text;
		if (!(text == "首页页脚"))
		{
			if (text == "偶数页脚")
			{
				textInfo.EvenFooterText[0] = CmbFooterLeftText.Text;
				textInfo.EvenFooterText[1] = CmbFooterMiddleText.Text;
				textInfo.EvenFooterText[2] = CmbFooterRightText.Text;
			}
			else
			{
				textInfo.PrimaryFooterText[0] = CmbFooterLeftText.Text;
				textInfo.PrimaryFooterText[1] = CmbFooterMiddleText.Text;
				textInfo.PrimaryFooterText[2] = CmbFooterRightText.Text;
			}
		}
		else
		{
			textInfo.FirstFooterText[0] = CmbFooterLeftText.Text;
			textInfo.FirstFooterText[1] = CmbFooterMiddleText.Text;
			textInfo.FirstFooterText[2] = CmbFooterRightText.Text;
		}
	}

	private void CmbHeaderSelect_SelectedIndexChanged(object sender, EventArgs e)
	{
		string text = CmbHeaderSelect.Text;
		if (!(text == "偶数页眉"))
		{
			if (text == "首页页眉")
			{
				CmbHeaderLeftText.Text = textInfo.FirstHeaderText[0];
				CmbHeaderMiddleText.Text = textInfo.FirstHeaderText[1];
				CmbHeaderRightText.Text = textInfo.FirstHeaderText[2];
			}
			else
			{
				CmbHeaderLeftText.Text = textInfo.PrimaryHeaderText[0];
				CmbHeaderMiddleText.Text = textInfo.PrimaryHeaderText[1];
				CmbHeaderRightText.Text = textInfo.PrimaryHeaderText[2];
			}
		}
		else
		{
			CmbHeaderLeftText.Text = textInfo.EvenHeaderText[0];
			CmbHeaderMiddleText.Text = textInfo.EvenHeaderText[1];
			CmbHeaderRightText.Text = textInfo.EvenHeaderText[2];
		}
	}

	private void CmbFooterSelect_SelectedIndexChanged(object sender, EventArgs e)
	{
		string text = CmbFooterSelect.Text;
		if (!(text == "偶数页脚"))
		{
			if (text == "首页页脚")
			{
				CmbFooterLeftText.Text = textInfo.FirstFooterText[0];
				CmbFooterMiddleText.Text = textInfo.FirstFooterText[1];
				CmbFooterRightText.Text = textInfo.FirstFooterText[2];
			}
			else
			{
				CmbFooterLeftText.Text = textInfo.PrimaryFooterText[0];
				CmbFooterMiddleText.Text = textInfo.PrimaryFooterText[1];
				CmbFooterRightText.Text = textInfo.PrimaryFooterText[2];
			}
		}
		else
		{
			CmbFooterLeftText.Text = textInfo.EvenFooterText[0];
			CmbFooterMiddleText.Text = textInfo.EvenFooterText[1];
			CmbFooterRightText.Text = textInfo.EvenFooterText[2];
		}
	}

	private void ChkRestartAtSection_CheckedChanged(object sender, EventArgs e)
	{
		textInfo.PageNumberStartAtSection = ChkRestartAtSection.Checked;
		NumUpDownPageStart.Enabled = ChkRestartAtSection.Checked;
	}

	private void PicBoxLOGOFile_DoubleClick(object sender, EventArgs e)
	{
		DlgOpenFiles.Filter = "图片文件（BMP、JPG、GIF、PNG、ICO）|*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.ico";
		DlgOpenFiles.Title = "选择一个图片文件";
		DlgOpenFiles.FileName = "LogoFile";
		if (DlgOpenFiles.ShowDialog() == DialogResult.OK)
		{
			PicBoxLOGOFile.Image = Image.FromFile(DlgOpenFiles.FileName);
			textInfo.LogoPath[CmbLogoIndex.SelectedIndex] = DlgOpenFiles.FileName;
			LabAddLogo.Visible = false;
		}
	}

	private void ChkSetPageMargin_CheckedChanged(object sender, EventArgs e)
	{
		CmbPageMarginType.Enabled = ChkSetPageMargin.Checked;
		Num_TopMargin.Enabled = ChkSetPageMargin.Checked;
		Num_BottomMargin.Enabled = ChkSetPageMargin.Checked;
		Num_LeftMargin.Enabled = ChkSetPageMargin.Checked;
		Num_RightMargin.Enabled = ChkSetPageMargin.Checked;
	}

	private void CmbPageMarginType_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!NotUserChanged)
		{
			if (CmbPageMarginType.SelectedIndex == 1)
			{
				Num_BottomMargin.Enabled = false;
				Num_LeftMargin.Enabled = false;
				Num_RightMargin.Enabled = false;
				Num_BottomMargin.Value = Num_TopMargin.Value;
				Num_LeftMargin.Value = Num_TopMargin.Value;
				Num_RightMargin.Value = Num_TopMargin.Value;
			}
			else if (CmbPageMarginType.SelectedIndex == 2)
			{
				Num_BottomMargin.Enabled = true;
				Num_LeftMargin.Enabled = true;
				Num_RightMargin.Enabled = false;
				Num_RightMargin.Value = Num_LeftMargin.Value;
			}
			else if (CmbPageMarginType.SelectedIndex == 3)
			{
				Num_BottomMargin.Enabled = false;
				Num_LeftMargin.Enabled = true;
				Num_RightMargin.Enabled = true;
				Num_BottomMargin.Value = Num_TopMargin.Value;
			}
			else
			{
				Num_BottomMargin.Enabled = true;
				Num_LeftMargin.Enabled = true;
				Num_RightMargin.Enabled = true;
			}
		}
	}

	private void Num_TopMargin_ValueChanged(object sender, EventArgs e)
	{
		if (!NotUserChanged)
		{
			if (CmbPageMarginType.SelectedIndex == 1)
			{
				Num_BottomMargin.Value = Num_TopMargin.Value;
				Num_LeftMargin.Value = Num_TopMargin.Value;
				Num_RightMargin.Value = Num_TopMargin.Value;
			}
			else if (CmbPageMarginType.SelectedIndex == 3)
			{
				Num_BottomMargin.Value = Num_TopMargin.Value;
			}
		}
	}

	private void Num_LeftMargin_ValueChanged(object sender, EventArgs e)
	{
		if (!NotUserChanged && CmbPageMarginType.SelectedIndex == 2)
		{
			Num_RightMargin.Value = Num_LeftMargin.Value;
		}
	}

	private void ChkSetBookbinding_CheckedChanged(object sender, EventArgs e)
	{
		CmbBookbinding.Enabled = ChkSetBookbinding.Checked;
		NumUpDownBookbinding.Enabled = ChkSetBookbinding.Checked;
	}

	private void UpdataSectionNo(Selection selection)
	{
		int index = selection.Sections.First.Index;
		int count = Globals.ThisAddIn.Application.ActiveDocument.Sections.Count;
		Section section = Globals.ThisAddIn.Application.ActiveDocument.Sections[index];
		LabCurrentSectionNo.Text = "当前第" + index + "节，全文共" + count + "节。";
		if (selection.StoryType == WdStoryType.wdMainTextStory)
		{
			NumUpDownHeaderHeight.Value = (decimal)Globals.ThisAddIn.Application.PointsToCentimeters(section.PageSetup.HeaderDistance);
			NumUpDownFooterHeight.Value = (decimal)Globals.ThisAddIn.Application.PointsToCentimeters(section.PageSetup.FooterDistance);
		}
	}

	internal void DeleteUpdataSectionNo(object sender, EventArgs e)
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowSelectionChange").RemoveEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(UpdataSectionNo));
	}

	private void ChkSameHeaderFooterHeight_CheckedChanged(object sender, EventArgs e)
	{
		textInfo.SameHeaderFooterHeight = ChkSameHeaderFooterHeight.Checked;
		NumUpDownFooterHeight.Enabled = ChkSameHeaderFooterHeight.Checked;
		NumUpDownHeaderHeight.Enabled = ChkSameHeaderFooterHeight.Checked;
	}

	private void CmbHeaderFooterFont_SelectedIndexChanged(object sender, EventArgs e)
	{
		NotUserChanged = true;
		if (CmbHeaderFooterFont.SelectedIndex == 1)
		{
			CmbHeaderFooterFontName.SelectedIndex = CmbHeaderFooterFontName.Items.IndexOf(fontInfo.FooterFontName);
			int num = FontSizePoint.IndexOf(fontInfo.FooterFontSize);
			if (num != -1)
			{
				CmbHeaderFooterFontSize.Text = null;
				CmbHeaderFooterFontSize.SelectedIndex = num;
			}
			else
			{
				CmbHeaderFooterFontSize.Text = fontInfo.FooterFontSize.ToString();
			}
			TogBold.Pressed = fontInfo.FooterFontBold;
			TogItalic.Pressed = fontInfo.FooterFontItalic;
		}
		else
		{
			CmbHeaderFooterFontName.SelectedIndex = CmbHeaderFooterFontName.Items.IndexOf(fontInfo.HeaderFontName);
			int num2 = FontSizePoint.IndexOf(fontInfo.HeaderFontSize);
			if (num2 != -1)
			{
				CmbHeaderFooterFontSize.Text = null;
				CmbHeaderFooterFontSize.SelectedIndex = num2;
			}
			else
			{
				CmbHeaderFooterFontSize.Text = fontInfo.HeaderFontSize.ToString();
			}
			TogBold.Pressed = fontInfo.HeaderFontBold;
			TogItalic.Pressed = fontInfo.HeaderFontItalic;
		}
		NotUserChanged = false;
	}

	private void CmbHeaderFooterFontName_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!NotUserChanged)
		{
			if (CmbHeaderFooterFont.SelectedIndex == 0)
			{
				fontInfo.HeaderFontName = CmbHeaderFooterFontName.SelectedItem as string;
			}
			else
			{
				fontInfo.FooterFontName = CmbHeaderFooterFontName.SelectedItem as string;
			}
		}
	}

	private void TogBold_Click(object sender, EventArgs e)
	{
		if (NotUserChanged)
		{
			return;
		}
		string name = (sender as ToggleButton).Name;
		if (!(name == "TogBold"))
		{
			if (name == "TogItalic")
			{
				if (CmbHeaderFooterFont.SelectedIndex == 0)
				{
					fontInfo.HeaderFontItalic = (sender as ToggleButton).Pressed;
				}
				else
				{
					fontInfo.FooterFontItalic = (sender as ToggleButton).Pressed;
				}
			}
		}
		else if (CmbHeaderFooterFont.SelectedIndex == 0)
		{
			fontInfo.HeaderFontBold = (sender as ToggleButton).Pressed;
		}
		else
		{
			fontInfo.FooterFontBold = (sender as ToggleButton).Pressed;
		}
	}

	private void CmbHeaderFooterFontSize_Leave(object sender, EventArgs e)
	{
		if (NotUserChanged || CmbHeaderFooterFontSize.SelectedIndex != -1 || !Regex.IsMatch(CmbHeaderFooterFontSize.Text, "^[1-9]{1,4}(\\.5|\\.0){0,1}$"))
		{
			return;
		}
		float num = Convert.ToSingle(CmbHeaderFooterFontSize.Text);
		if (num >= 1f || num <= 1638f)
		{
			if (CmbHeaderFooterFont.SelectedIndex == 0)
			{
				fontInfo.HeaderFontSize = num;
			}
			else
			{
				fontInfo.FooterFontSize = num;
			}
		}
	}

	private void CmbHeaderFooterFontSize_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (!NotUserChanged)
		{
			if (CmbHeaderFooterFont.SelectedIndex == 0)
			{
				fontInfo.HeaderFontSize = FontSizePoint[CmbHeaderFooterFontSize.SelectedIndex];
			}
			else
			{
				fontInfo.FooterFontSize = FontSizePoint[CmbHeaderFooterFontSize.SelectedIndex];
			}
		}
	}

	private void RdoApplyToDocument_CheckedChanged(object sender, EventArgs e)
	{
		textInfo.ApplyModel = (sender as RadioButton).Name switch
		{
			"RdoApplyToSection" => 0, 
			"RdoApplyToDocument" => 1, 
			"RdoApplyToSectionEnd" => 2, 
			_ => 1, 
		};
	}

	private void CmbLOGOSize_SelectedIndexChanged(object sender, EventArgs e)
	{
		textInfo.LogoHeight = Convert.ToSingle(CmbLOGOSize.Text.Replace("倍字高", ""));
	}

	private void CmbHeaderlineTypeSelect_SelectedIndexChanged(object sender, EventArgs e)
	{
		if ((sender as LineTypeSelectComboBox).Name == "CmbHeaderlineTypeSelect")
		{
			textInfo.HeaderLineType = CmbHeaderlineTypeSelect.SelectedIndex;
		}
		else
		{
			textInfo.FooterLineType = CmbFooterlineTypeSelect.SelectedIndex;
		}
	}

	private void CmbLogoIndex_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (textInfo.LogoPath[CmbLogoIndex.SelectedIndex] == "")
		{
			PicBoxLOGOFile.Image = null;
			LabAddLogo.Visible = true;
		}
		else if (File.Exists(textInfo.LogoPath[CmbLogoIndex.SelectedIndex]))
		{
			PicBoxLOGOFile.Image = Image.FromFile(textInfo.LogoPath[CmbLogoIndex.SelectedIndex]);
			LabAddLogo.Visible = false;
		}
		else if (MessageBox.Show("指定的图片文件不存在，是否重新选择！", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
		{
			PicBoxLOGOFile_DoubleClick(null, null);
		}
		else
		{
			textInfo.LogoPath[CmbLogoIndex.SelectedIndex] = "";
			PicBoxLOGOFile.Image = null;
			LabAddLogo.Visible = true;
		}
	}

	private void NumUpDownPageStart_ValueChanged(object sender, EventArgs e)
	{
		textInfo.StartNumber = (int)NumUpDownPageStart.Value;
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
		this.groupBox3 = new System.Windows.Forms.GroupBox();
		this.ChkClearCurrent = new System.Windows.Forms.CheckBox();
		this.CmbLogoIndex = new System.Windows.Forms.ComboBox();
		this.label1 = new System.Windows.Forms.Label();
		this.CmbHeaderFooterFontName = new System.Windows.Forms.ComboBox();
		this.CmbHeaderFooterFont = new System.Windows.Forms.ComboBox();
		this.TogItalic = new WordFormatHelper.ToggleButton();
		this.TogBold = new WordFormatHelper.ToggleButton();
		this.RdoApplyToSection = new System.Windows.Forms.RadioButton();
		this.RdoApplyToSectionEnd = new System.Windows.Forms.RadioButton();
		this.RdoApplyToDocument = new System.Windows.Forms.RadioButton();
		this.ChkSameHeaderFooterHeight = new System.Windows.Forms.CheckBox();
		this.ChkRestartAtSection = new System.Windows.Forms.CheckBox();
		this.LabCurrentSectionNo = new System.Windows.Forms.Label();
		this.CmbFooterlineTypeSelect = new WordFormatHelper.LineTypeSelectComboBox();
		this.CmbHeaderlineTypeSelect = new WordFormatHelper.LineTypeSelectComboBox();
		this.CmbLOGOSize = new System.Windows.Forms.ComboBox();
		this.LabAddLogo = new System.Windows.Forms.Label();
		this.label12 = new System.Windows.Forms.Label();
		this.CmbHeaderFooterFontSize = new System.Windows.Forms.ComboBox();
		this.label11 = new System.Windows.Forms.Label();
		this.NumUpDownPageStart = new System.Windows.Forms.NumericUpDown();
		this.NumUpDownFooterHeight = new WordFormatHelper.NumericUpDownWithUnit();
		this.NumUpDownHeaderHeight = new WordFormatHelper.NumericUpDownWithUnit();
		this.label9 = new System.Windows.Forms.Label();
		this.label8 = new System.Windows.Forms.Label();
		this.CmbFooterRightText = new System.Windows.Forms.ComboBox();
		this.CmbHeaderRightText = new System.Windows.Forms.ComboBox();
		this.CmbFooterSelect = new System.Windows.Forms.ComboBox();
		this.CmbFooterLeftText = new System.Windows.Forms.ComboBox();
		this.CmbHeaderSelect = new System.Windows.Forms.ComboBox();
		this.CmbHeaderLeftText = new System.Windows.Forms.ComboBox();
		this.ChkOddEvenPageDiffrent = new System.Windows.Forms.CheckBox();
		this.ChkFirstPageDiffrent = new System.Windows.Forms.CheckBox();
		this.CmbFooterMiddleText = new System.Windows.Forms.ComboBox();
		this.CmbHeaderMiddleText = new System.Windows.Forms.ComboBox();
		this.BtnSetHeaderFooter = new System.Windows.Forms.Button();
		this.PicBoxLOGOFile = new System.Windows.Forms.PictureBox();
		this.label18 = new System.Windows.Forms.Label();
		this.label17 = new System.Windows.Forms.Label();
		this.label16 = new System.Windows.Forms.Label();
		this.label15 = new System.Windows.Forms.Label();
		this.label14 = new System.Windows.Forms.Label();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.label5 = new System.Windows.Forms.Label();
		this.Num_RightMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label4 = new System.Windows.Forms.Label();
		this.Num_LeftMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label3 = new System.Windows.Forms.Label();
		this.Num_BottomMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.label2 = new System.Windows.Forms.Label();
		this.CmbPageMarginType = new System.Windows.Forms.ComboBox();
		this.ChkSetBookbinding = new System.Windows.Forms.CheckBox();
		this.ChkSetPageMargin = new System.Windows.Forms.CheckBox();
		this.ChkApplySectionMargin = new System.Windows.Forms.CheckBox();
		this.BtnApplyPageMargin = new System.Windows.Forms.Button();
		this.CmbBookbinding = new System.Windows.Forms.ComboBox();
		this.Num_TopMargin = new WordFormatHelper.NumericUpDownWithUnit();
		this.NumUpDownBookbinding = new WordFormatHelper.NumericUpDownWithUnit();
		this.DlgOpenFiles = new System.Windows.Forms.OpenFileDialog();
		this.groupBox3.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPageStart).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownFooterHeight).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownHeaderHeight).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.PicBoxLOGOFile).BeginInit();
		this.groupBox2.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_RightMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_LeftMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BottomMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.Num_TopMargin).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownBookbinding).BeginInit();
		base.SuspendLayout();
		this.groupBox3.Controls.Add(this.ChkClearCurrent);
		this.groupBox3.Controls.Add(this.CmbLogoIndex);
		this.groupBox3.Controls.Add(this.label1);
		this.groupBox3.Controls.Add(this.CmbHeaderFooterFontName);
		this.groupBox3.Controls.Add(this.CmbHeaderFooterFont);
		this.groupBox3.Controls.Add(this.TogItalic);
		this.groupBox3.Controls.Add(this.TogBold);
		this.groupBox3.Controls.Add(this.RdoApplyToSection);
		this.groupBox3.Controls.Add(this.RdoApplyToSectionEnd);
		this.groupBox3.Controls.Add(this.RdoApplyToDocument);
		this.groupBox3.Controls.Add(this.ChkSameHeaderFooterHeight);
		this.groupBox3.Controls.Add(this.ChkRestartAtSection);
		this.groupBox3.Controls.Add(this.LabCurrentSectionNo);
		this.groupBox3.Controls.Add(this.CmbFooterlineTypeSelect);
		this.groupBox3.Controls.Add(this.CmbHeaderlineTypeSelect);
		this.groupBox3.Controls.Add(this.CmbLOGOSize);
		this.groupBox3.Controls.Add(this.LabAddLogo);
		this.groupBox3.Controls.Add(this.label12);
		this.groupBox3.Controls.Add(this.CmbHeaderFooterFontSize);
		this.groupBox3.Controls.Add(this.label11);
		this.groupBox3.Controls.Add(this.NumUpDownPageStart);
		this.groupBox3.Controls.Add(this.NumUpDownFooterHeight);
		this.groupBox3.Controls.Add(this.NumUpDownHeaderHeight);
		this.groupBox3.Controls.Add(this.label9);
		this.groupBox3.Controls.Add(this.label8);
		this.groupBox3.Controls.Add(this.CmbFooterRightText);
		this.groupBox3.Controls.Add(this.CmbHeaderRightText);
		this.groupBox3.Controls.Add(this.CmbFooterSelect);
		this.groupBox3.Controls.Add(this.CmbFooterLeftText);
		this.groupBox3.Controls.Add(this.CmbHeaderSelect);
		this.groupBox3.Controls.Add(this.CmbHeaderLeftText);
		this.groupBox3.Controls.Add(this.ChkOddEvenPageDiffrent);
		this.groupBox3.Controls.Add(this.ChkFirstPageDiffrent);
		this.groupBox3.Controls.Add(this.CmbFooterMiddleText);
		this.groupBox3.Controls.Add(this.CmbHeaderMiddleText);
		this.groupBox3.Controls.Add(this.BtnSetHeaderFooter);
		this.groupBox3.Controls.Add(this.PicBoxLOGOFile);
		this.groupBox3.Controls.Add(this.label18);
		this.groupBox3.Controls.Add(this.label17);
		this.groupBox3.Controls.Add(this.label16);
		this.groupBox3.Controls.Add(this.label15);
		this.groupBox3.Controls.Add(this.label14);
		this.groupBox3.Location = new System.Drawing.Point(4, 125);
		this.groupBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.groupBox3.Name = "groupBox3";
		this.groupBox3.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.groupBox3.Size = new System.Drawing.Size(850, 260);
		this.groupBox3.TabIndex = 5;
		this.groupBox3.TabStop = false;
		this.groupBox3.Text = "页眉页脚";
		this.ChkClearCurrent.AutoSize = true;
		this.ChkClearCurrent.Location = new System.Drawing.Point(632, 170);
		this.ChkClearCurrent.Name = "ChkClearCurrent";
		this.ChkClearCurrent.Size = new System.Drawing.Size(210, 24);
		this.ChkClearCurrent.TabIndex = 94;
		this.ChkClearCurrent.Text = "内容为空时清除既有页眉页脚";
		this.ChkClearCurrent.UseVisualStyleBackColor = true;
		this.CmbLogoIndex.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbLogoIndex.FormattingEnabled = true;
		this.CmbLogoIndex.Items.AddRange(new object[3] { "1", "2", "3" });
		this.CmbLogoIndex.Location = new System.Drawing.Point(543, 168);
		this.CmbLogoIndex.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbLogoIndex.Name = "CmbLogoIndex";
		this.CmbLogoIndex.Size = new System.Drawing.Size(46, 28);
		this.CmbLogoIndex.TabIndex = 93;
		this.CmbLogoIndex.SelectedIndexChanged += new System.EventHandler(CmbLogoIndex_SelectedIndexChanged);
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(490, 171);
		this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(49, 20);
		this.label1.TabIndex = 92;
		this.label1.Text = "Logo-";
		this.CmbHeaderFooterFontName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbHeaderFooterFontName.FormattingEnabled = true;
		this.CmbHeaderFooterFontName.Location = new System.Drawing.Point(599, 30);
		this.CmbHeaderFooterFontName.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderFooterFontName.Name = "CmbHeaderFooterFontName";
		this.CmbHeaderFooterFontName.Size = new System.Drawing.Size(109, 28);
		this.CmbHeaderFooterFontName.TabIndex = 91;
		this.CmbHeaderFooterFontName.SelectedIndexChanged += new System.EventHandler(CmbHeaderFooterFontName_SelectedIndexChanged);
		this.CmbHeaderFooterFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbHeaderFooterFont.FormattingEnabled = true;
		this.CmbHeaderFooterFont.Items.AddRange(new object[2] { "页眉", "页脚" });
		this.CmbHeaderFooterFont.Location = new System.Drawing.Point(526, 30);
		this.CmbHeaderFooterFont.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderFooterFont.Name = "CmbHeaderFooterFont";
		this.CmbHeaderFooterFont.Size = new System.Drawing.Size(65, 28);
		this.CmbHeaderFooterFont.TabIndex = 90;
		this.CmbHeaderFooterFont.SelectedIndexChanged += new System.EventHandler(CmbHeaderFooterFont_SelectedIndexChanged);
		this.TogItalic.BackColor = System.Drawing.Color.AliceBlue;
		this.TogItalic.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, 134);
		this.TogItalic.Location = new System.Drawing.Point(813, 30);
		this.TogItalic.Name = "TogItalic";
		this.TogItalic.Pressed = false;
		this.TogItalic.Size = new System.Drawing.Size(28, 28);
		this.TogItalic.TabIndex = 89;
		this.TogItalic.Text = "I";
		this.TogItalic.UseVisualStyleBackColor = false;
		this.TogItalic.Click += new System.EventHandler(TogBold_Click);
		this.TogBold.BackColor = System.Drawing.Color.AliceBlue;
		this.TogBold.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 134);
		this.TogBold.Location = new System.Drawing.Point(783, 30);
		this.TogBold.Name = "TogBold";
		this.TogBold.Pressed = false;
		this.TogBold.Size = new System.Drawing.Size(28, 28);
		this.TogBold.TabIndex = 88;
		this.TogBold.Text = "B";
		this.TogBold.UseVisualStyleBackColor = false;
		this.TogBold.Click += new System.EventHandler(TogBold_Click);
		this.RdoApplyToSection.AutoSize = true;
		this.RdoApplyToSection.Location = new System.Drawing.Point(245, 198);
		this.RdoApplyToSection.Name = "RdoApplyToSection";
		this.RdoApplyToSection.Size = new System.Drawing.Size(97, 24);
		this.RdoApplyToSection.TabIndex = 87;
		this.RdoApplyToSection.TabStop = true;
		this.RdoApplyToSection.Text = "应用于本节";
		this.RdoApplyToSection.UseVisualStyleBackColor = true;
		this.RdoApplyToSection.CheckedChanged += new System.EventHandler(RdoApplyToDocument_CheckedChanged);
		this.RdoApplyToSectionEnd.AutoSize = true;
		this.RdoApplyToSectionEnd.Location = new System.Drawing.Point(114, 198);
		this.RdoApplyToSectionEnd.Name = "RdoApplyToSectionEnd";
		this.RdoApplyToSectionEnd.Size = new System.Drawing.Size(125, 24);
		this.RdoApplyToSectionEnd.TabIndex = 86;
		this.RdoApplyToSectionEnd.TabStop = true;
		this.RdoApplyToSectionEnd.Text = "应用于本节之后";
		this.RdoApplyToSectionEnd.UseVisualStyleBackColor = true;
		this.RdoApplyToSectionEnd.CheckedChanged += new System.EventHandler(RdoApplyToDocument_CheckedChanged);
		this.RdoApplyToDocument.AutoSize = true;
		this.RdoApplyToDocument.Checked = true;
		this.RdoApplyToDocument.Location = new System.Drawing.Point(7, 198);
		this.RdoApplyToDocument.Name = "RdoApplyToDocument";
		this.RdoApplyToDocument.Size = new System.Drawing.Size(97, 24);
		this.RdoApplyToDocument.TabIndex = 85;
		this.RdoApplyToDocument.TabStop = true;
		this.RdoApplyToDocument.Text = "应用于全文";
		this.RdoApplyToDocument.UseVisualStyleBackColor = true;
		this.RdoApplyToDocument.CheckedChanged += new System.EventHandler(RdoApplyToDocument_CheckedChanged);
		this.ChkSameHeaderFooterHeight.AutoSize = true;
		this.ChkSameHeaderFooterHeight.Location = new System.Drawing.Point(8, 32);
		this.ChkSameHeaderFooterHeight.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkSameHeaderFooterHeight.Name = "ChkSameHeaderFooterHeight";
		this.ChkSameHeaderFooterHeight.Size = new System.Drawing.Size(84, 24);
		this.ChkSameHeaderFooterHeight.TabIndex = 84;
		this.ChkSameHeaderFooterHeight.Text = "统一高度";
		this.ChkSameHeaderFooterHeight.UseVisualStyleBackColor = true;
		this.ChkSameHeaderFooterHeight.CheckedChanged += new System.EventHandler(ChkSameHeaderFooterHeight_CheckedChanged);
		this.ChkRestartAtSection.AutoSize = true;
		this.ChkRestartAtSection.Location = new System.Drawing.Point(8, 226);
		this.ChkRestartAtSection.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkRestartAtSection.Name = "ChkRestartAtSection";
		this.ChkRestartAtSection.Size = new System.Drawing.Size(140, 24);
		this.ChkRestartAtSection.TabIndex = 83;
		this.ChkRestartAtSection.Text = "页码起始起始编号";
		this.ChkRestartAtSection.UseVisualStyleBackColor = true;
		this.ChkRestartAtSection.CheckedChanged += new System.EventHandler(ChkRestartAtSection_CheckedChanged);
		this.LabCurrentSectionNo.Location = new System.Drawing.Point(617, 195);
		this.LabCurrentSectionNo.Name = "LabCurrentSectionNo";
		this.LabCurrentSectionNo.Size = new System.Drawing.Size(225, 20);
		this.LabCurrentSectionNo.TabIndex = 82;
		this.LabCurrentSectionNo.Text = "LabInfo";
		this.LabCurrentSectionNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
		this.CmbFooterlineTypeSelect.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		this.CmbFooterlineTypeSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbFooterlineTypeSelect.FormattingEnabled = true;
		this.CmbFooterlineTypeSelect.Items.AddRange(new object[6] { "细实线", "双细实线", "细粗实线", "粗细实线", "粗实线", "无" });
		this.CmbFooterlineTypeSelect.Location = new System.Drawing.Point(691, 133);
		this.CmbFooterlineTypeSelect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbFooterlineTypeSelect.Name = "CmbFooterlineTypeSelect";
		this.CmbFooterlineTypeSelect.Size = new System.Drawing.Size(150, 27);
		this.CmbFooterlineTypeSelect.TabIndex = 20;
		this.CmbFooterlineTypeSelect.SelectedIndexChanged += new System.EventHandler(CmbHeaderlineTypeSelect_SelectedIndexChanged);
		this.CmbHeaderlineTypeSelect.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
		this.CmbHeaderlineTypeSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbHeaderlineTypeSelect.FormattingEnabled = true;
		this.CmbHeaderlineTypeSelect.Items.AddRange(new object[6] { "细实线", "双细实线", "细粗实线", "粗细实线", "粗实线", "无" });
		this.CmbHeaderlineTypeSelect.Location = new System.Drawing.Point(691, 98);
		this.CmbHeaderlineTypeSelect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderlineTypeSelect.Name = "CmbHeaderlineTypeSelect";
		this.CmbHeaderlineTypeSelect.Size = new System.Drawing.Size(150, 27);
		this.CmbHeaderlineTypeSelect.TabIndex = 15;
		this.CmbHeaderlineTypeSelect.SelectedIndexChanged += new System.EventHandler(CmbHeaderlineTypeSelect_SelectedIndexChanged);
		this.CmbLOGOSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbLOGOSize.FormattingEnabled = true;
		this.CmbLOGOSize.Items.AddRange(new object[8] { "1倍字高", "1.2倍字高", "1.5倍字高", "2倍字高", "2.5倍字高", "3倍字高", "0.75倍字高", "0.5倍字高" });
		this.CmbLOGOSize.Location = new System.Drawing.Point(489, 222);
		this.CmbLOGOSize.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbLOGOSize.Name = "CmbLOGOSize";
		this.CmbLOGOSize.Size = new System.Drawing.Size(100, 28);
		this.CmbLOGOSize.TabIndex = 28;
		this.CmbLOGOSize.SelectedIndexChanged += new System.EventHandler(CmbLOGOSize_SelectedIndexChanged);
		this.LabAddLogo.AutoSize = true;
		this.LabAddLogo.Font = new System.Drawing.Font("微软雅黑", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.LabAddLogo.Location = new System.Drawing.Point(379, 203);
		this.LabAddLogo.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.LabAddLogo.Name = "LabAddLogo";
		this.LabAddLogo.Size = new System.Drawing.Size(86, 17);
		this.LabAddLogo.TabIndex = 77;
		this.LabAddLogo.Text = "双击添加Logo";
		this.LabAddLogo.DoubleClick += new System.EventHandler(PicBoxLOGOFile_DoubleClick);
		this.label12.AutoSize = true;
		this.label12.Location = new System.Drawing.Point(490, 197);
		this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label12.Name = "label12";
		this.label12.Size = new System.Drawing.Size(99, 20);
		this.label12.TabIndex = 77;
		this.label12.Text = "Logo相对大小";
		this.CmbHeaderFooterFontSize.Location = new System.Drawing.Point(716, 30);
		this.CmbHeaderFooterFontSize.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderFooterFontSize.Name = "CmbHeaderFooterFontSize";
		this.CmbHeaderFooterFontSize.Size = new System.Drawing.Size(60, 28);
		this.CmbHeaderFooterFontSize.TabIndex = 23;
		this.CmbHeaderFooterFontSize.SelectedIndexChanged += new System.EventHandler(CmbHeaderFooterFontSize_SelectedIndexChanged);
		this.CmbHeaderFooterFontSize.Leave += new System.EventHandler(CmbHeaderFooterFontSize_Leave);
		this.label11.AutoSize = true;
		this.label11.Location = new System.Drawing.Point(488, 34);
		this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label11.Name = "label11";
		this.label11.Size = new System.Drawing.Size(37, 20);
		this.label11.TabIndex = 75;
		this.label11.Text = "字体";
		this.NumUpDownPageStart.Enabled = false;
		this.NumUpDownPageStart.Location = new System.Drawing.Point(154, 225);
		this.NumUpDownPageStart.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownPageStart.Maximum = new decimal(new int[4] { 999999, 0, 0, 0 });
		this.NumUpDownPageStart.Name = "NumUpDownPageStart";
		this.NumUpDownPageStart.Size = new System.Drawing.Size(85, 26);
		this.NumUpDownPageStart.TabIndex = 27;
		this.NumUpDownPageStart.Value = new decimal(new int[4] { 1, 0, 0, 0 });
		this.NumUpDownPageStart.ValueChanged += new System.EventHandler(NumUpDownPageStart_ValueChanged);
		this.NumUpDownFooterHeight.DecimalPlaces = 2;
		this.NumUpDownFooterHeight.Enabled = false;
		this.NumUpDownFooterHeight.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownFooterHeight.Label = "厘米";
		this.NumUpDownFooterHeight.Location = new System.Drawing.Point(340, 31);
		this.NumUpDownFooterHeight.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownFooterHeight.Name = "NumUpDownFooterHeight";
		this.NumUpDownFooterHeight.Size = new System.Drawing.Size(90, 26);
		this.NumUpDownFooterHeight.TabIndex = 22;
		this.NumUpDownHeaderHeight.DecimalPlaces = 2;
		this.NumUpDownHeaderHeight.Enabled = false;
		this.NumUpDownHeaderHeight.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownHeaderHeight.Label = "厘米";
		this.NumUpDownHeaderHeight.Location = new System.Drawing.Point(179, 31);
		this.NumUpDownHeaderHeight.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownHeaderHeight.Name = "NumUpDownHeaderHeight";
		this.NumUpDownHeaderHeight.Size = new System.Drawing.Size(90, 26);
		this.NumUpDownHeaderHeight.TabIndex = 21;
		this.label9.AutoSize = true;
		this.label9.Location = new System.Drawing.Point(273, 34);
		this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label9.Name = "label9";
		this.label9.Size = new System.Drawing.Size(65, 20);
		this.label9.TabIndex = 72;
		this.label9.Text = "页脚高度";
		this.label8.AutoSize = true;
		this.label8.Location = new System.Drawing.Point(111, 34);
		this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label8.Name = "label8";
		this.label8.Size = new System.Drawing.Size(65, 20);
		this.label8.TabIndex = 72;
		this.label8.Text = "页眉高度";
		this.CmbFooterRightText.FormattingEnabled = true;
		this.CmbFooterRightText.Items.AddRange(new object[9] { "#", "- # -", "第 # 页", "第 # 页 / 共 $ 页", "Page #", "Page # of $", "[LOGO1]", "[LOGO2]", "[LOGO3]" });
		this.CmbFooterRightText.Location = new System.Drawing.Point(505, 132);
		this.CmbFooterRightText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbFooterRightText.Name = "CmbFooterRightText";
		this.CmbFooterRightText.Size = new System.Drawing.Size(180, 28);
		this.CmbFooterRightText.TabIndex = 19;
		this.CmbFooterRightText.Validated += new System.EventHandler(CmbHeaderFooter_Validated);
		this.CmbHeaderRightText.FormattingEnabled = true;
		this.CmbHeaderRightText.Items.AddRange(new object[9] { "#", "- # -", "第 # 页", "第 # 页 / 共 $ 页", "Page #", "Page # of $", "[LOGO1]", "[LOGO2]", "[LOGO3]" });
		this.CmbHeaderRightText.Location = new System.Drawing.Point(505, 97);
		this.CmbHeaderRightText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderRightText.Name = "CmbHeaderRightText";
		this.CmbHeaderRightText.Size = new System.Drawing.Size(180, 28);
		this.CmbHeaderRightText.TabIndex = 14;
		this.CmbHeaderRightText.Validated += new System.EventHandler(CmbHeaderFooter_Validated);
		this.CmbFooterSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbFooterSelect.FormattingEnabled = true;
		this.CmbFooterSelect.Items.AddRange(new object[1] { "页脚" });
		this.CmbFooterSelect.Location = new System.Drawing.Point(8, 132);
		this.CmbFooterSelect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbFooterSelect.Name = "CmbFooterSelect";
		this.CmbFooterSelect.Size = new System.Drawing.Size(110, 28);
		this.CmbFooterSelect.TabIndex = 16;
		this.CmbFooterSelect.SelectedIndexChanged += new System.EventHandler(CmbFooterSelect_SelectedIndexChanged);
		this.CmbFooterLeftText.FormattingEnabled = true;
		this.CmbFooterLeftText.Items.AddRange(new object[9] { "#", "- # -", "第 # 页", "第 # 页 / 共 $ 页", "Page #", "Page # of $", "[LOGO1]", "[LOGO2]", "[LOGO3]" });
		this.CmbFooterLeftText.Location = new System.Drawing.Point(133, 132);
		this.CmbFooterLeftText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbFooterLeftText.Name = "CmbFooterLeftText";
		this.CmbFooterLeftText.Size = new System.Drawing.Size(180, 28);
		this.CmbFooterLeftText.TabIndex = 17;
		this.CmbFooterLeftText.Validated += new System.EventHandler(CmbHeaderFooter_Validated);
		this.CmbHeaderSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbHeaderSelect.FormattingEnabled = true;
		this.CmbHeaderSelect.Items.AddRange(new object[1] { "页眉" });
		this.CmbHeaderSelect.Location = new System.Drawing.Point(8, 97);
		this.CmbHeaderSelect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderSelect.Name = "CmbHeaderSelect";
		this.CmbHeaderSelect.Size = new System.Drawing.Size(110, 28);
		this.CmbHeaderSelect.TabIndex = 11;
		this.CmbHeaderSelect.SelectedIndexChanged += new System.EventHandler(CmbHeaderSelect_SelectedIndexChanged);
		this.CmbHeaderLeftText.FormattingEnabled = true;
		this.CmbHeaderLeftText.Items.AddRange(new object[9] { "#", "- # -", "第 # 页", "第 # 页 / 共 $ 页", "Page #", "Page # of $", "[LOGO1]", "[LOGO2]", "[LOGO3]" });
		this.CmbHeaderLeftText.Location = new System.Drawing.Point(133, 97);
		this.CmbHeaderLeftText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderLeftText.Name = "CmbHeaderLeftText";
		this.CmbHeaderLeftText.Size = new System.Drawing.Size(180, 28);
		this.CmbHeaderLeftText.TabIndex = 12;
		this.CmbHeaderLeftText.Validated += new System.EventHandler(CmbHeaderFooter_Validated);
		this.ChkOddEvenPageDiffrent.AutoSize = true;
		this.ChkOddEvenPageDiffrent.Location = new System.Drawing.Point(114, 171);
		this.ChkOddEvenPageDiffrent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkOddEvenPageDiffrent.Name = "ChkOddEvenPageDiffrent";
		this.ChkOddEvenPageDiffrent.Size = new System.Drawing.Size(98, 24);
		this.ChkOddEvenPageDiffrent.TabIndex = 25;
		this.ChkOddEvenPageDiffrent.Text = "奇偶页不同";
		this.ChkOddEvenPageDiffrent.UseVisualStyleBackColor = true;
		this.ChkOddEvenPageDiffrent.CheckedChanged += new System.EventHandler(ChkOddEvenPageDiffrent_CheckedChanged);
		this.ChkFirstPageDiffrent.AutoSize = true;
		this.ChkFirstPageDiffrent.Location = new System.Drawing.Point(8, 171);
		this.ChkFirstPageDiffrent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkFirstPageDiffrent.Name = "ChkFirstPageDiffrent";
		this.ChkFirstPageDiffrent.Size = new System.Drawing.Size(84, 24);
		this.ChkFirstPageDiffrent.TabIndex = 24;
		this.ChkFirstPageDiffrent.Text = "首页不同";
		this.ChkFirstPageDiffrent.UseVisualStyleBackColor = true;
		this.ChkFirstPageDiffrent.CheckedChanged += new System.EventHandler(ChkFirstPageDiffrent_CheckedChanged);
		this.CmbFooterMiddleText.FormattingEnabled = true;
		this.CmbFooterMiddleText.Items.AddRange(new object[9] { "#", "- # -", "第 # 页", "第 # 页 / 共 $ 页", "Page #", "Page # of $", "[LOGO1]", "[LOGO2]", "[LOGO3]" });
		this.CmbFooterMiddleText.Location = new System.Drawing.Point(319, 132);
		this.CmbFooterMiddleText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbFooterMiddleText.Name = "CmbFooterMiddleText";
		this.CmbFooterMiddleText.Size = new System.Drawing.Size(180, 28);
		this.CmbFooterMiddleText.TabIndex = 18;
		this.CmbFooterMiddleText.Validated += new System.EventHandler(CmbHeaderFooter_Validated);
		this.CmbHeaderMiddleText.FormattingEnabled = true;
		this.CmbHeaderMiddleText.Items.AddRange(new object[9] { "#", "- # -", "第 # 页", "第 # 页 / 共 $ 页", "Page #", "Page # of $", "[LOGO1]", "[LOGO2]", "[LOGO3]" });
		this.CmbHeaderMiddleText.Location = new System.Drawing.Point(319, 97);
		this.CmbHeaderMiddleText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbHeaderMiddleText.Name = "CmbHeaderMiddleText";
		this.CmbHeaderMiddleText.Size = new System.Drawing.Size(180, 28);
		this.CmbHeaderMiddleText.TabIndex = 13;
		this.CmbHeaderMiddleText.Validated += new System.EventHandler(CmbHeaderFooter_Validated);
		this.BtnSetHeaderFooter.Location = new System.Drawing.Point(722, 220);
		this.BtnSetHeaderFooter.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.BtnSetHeaderFooter.Name = "BtnSetHeaderFooter";
		this.BtnSetHeaderFooter.Size = new System.Drawing.Size(120, 30);
		this.BtnSetHeaderFooter.TabIndex = 29;
		this.BtnSetHeaderFooter.Text = "应用设置";
		this.BtnSetHeaderFooter.UseVisualStyleBackColor = true;
		this.BtnSetHeaderFooter.Click += new System.EventHandler(BtnSetHeaderFooter_Click);
		this.PicBoxLOGOFile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
		this.PicBoxLOGOFile.Location = new System.Drawing.Point(362, 171);
		this.PicBoxLOGOFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.PicBoxLOGOFile.Name = "PicBoxLOGOFile";
		this.PicBoxLOGOFile.Size = new System.Drawing.Size(120, 80);
		this.PicBoxLOGOFile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
		this.PicBoxLOGOFile.TabIndex = 73;
		this.PicBoxLOGOFile.TabStop = false;
		this.PicBoxLOGOFile.DoubleClick += new System.EventHandler(PicBoxLOGOFile_DoubleClick);
		this.label18.AutoSize = true;
		this.label18.Location = new System.Drawing.Point(699, 72);
		this.label18.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label18.Name = "label18";
		this.label18.Size = new System.Drawing.Size(135, 20);
		this.label18.TabIndex = 81;
		this.label18.Text = "分隔线（不分类型）";
		this.label17.AutoSize = true;
		this.label17.Location = new System.Drawing.Point(563, 72);
		this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label17.Name = "label17";
		this.label17.Size = new System.Drawing.Size(65, 20);
		this.label17.TabIndex = 81;
		this.label17.Text = "右侧内容";
		this.label16.AutoSize = true;
		this.label16.Location = new System.Drawing.Point(377, 72);
		this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label16.Name = "label16";
		this.label16.Size = new System.Drawing.Size(65, 20);
		this.label16.TabIndex = 81;
		this.label16.Text = "居中内容";
		this.label15.AutoSize = true;
		this.label15.Location = new System.Drawing.Point(191, 72);
		this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label15.Name = "label15";
		this.label15.Size = new System.Drawing.Size(65, 20);
		this.label15.TabIndex = 81;
		this.label15.Text = "左侧内容";
		this.label14.AutoSize = true;
		this.label14.Location = new System.Drawing.Point(45, 72);
		this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label14.Name = "label14";
		this.label14.Size = new System.Drawing.Size(37, 20);
		this.label14.TabIndex = 81;
		this.label14.Text = "类型";
		this.groupBox2.Controls.Add(this.label5);
		this.groupBox2.Controls.Add(this.Num_RightMargin);
		this.groupBox2.Controls.Add(this.label4);
		this.groupBox2.Controls.Add(this.Num_LeftMargin);
		this.groupBox2.Controls.Add(this.label3);
		this.groupBox2.Controls.Add(this.Num_BottomMargin);
		this.groupBox2.Controls.Add(this.label2);
		this.groupBox2.Controls.Add(this.CmbPageMarginType);
		this.groupBox2.Controls.Add(this.ChkSetBookbinding);
		this.groupBox2.Controls.Add(this.ChkSetPageMargin);
		this.groupBox2.Controls.Add(this.ChkApplySectionMargin);
		this.groupBox2.Controls.Add(this.BtnApplyPageMargin);
		this.groupBox2.Controls.Add(this.CmbBookbinding);
		this.groupBox2.Controls.Add(this.Num_TopMargin);
		this.groupBox2.Controls.Add(this.NumUpDownBookbinding);
		this.groupBox2.Location = new System.Drawing.Point(4, 5);
		this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.groupBox2.Size = new System.Drawing.Size(850, 110);
		this.groupBox2.TabIndex = 4;
		this.groupBox2.TabStop = false;
		this.groupBox2.Text = "页边距";
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(274, 67);
		this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(23, 20);
		this.label5.TabIndex = 88;
		this.label5.Text = "右";
		this.Num_RightMargin.DecimalPlaces = 2;
		this.Num_RightMargin.Enabled = false;
		this.Num_RightMargin.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.Num_RightMargin.Label = "厘米";
		this.Num_RightMargin.Location = new System.Drawing.Point(302, 64);
		this.Num_RightMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Num_RightMargin.Name = "Num_RightMargin";
		this.Num_RightMargin.Size = new System.Drawing.Size(96, 26);
		this.Num_RightMargin.TabIndex = 87;
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(137, 67);
		this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(23, 20);
		this.label4.TabIndex = 86;
		this.label4.Text = "左";
		this.Num_LeftMargin.DecimalPlaces = 2;
		this.Num_LeftMargin.Enabled = false;
		this.Num_LeftMargin.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.Num_LeftMargin.Label = "厘米";
		this.Num_LeftMargin.Location = new System.Drawing.Point(166, 64);
		this.Num_LeftMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Num_LeftMargin.Name = "Num_LeftMargin";
		this.Num_LeftMargin.Size = new System.Drawing.Size(96, 26);
		this.Num_LeftMargin.TabIndex = 85;
		this.Num_LeftMargin.ValueChanged += new System.EventHandler(Num_LeftMargin_ValueChanged);
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(274, 31);
		this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(23, 20);
		this.label3.TabIndex = 84;
		this.label3.Text = "下";
		this.Num_BottomMargin.DecimalPlaces = 2;
		this.Num_BottomMargin.Enabled = false;
		this.Num_BottomMargin.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.Num_BottomMargin.Label = "厘米";
		this.Num_BottomMargin.Location = new System.Drawing.Point(302, 28);
		this.Num_BottomMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Num_BottomMargin.Name = "Num_BottomMargin";
		this.Num_BottomMargin.Size = new System.Drawing.Size(96, 26);
		this.Num_BottomMargin.TabIndex = 83;
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(137, 31);
		this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(23, 20);
		this.label2.TabIndex = 82;
		this.label2.Text = "上";
		this.CmbPageMarginType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbPageMarginType.Enabled = false;
		this.CmbPageMarginType.FormattingEnabled = true;
		this.CmbPageMarginType.Items.AddRange(new object[4] { "自由设置", "四边相等", "左右相等", "上下相等" });
		this.CmbPageMarginType.Location = new System.Drawing.Point(19, 63);
		this.CmbPageMarginType.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbPageMarginType.Name = "CmbPageMarginType";
		this.CmbPageMarginType.Size = new System.Drawing.Size(99, 28);
		this.CmbPageMarginType.TabIndex = 13;
		this.CmbPageMarginType.SelectedIndexChanged += new System.EventHandler(CmbPageMarginType_SelectedIndexChanged);
		this.ChkSetBookbinding.AutoSize = true;
		this.ChkSetBookbinding.Location = new System.Drawing.Point(489, 29);
		this.ChkSetBookbinding.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkSetBookbinding.Name = "ChkSetBookbinding";
		this.ChkSetBookbinding.Size = new System.Drawing.Size(98, 24);
		this.ChkSetBookbinding.TabIndex = 12;
		this.ChkSetBookbinding.Text = "设置装订线";
		this.ChkSetBookbinding.UseVisualStyleBackColor = true;
		this.ChkSetBookbinding.CheckedChanged += new System.EventHandler(ChkSetBookbinding_CheckedChanged);
		this.ChkSetPageMargin.AutoSize = true;
		this.ChkSetPageMargin.Location = new System.Drawing.Point(19, 29);
		this.ChkSetPageMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkSetPageMargin.Name = "ChkSetPageMargin";
		this.ChkSetPageMargin.Size = new System.Drawing.Size(98, 24);
		this.ChkSetPageMargin.TabIndex = 11;
		this.ChkSetPageMargin.Text = "设置页边距";
		this.ChkSetPageMargin.UseVisualStyleBackColor = true;
		this.ChkSetPageMargin.CheckedChanged += new System.EventHandler(ChkSetPageMargin_CheckedChanged);
		this.ChkApplySectionMargin.AutoSize = true;
		this.ChkApplySectionMargin.Checked = true;
		this.ChkApplySectionMargin.CheckState = System.Windows.Forms.CheckState.Checked;
		this.ChkApplySectionMargin.Location = new System.Drawing.Point(621, 74);
		this.ChkApplySectionMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.ChkApplySectionMargin.Name = "ChkApplySectionMargin";
		this.ChkApplySectionMargin.Size = new System.Drawing.Size(98, 24);
		this.ChkApplySectionMargin.TabIndex = 6;
		this.ChkApplySectionMargin.Text = "应用于本节";
		this.ChkApplySectionMargin.UseVisualStyleBackColor = true;
		this.BtnApplyPageMargin.Location = new System.Drawing.Point(722, 70);
		this.BtnApplyPageMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.BtnApplyPageMargin.Name = "BtnApplyPageMargin";
		this.BtnApplyPageMargin.Size = new System.Drawing.Size(120, 30);
		this.BtnApplyPageMargin.TabIndex = 10;
		this.BtnApplyPageMargin.Text = "应用设置";
		this.BtnApplyPageMargin.UseVisualStyleBackColor = true;
		this.BtnApplyPageMargin.Click += new System.EventHandler(BtnApplyPageMargin_Click);
		this.CmbBookbinding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.CmbBookbinding.Enabled = false;
		this.CmbBookbinding.FormattingEnabled = true;
		this.CmbBookbinding.Items.AddRange(new object[3] { "左侧装订", "顶部装订", "左右对称装订" });
		this.CmbBookbinding.Location = new System.Drawing.Point(599, 27);
		this.CmbBookbinding.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.CmbBookbinding.Name = "CmbBookbinding";
		this.CmbBookbinding.Size = new System.Drawing.Size(115, 28);
		this.CmbBookbinding.TabIndex = 4;
		this.Num_TopMargin.DecimalPlaces = 2;
		this.Num_TopMargin.Enabled = false;
		this.Num_TopMargin.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.Num_TopMargin.Label = "厘米";
		this.Num_TopMargin.Location = new System.Drawing.Point(166, 28);
		this.Num_TopMargin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Num_TopMargin.Name = "Num_TopMargin";
		this.Num_TopMargin.Size = new System.Drawing.Size(96, 26);
		this.Num_TopMargin.TabIndex = 0;
		this.Num_TopMargin.ValueChanged += new System.EventHandler(Num_TopMargin_ValueChanged);
		this.NumUpDownBookbinding.DecimalPlaces = 2;
		this.NumUpDownBookbinding.Enabled = false;
		this.NumUpDownBookbinding.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.NumUpDownBookbinding.Label = "厘米";
		this.NumUpDownBookbinding.Location = new System.Drawing.Point(722, 29);
		this.NumUpDownBookbinding.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.NumUpDownBookbinding.Name = "NumUpDownBookbinding";
		this.NumUpDownBookbinding.Size = new System.Drawing.Size(120, 26);
		this.NumUpDownBookbinding.TabIndex = 5;
		this.DlgOpenFiles.FileName = "openFileDialog1";
		base.AutoScaleDimensions = new System.Drawing.SizeF(8f, 20f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.groupBox3);
		base.Controls.Add(this.groupBox2);
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		base.Name = "PageSetUI";
		base.Size = new System.Drawing.Size(860, 390);
		this.groupBox3.ResumeLayout(false);
		this.groupBox3.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownPageStart).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownFooterHeight).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownHeaderHeight).EndInit();
		((System.ComponentModel.ISupportInitialize)this.PicBoxLOGOFile).EndInit();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.Num_RightMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_LeftMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_BottomMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.Num_TopMargin).EndInit();
		((System.ComponentModel.ISupportInitialize)this.NumUpDownBookbinding).EndInit();
		base.ResumeLayout(false);
	}
}
}