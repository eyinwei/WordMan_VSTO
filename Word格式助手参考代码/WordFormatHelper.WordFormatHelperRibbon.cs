// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.WordFormatHelperRibbon
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using WordFormatHelper;
using WordFormatHelper.Properties;
using WordFormatHelper.Tools;

public class WordFormatHelperRibbon : RibbonBase
{
	private Form UIForm;

	private Form ExportToPDFUI;

	private Form OCRUI;

	private WordFormatHelper.Settings SetPane;

	private CustomTaskPane SettingsUI;

	private readonly Dictionary<int, CustomTaskPane> hwndPaneDic = new Dictionary<int, CustomTaskPane>();

	private List<Range> multiSelection;

	private const int ENHANCEDMETAFILE = 14;

	private IContainer components;

	internal RibbonTab TabAddIns;

	private RibbonTab Tab_WordFormatAssistant;

	internal RibbonGroup Gp_PageSet;

	internal RibbonGallery Ga_FastMargin;

	internal RibbonGallery Ga_FastHeaderFooter;

	internal RibbonGallery Ga_HeaderLine;

	private RibbonButton Btn_SetHeaderSingleLine;

	private RibbonButton Btn_SetHeaderThickLine;

	private RibbonButton Btn_SetHeaderDoubleLine;

	private RibbonButton Btn_SetHeaderThinThickLine;

	private RibbonButton Btn_SetHeaderThickThinLine;

	private RibbonButton Btn_DeleteHeaderLine;

	internal RibbonGallery Ga_FooterLine;

	private RibbonButton Btn_SetFooterSingleLine;

	private RibbonButton Btn_SetFooterThickLine;

	private RibbonButton Btn_SetFooterDoubleLine;

	private RibbonButton Btn_SetFooterThinThickLine;

	private RibbonButton Btn_SetFooterThickThinLine;

	private RibbonButton Btn_DeleteFooterLine;

	internal RibbonGroup Gp_TextFormatSet;

	internal RibbonButton Btn_PunctuationEng2Chn;

	internal RibbonButton Btn_PunctuationChn2Eng;

	internal RibbonButton Btn_DeleteSpace;

	internal RibbonGallery Ga_InsertDate;

	internal RibbonGallery Ga_Bracketed;

	internal RibbonButton Btn_ToChinese;

	internal RibbonSplitButton Btn_ListPunctuation;

	internal RibbonButton Btn_ListSemicolonPeriod;

	internal RibbonButton Btn_ListCommaPeriod;

	internal RibbonButton Btn_ListPeriod;

	internal RibbonButton Btn_ListNoPunctuation;

	internal RibbonSplitButton Btn_ParagrahIndent;

	internal RibbonButton Btn_ParagrahIndent2Char;

	internal RibbonButton Btn_ParagrahNoIndent;

	internal RibbonGroup Gp_List;

	internal RibbonButton Btn_FastListFormat;

	internal RibbonGroup Gp_LevelList;

	internal RibbonGallery Ga_FastLevelList;

	private RibbonButton Btn_Set2LevelList;

	private RibbonButton Btn_Set3LevelList;

	private RibbonButton Btn_Set4LevelList;

	private RibbonButton Btn_Set5LevelList;

	internal RibbonGroup Gp_TOC;

	internal RibbonButton Btn_TOC;

	internal RibbonGroup Gp_TablePictureSet;

	internal RibbonToggleButton Tog_ApplyToAll;

	internal RibbonSeparator separator2;

	internal RibbonButton Btn_TableLeft;

	internal RibbonButton Btn_TableCenter;

	internal RibbonButton Btn_TableRight;

	internal RibbonButton Btn_FirstRowBold;

	internal RibbonButton Btn_FirstColumnBold;

	internal RibbonButton Btn_TableThickOutside;

	internal RibbonButton Btn_TableSingleSpace;

	internal RibbonButton Btn_RemoveUselessLine;

	internal RibbonButton Btn_RemoveLeftIndent;

	internal RibbonGallery GA_NewTables;

	private RibbonButton Btn_TableStyle;

	internal RibbonSeparator separator1;

	internal RibbonButton Btn_PictureLeft;

	internal RibbonButton Btn_PictureCenter;

	internal RibbonButton Btn_PictureRight;

	internal RibbonButton Btn_SetPictureWidth;

	internal RibbonButton Btn_SetPictureHeight;

	internal RibbonButton Btn_PictureSingleSpace;

	internal RibbonEditBox Ebox_PictureWidth;

	internal RibbonEditBox Ebox_PictureHeight;

	internal RibbonGroup Gp_DocumentSet;

	internal RibbonButton Btn_FastSetStyle;

	internal RibbonGroup Gp_QRCoder;

	internal RibbonButton Btn_FastQRCoder;

	internal RibbonGroup Gp_Help;

	internal RibbonGallery Ga_SetUnits;

	private RibbonButton Btn_UnitInch;

	private RibbonButton Btn_UnitCM;

	private RibbonButton Btn_UnitMM;

	private RibbonButton Btn_UnitPt;

	private RibbonButton Btn_UnitPicas;

	internal RibbonButton Btn_DefaultValue;

	internal RibbonGallery GA_AboutAndHelp;

	private RibbonButton Btn_HelpOnline;

	private RibbonButton Btn_AboutUs;

	internal RibbonButton Btn_TableFullWidth;

	internal RibbonGallery Ga_Utilities;

	private RibbonButton Btn_ExportPDF;

	internal RibbonButton Btn_ListStartFromOne;

	internal RibbonButton Btn_RowHeightFitText;

	internal RibbonButton Btn_RepeatTitle;

	internal RibbonButton Btn_TransToList;

	internal RibbonButton Btn_DeleteBlankLine;

	internal RibbonSplitButton Btn_FormatPainter;

	internal RibbonButton Btn_FormatPainterUI;

	private RibbonButton Btn_OCR;

	private RibbonButton Btn_ExportPng;

	internal RibbonSplitButton Btn_SuperscriptAndSubscript;

	internal RibbonButton Btn_SquareSuperscript;

	internal RibbonButton Btn_CubeSuperscript;

	internal RibbonButton Btn_NumberSuperscript;

	internal RibbonButton Btn_NumberSubscript;

	internal RibbonButton Btn_CustomScript;

	private void WordFormatHelperRibbon_Load(object sender, RibbonUIEventArgs e)
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "DocumentChange").AddEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_DocumentChangeEventHandler(FunctionEnabled));
		Ga_SetUnits.Buttons[(int)Globals.ThisAddIn.Application.Options.MeasurementUnit].OfficeImageId = "AcceptInvitation";
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowSelectionChange").AddEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowSelectionChangeEventHandler(MultiSelectMode));
		AddTips();
		Btn_SuperscriptAndSubscript.Label = Btn_SquareSuperscript.Label;
		Btn_SuperscriptAndSubscript.Image = Btn_SquareSuperscript.Image;
		Btn_SuperscriptAndSubscript.ScreenTip = Btn_SquareSuperscript.ScreenTip;
		Btn_SuperscriptAndSubscript.SuperTip = Btn_SquareSuperscript.SuperTip;
	}

	private void AddTips()
	{
		foreach (RibbonGroup group in Tab_WordFormatAssistant.Groups)
		{
			foreach (RibbonControl item in group.Items)
			{
				string text = Resources.ResourceManager.GetString(item.Name);
				if (text != null)
				{
					item.GetType().GetProperty("SuperTip")?.SetValue(item, text);
				}
				if (item is RibbonGallery ribbonGallery)
				{
					foreach (RibbonButton button in ribbonGallery.Buttons)
					{
						button.GetType().GetProperty("SuperTip")?.SetValue(button, Resources.ResourceManager.GetString(button.Name));
					}
				}
				if (!(item is RibbonSplitButton ribbonSplitButton))
				{
					continue;
				}
				foreach (RibbonControl item2 in ribbonSplitButton.Items)
				{
					item2.GetType().GetProperty("SuperTip")?.SetValue(item2, Resources.ResourceManager.GetString(item2.Name));
				}
			}
		}
	}

	private void ShowFormatHelperUI(object sender, RibbonControlEventArgs e)
	{
		UserControl userControl = null;
		string text = "";
		string text2 = "";
		if (sender is RibbonGroup ribbonGroup)
		{
			text2 = ribbonGroup.Name;
		}
		else if (sender is RibbonButton ribbonButton)
		{
			text2 = ribbonButton.Name;
		}
		switch (text2)
		{
		case "Gp_LevelList":
			text = "多级列表";
			userControl = new LevelListSetUI();
			break;
		case "Gp_TablePictureSet":
			text = "图表格式";
			userControl = new TablePictureSetUI();
			break;
		case "Gp_List":
			text = "列表格式";
			userControl = new ListSetUI();
			break;
		case "Gp_QRCoder":
			text = "二维码条码设置";
			userControl = new QRCoderSetUI();
			break;
		case "Gp_TextFormatSet":
			text = "文本格式";
			userControl = new TextFormatUI();
			break;
		case "Gp_PageSet":
			text = "页面/页眉页脚格式";
			userControl = new PageSetUI();
			break;
		case "Btn_FastSetStyle":
			text = "文档样式设置";
			userControl = new StyleSetGuider();
			break;
		case "Btn_TableStyle":
			text = "表格样式设置";
			userControl = new TableStyleSettings();
			break;
		case "Btn_TOC":
			text = "目录设置";
			userControl = new TOCSet();
			break;
		case "Btn_RowHeightFitText":
			text = "表格行高适配内容";
			userControl = new TableRowHeightUI();
			break;
		case "Btn_FormatPainterUI":
			text = "固定格式刷";
			userControl = new FormatPainterUI();
			break;
		case "Btn_CustomScript":
			text = "上下标格式";
			userControl = new ScriptFormatUI();
			break;
		}
		if (UIForm == null || UIForm.IsDisposed)
		{
			UIForm = new Form
			{
				FormBorderStyle = FormBorderStyle.FixedSingle,
				MaximizeBox = false,
				MinimizeBox = false,
				Icon = Resources.WAIcon,
				Text = text,
				AutoScaleMode = AutoScaleMode.Dpi
			};
			UIForm.Controls.Clear();
			UIForm.Controls.Add(userControl);
			UIForm.ClientSize = userControl.Size;
			UIForm.StartPosition = FormStartPosition.CenterScreen;
			UIForm.ShowInTaskbar = false;
			UIForm.Disposed += WhenUIDisposed;
			UIForm.HelpRequested += UserHelp;
			UIForm.Show(Globals.ThisAddIn.Application.ActiveWindow as IWin32Window);
		}
		else
		{
			if (text != UIForm.Text)
			{
				UIForm.Controls.Clear();
				UIForm.Controls.Add(userControl);
				UIForm.ClientSize = userControl.Size;
				UIForm.Text = text;
			}
			UIForm.Activate();
		}
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowActivate").AddEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowActivateEventHandler(ShowDialogBetweenWindow));
	}

	private void WhenUIDisposed(object sender, EventArgs e)
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowActivate").RemoveEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowActivateEventHandler(ShowDialogBetweenWindow));
	}

	private void ShowDialogBetweenWindow(Document Doc, Window Wn)
	{
		if (!UIForm.IsDisposed)
		{
			if (UIForm.IsHandleCreated)
			{
				ThisAddIn.SetWindowLongPtrImp(UIForm.Handle, -8, Wn.Hwnd);
				UIForm.Activate();
			}
			else
			{
				UIForm.Visible = false;
				UIForm.Show(Wn as IWin32Window);
			}
		}
	}

	private void Btn_DefaultValue_Click(object sender, RibbonControlEventArgs e)
	{
		int hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
		ShowSettingPane(hwnd);
		SettingsUI.Visible = true;
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "WindowActivate").AddEventHandler(Globals.ThisAddIn.Application, new ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate));
	}

	private void Application_WindowActivate(Document Doc, Window Wn)
	{
		if (hwndPaneDic.ContainsKey(Wn.Hwnd) && SettingsUI.Visible)
		{
			(hwndPaneDic[Wn.Hwnd].Control as WordFormatHelper.Settings).ReadSettings();
		}
	}

	private void ShowSettingPane(int wHwnd)
	{
		if (hwndPaneDic.ContainsKey(wHwnd))
		{
			SettingsUI = hwndPaneDic[wHwnd];
			return;
		}
		SetPane = new WordFormatHelper.Settings();
		SettingsUI = Globals.ThisAddIn.CustomTaskPanes.Add(SetPane, "格式助手设置");
		SettingsUI.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
		SettingsUI.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
		CustomTaskPane settingsUI = SettingsUI;
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		object fVertical = Type.Missing;
		settingsUI.Width = (int)application.PixelsToPoints(330f, ref fVertical) + 20;
		hwndPaneDic.Add(wHwnd, SettingsUI);
	}

	internal void FunctionEnabled()
	{
		if (Globals.ThisAddIn.Application.Documents.Count == 0)
		{
			foreach (RibbonGroup group in Tab_WordFormatAssistant.Groups)
			{
				if (group.DialogLauncher != null)
				{
					group.DialogLauncher.Enabled = false;
				}
				foreach (RibbonControl item in group.Items)
				{
					(item as RibbonControl).Enabled = false;
				}
			}
			return;
		}
		foreach (RibbonGroup group2 in Tab_WordFormatAssistant.Groups)
		{
			if (group2.DialogLauncher != null)
			{
				group2.DialogLauncher.Enabled = true;
			}
			foreach (RibbonControl item2 in group2.Items)
			{
				(item2 as RibbonControl).Enabled = true;
			}
		}
	}

	private void Btn_FastQRCoder_Click(object sender, RibbonControlEventArgs e)
	{
		string text = "";
		if (Globals.ThisAddIn.Application.Selection.Type == WdSelectionType.wdSelectionIP)
		{
			InputForm inputForm = new InputForm("请输入需要生成二维码的文字内容");
			inputForm.ShowDialog();
			if (inputForm.OK)
			{
				text = inputForm.InputText;
			}
			inputForm.Dispose();
		}
		else if (Globals.ThisAddIn.Application.Selection.Hyperlinks.Count > 0)
		{
			Hyperlinks hyperlinks = Globals.ThisAddIn.Application.Selection.Hyperlinks;
			object Index = 1;
			text = hyperlinks[ref Index].Address;
		}
		else
		{
			text = Globals.ThisAddIn.Application.Selection.Range.Text;
		}
		text = text.Trim(" \r\n".ToCharArray());
		if (!(text == ""))
		{
			Globals.ThisAddIn.CreateQRCodeImage(text);
		}
	}

	private void Ga_SetUnits_ButtonClick(object sender, RibbonControlEventArgs e)
	{
		RibbonButton ribbonButton = sender as RibbonButton;
		foreach (RibbonButton button in Ga_SetUnits.Buttons)
		{
			button.OfficeImageId = "";
		}
		ribbonButton.OfficeImageId = "AcceptInvitation";
		Globals.ThisAddIn.Application.Options.MeasurementUnit = (WdMeasurementUnits)Ga_SetUnits.Buttons.IndexOf(ribbonButton);
	}

	private void MultiSelectMode(Selection Sel)
	{
		if ((multiSelection == null) & (Sel.Type == WdSelectionType.wdSelectionNormal))
		{
			multiSelection = new List<Range>(1) { Sel.Range };
		}
		else if ((Keyboard.IsKeyDown(Key.LeftCtrl) | Keyboard.IsKeyDown(Key.RightCtrl)) & (Sel.Type == WdSelectionType.wdSelectionNormal))
		{
			multiSelection.Add(Sel.Range);
		}
		else if (Sel.Type == WdSelectionType.wdSelectionNormal)
		{
			multiSelection = new List<Range>(1) { Sel.Range };
		}
		else
		{
			multiSelection = null;
		}
	}

	private void Tog_ApplyToAll_Click(object sender, RibbonControlEventArgs e)
	{
		if (Tog_ApplyToAll.Checked)
		{
			Tog_ApplyToAll.Image = Resources.Mode_Document_On;
		}
		else
		{
			Tog_ApplyToAll.Image = Resources.Mode_Document_Off;
		}
	}

	private void GA_AboutAndHelp_ButtonClick(object sender, RibbonControlEventArgs e)
	{
		if ((sender as RibbonButton).Name == "Btn_HelpOnline")
		{
			Process.Start("Https://Enocheasty.github.io/WordAssistantHelp/");
			return;
		}
		AboutMe aboutMe = new AboutMe();
		aboutMe.Icon = Resources.WAIcon;
		aboutMe.StartPosition = FormStartPosition.CenterScreen;
		aboutMe.Show();
	}

	private void UserHelp(object sender, HelpEventArgs hlpevent)
	{
		string text;
		switch ((sender as Form).Text)
		{
		case "页面/页眉页脚格式":
			text = "WAHelpPageSetup.html";
			break;
		case "文本格式":
		case "固定格式刷":
			text = "WAHelpTextFormat.html";
			break;
		case "多级列表":
		case "列表格式":
			text = "WAHelpListLevelList.html";
			break;
		case "目录设置":
			text = "WAHelpContents.html";
			break;
		case "图表格式":
		case "表格样式设置":
		case "表格行高适配内容":
			text = "WAHelpTablePicture.html";
			break;
		case "文档样式设置":
			text = "WAHelpStyles.html";
			break;
		case "二维码条码设置":
			text = "WAHelpQRCoder.html";
			break;
		default:
			text = "";
			break;
		}
		string text2 = text;
		string parameter = (sender as Form).Text switch
		{
			"页面/页眉页脚格式" => "PageMargin&HeaderFooterUI", 
			"固定格式刷" => "FixFormatPainter", 
			"列表格式" => "ListFormat", 
			"多级列表" => "LevelListFormat", 
			"表格样式设置" => "NewTableStyle", 
			"表格行高适配内容" => "", 
			_ => "", 
		};
		Help.ShowHelp(null, "https://enocheasty.github.io/WordAssistantHelp/" + text2, HelpNavigator.Topic, parameter);
	}

	private void Ga_Utilities_ButtonClick(object sender, RibbonControlEventArgs e)
	{
		switch ((sender as RibbonButton).Name)
		{
		case "Btn_ExportPDF":
			if (ExportToPDFUI == null)
			{
				ExportToPDFUI = new ExportToPDF();
			}
			if (ExportToPDFUI.IsDisposed)
			{
				ExportToPDFUI = new ExportToPDF();
			}
			ExportToPDFUI?.Show();
			ExportToPDFUI?.Activate();
			break;
		case "Btn_OCR":
			if (OCRUI == null)
			{
				OCRUI = new OCRTools();
			}
			if (OCRUI.IsDisposed)
			{
				OCRUI = new OCRTools();
			}
			OCRUI?.Show(Globals.ThisAddIn.Application.ActiveWindow as IWin32Window);
			OCRUI?.Activate();
			break;
		case "Btn_ExportPng":
			ExportSelectionToPng();
			break;
		}
	}

	[DllImport("User32")]
	private static extern bool OpenClipboard(IntPtr hWndNewOwner);

	[DllImport("User32")]
	private static extern bool CloseClipboard();

	[DllImport("User32")]
	private static extern IntPtr GetClipboardData(int uFormat);

	[DllImport("Gdi32", CharSet = CharSet.Unicode)]
	private static extern IntPtr CopyEnhMetaFile(IntPtr emfHandle, IntPtr fileName);

	private void ExportSelectionToPng()
	{
		if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionNormal)
		{
			return;
		}
		Globals.ThisAddIn.Application.Selection.CopyAsPicture();
		if (!OpenClipboard(IntPtr.Zero))
		{
			return;
		}
		IntPtr clipboardData = GetClipboardData(14);
		if (clipboardData != IntPtr.Zero)
		{
			using Metafile metafile = new Metafile(CopyEnhMetaFile(clipboardData, IntPtr.Zero), deleteEmf: true);
			if (metafile != null)
			{
				using Bitmap bitmap = new Bitmap(metafile.Width, metafile.Height);
				using Graphics graphics = Graphics.FromImage(bitmap);
				graphics.SmoothingMode = SmoothingMode.HighQuality;
				graphics.DrawImage(metafile, 0, 0, bitmap.Width, bitmap.Height);
				SaveFileDialog saveFileDialog = new SaveFileDialog
				{
					FileName = "导出图片",
					AddExtension = true,
					Filter = "图片文件|*.png",
					DefaultExt = ".png"
				};
				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					bitmap.Save(saveFileDialog.FileName);
					MessageBox.Show("图片已导出至:" + saveFileDialog.FileName, "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				}
			}
		}
		CloseClipboard();
	}

	private void SetMarginBtn_Click(object sender, RibbonControlEventArgs e)
	{
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		float[] array = new float[4];
		switch ((sender as RibbonGallery).SelectedItemIndex)
		{
		case 0:
			array[0] = defaultValue.PageTopMargin;
			array[1] = defaultValue.PageBottomMargin;
			array[2] = defaultValue.PageLeftMargin;
			array[3] = defaultValue.PageRightMargin;
			break;
		case 1:
			array[0] = (array[1] = (array[2] = (array[3] = 2f)));
			break;
		case 2:
			array[0] = (array[1] = (array[2] = (array[3] = 2.5f)));
			break;
		case 3:
			array[0] = (array[1] = (array[2] = (array[3] = 3f)));
			break;
		case 4:
			array[0] = (array[1] = (array[2] = (array[3] = 3.5f)));
			break;
		default:
			array[0] = (array[1] = (array[2] = (array[3] = 0f)));
			break;
		}
		try
		{
			activeDocument.PageSetup.TopMargin = application.CentimetersToPoints(array[0]);
			activeDocument.PageSetup.BottomMargin = application.CentimetersToPoints(array[1]);
			activeDocument.PageSetup.LeftMargin = application.CentimetersToPoints(array[2]);
			activeDocument.PageSetup.RightMargin = application.CentimetersToPoints(array[3]);
		}
		catch
		{
			if (MessageBox.Show("设置页边距的过程出现问题！可能由于文档不同分节页面设置不统一，是否按节设置页边距？", "Word格式助手", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) != DialogResult.Yes)
			{
				return;
			}
			foreach (Section section in activeDocument.Sections)
			{
				try
				{
					section.PageSetup.TopMargin = application.CentimetersToPoints(array[0]);
					section.PageSetup.BottomMargin = application.CentimetersToPoints(array[1]);
					section.PageSetup.LeftMargin = application.CentimetersToPoints(array[2]);
					section.PageSetup.RightMargin = application.CentimetersToPoints(array[3]);
				}
				catch
				{
					MessageBox.Show("第" + section.Index + "设置页边距出现问题，请检查是否因分栏引起无法设置边距！", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
			}
		}
	}

	private void SetHeaderFooterLine_Click(object sender, RibbonControlEventArgs e)
	{
		int lineType;
		bool flag;
		switch ((sender as RibbonButton).Name)
		{
		case "Btn_SetFooterSingleLine":
		case "Btn_SetHeaderSingleLine":
			lineType = 0;
			flag = false;
			break;
		case "Btn_SetFooterThickLine":
		case "Btn_SetHeaderThickLine":
			lineType = 4;
			flag = false;
			break;
		case "Btn_SetFooterDoubleLine":
		case "Btn_SetHeaderDoubleLine":
			lineType = 1;
			flag = false;
			break;
		case "Btn_SetFooterThinThickLine":
		case "Btn_SetHeaderThinThickLine":
			lineType = 2;
			flag = false;
			break;
		case "Btn_SetFooterThickThinLine":
		case "Btn_SetHeaderThickThinLine":
			lineType = 3;
			flag = false;
			break;
		case "Btn_DeleteFooterLine":
		case "Btn_DeleteHeaderLine":
			lineType = 0;
			flag = true;
			break;
		default:
			lineType = 0;
			flag = false;
			break;
		}
		bool flag2 = Regex.IsMatch((sender as RibbonButton).Name, "Header");
		Section first = Globals.ThisAddIn.Application.Selection.Sections.First;
		Section section = ((first.Index < Globals.ThisAddIn.Application.ActiveDocument.Sections.Count) ? Globals.ThisAddIn.Application.ActiveDocument.Sections[first.Index + 1] : null);
		HeaderFooter headerFooter = (flag2 ? first.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary] : first.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary]);
		headerFooter.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
		headerFooter.LinkToPrevious = false;
		if (section != null)
		{
			if (flag2)
			{
				section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
			}
			else
			{
				section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
			}
		}
		if (!flag)
		{
			Globals.ThisAddIn.InsertSplitLine(headerFooter, flag2, lineType);
		}
		if (first.PageSetup.OddAndEvenPagesHeaderFooter != -1)
		{
			return;
		}
		headerFooter = (flag2 ? first.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages] : first.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
		headerFooter.Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
		headerFooter.LinkToPrevious = false;
		if (section != null)
		{
			if (flag2)
			{
				section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
			}
			else
			{
				section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
			}
		}
		if (!flag)
		{
			Globals.ThisAddIn.InsertSplitLine(headerFooter, flag2, lineType);
		}
	}

	private void Ga_FastHeaderFooter_Click(object sender, RibbonControlEventArgs e)
	{
		bool flag = false;
		Section first = Globals.ThisAddIn.Application.Selection.Sections.First;
		HeaderFooter headerFooter = first.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
		headerFooter.LinkToPrevious = false;
		Section section = ((first.Index < Globals.ThisAddIn.Application.ActiveDocument.Sections.Count) ? Globals.ThisAddIn.Application.ActiveDocument.Sections[first.Index + 1] : null);
		if (section != null)
		{
			section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
		}
		bool flag2 = first.PageSetup.OddAndEvenPagesHeaderFooter == -1;
		string[] array = new string[3] { "", "", "" };
		switch (Ga_FastHeaderFooter.SelectedItemIndex)
		{
		case 0:
		case 4:
			flag = Ga_FastHeaderFooter.SelectedItemIndex == 0;
			array[1] = "第#页 / 共$页";
			break;
		case 1:
		case 5:
			flag = Ga_FastHeaderFooter.SelectedItemIndex == 1;
			array[1] = "第#页";
			break;
		case 2:
		case 6:
			flag = Ga_FastHeaderFooter.SelectedItemIndex == 2;
			array[1] = "Page #";
			break;
		case 3:
		case 7:
			flag = Ga_FastHeaderFooter.SelectedItemIndex == 3;
			array[1] = "Page # of $";
			break;
		}
		Globals.ThisAddIn.ResetHeaderFooterStyle();
		float num = first.PageSetup.PageWidth - first.PageSetup.LeftMargin - first.PageSetup.RightMargin;
		num = ((first.PageSetup.GutterPos == WdGutterStyle.wdGutterPosTop) ? num : (num - first.PageSetup.Gutter));
		Globals.ThisAddIn.InsertHeaderFooter(ApplyToSection: true, headerFooter, array, ThisAddIn.HeaderFooterTextType.Center, num, new string[1] { "" }, 0f);
		if (flag)
		{
			Globals.ThisAddIn.InsertSplitLine(headerFooter, isHeader: false, 0);
		}
		if (flag2)
		{
			headerFooter = Globals.ThisAddIn.Application.Selection.Sections.First.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
			if (section != null)
			{
				section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
			}
			Globals.ThisAddIn.InsertHeaderFooter(ApplyToSection: true, headerFooter, array, ThisAddIn.HeaderFooterTextType.Center, num, new string[1] { "" }, 0f);
			if (flag)
			{
				Globals.ThisAddIn.InsertSplitLine(headerFooter, isHeader: false, 0);
			}
		}
	}

	private void SetLevelList_Click(object sender, RibbonControlEventArgs e)
	{
		int levels = (sender as RibbonButton).Name switch
		{
			"Btn_Set2LevelList" => 2, 
			"Btn_Set3LevelList" => 3, 
			"Btn_Set4LevelList" => 4, 
			"Btn_Set5LevelList" => 5, 
			_ => 2, 
		};
		Globals.ThisAddIn.AutoCreateLevelList(levels, 0f, 0f, 0f);
	}

	private void Btn_FastListFormat_Click(object sender, RibbonControlEventArgs e)
	{
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		float listNumIndent = defaultValue.ListNumIndent;
		float listTextIndent = defaultValue.ListTextIndent;
		float listAfterIndent = defaultValue.ListAfterIndent;
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			ListFormat listFormat = Globals.ThisAddIn.Application.Selection.Range.ListFormat;
			if (listFormat.ListType != WdListType.wdListOutlineNumbering && listFormat.ListType != WdListType.wdListNoNumbering)
			{
				Globals.ThisAddIn.ListFormat(listFormat.List, -1, null, listNumIndent, listTextIndent, listAfterIndent);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void ListPunctuation_ButtonClick(object sender, RibbonControlEventArgs e)
	{
		string[] source = new string[11]
		{
			",", ".", ";", "，", "。", "；", ":", "：", "?", "？",
			"、"
		};
		string[] array = Array.Empty<string>();
		string text = (sender as RibbonControl).Name;
		if (text == "Btn_ListPunctuation")
		{
			switch ((sender as RibbonSplitButton).Label)
			{
			case "分号句号":
				text = "Btn_ListSemicolonPeriod";
				break;
			case "逗号句号":
				text = "Btn_ListCommaPeriod";
				break;
			case "全为句号":
				text = "Btn_ListPeriod";
				break;
			case "删除标点":
				text = "Btn_ListNoPunctuation";
				break;
			}
		}
		else
		{
			Btn_ListPunctuation.Label = (sender as RibbonButton).Label;
			Btn_ListPunctuation.Image = (sender as RibbonButton).Image;
		}
		switch (text)
		{
		case "Btn_ListPeriod":
			array = new string[2] { "。", "。" };
			break;
		case "Btn_ListSemicolonPeriod":
			array = new string[2] { "；", "。" };
			break;
		case "Btn_ListNoPunctuation":
			array = new string[2] { "", "" };
			break;
		case "Btn_ListCommaPeriod":
			array = new string[2] { "，", "。" };
			break;
		}
		if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionNormal)
		{
			return;
		}
		ListFormat listFormat = Globals.ThisAddIn.Application.Selection.Paragraphs[1].Range.ListFormat;
		List list = listFormat.List;
		if (list == null || (listFormat.ListType != WdListType.wdListSimpleNumbering && listFormat.ListType != WdListType.wdListBullet))
		{
			return;
		}
		for (int i = 1; i <= list.ListParagraphs.Count; i++)
		{
			int num = list.ListParagraphs[i].Range.End - 1;
			Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
			object Start = num - 1;
			object End = num;
			Range range = activeDocument.Range(ref Start, ref End);
			if (source.Contains(range.Text))
			{
				if (range.Text != ":" && range.Text != "：")
				{
					if (i == list.ListParagraphs.Count)
					{
						range.Text = array[1];
					}
					else
					{
						range.Text = array[0];
					}
				}
			}
			else if (i == list.ListParagraphs.Count)
			{
				range.Text += array[1];
			}
			else
			{
				range.Text += array[0];
			}
		}
	}

	private void Btn_ListStartFromOne_Click(object sender, RibbonControlEventArgs e)
	{
		if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionIP || Globals.ThisAddIn.Application.Selection.Range.ListParagraphs.Count != 1)
		{
			return;
		}
		ListFormat listFormat = Globals.ThisAddIn.Application.Selection.Range.ListFormat;
		float firstLineIndent = Globals.ThisAddIn.Application.Selection.Range.ParagraphFormat.FirstLineIndent;
		float leftIndent = Globals.ThisAddIn.Application.Selection.Range.ParagraphFormat.LeftIndent;
		WdContinue wdContinue = listFormat.CanContinuePreviousList(listFormat.ListTemplate);
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (wdContinue == WdContinue.wdContinueList)
			{
				Globals.ThisAddIn.Application.Selection.Range.ListParagraphs[1].SeparateList();
			}
			foreach (Paragraph listParagraph in Globals.ThisAddIn.Application.Selection.Range.ListParagraphs[1].Range.ListFormat.List.ListParagraphs)
			{
				listParagraph.LeftIndent = leftIndent;
				listParagraph.FirstLineIndent = firstLineIndent;
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Btn_TransToList_Click(object sender, RibbonControlEventArgs e)
	{
		ListFormat listFormat = Globals.ThisAddIn.Application.Selection.Range.ListFormat;
		if (listFormat.ListType == WdListType.wdListOutlineNumbering)
		{
			listFormat.ListTemplate.OutlineNumbered = false;
		}
	}

	private void SetTableFormat_Click(object sender, RibbonControlEventArgs e)
	{
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		bool firstRowBold = false;
		bool firstColumnBold = false;
		bool removeUselessLine = false;
		bool setBorder = false;
		switch ((sender as RibbonButton).Name)
		{
		case "Btn_FirstRowBold":
			firstRowBold = true;
			break;
		case "Btn_FirstColumnBold":
			firstColumnBold = true;
			break;
		case "Btn_RemoveUselessLine":
			removeUselessLine = true;
			break;
		case "Btn_TableThickOutside":
			setBorder = true;
			break;
		}
		Tables tables = (Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.Tables : Globals.ThisAddIn.Application.Selection.Tables);
		if (tables.Count == 0)
		{
			return;
		}
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			foreach (Table item in tables)
			{
				if (item.NestingLevel == 1)
				{
					Globals.ThisAddIn.SetTableFormat(item, firstRowBold, firstColumnBold, setInnerMargin: false, 0, 0f, removeUselessLine, setBorder, defaultValue.TableOutside_LineType, defaultValue.TableOutside_LineWidth);
				}
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void SetTableFormatA_Click(object sender, RibbonControlEventArgs e)
	{
		Tables tables = (Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.Tables : Globals.ThisAddIn.Application.Selection.Tables);
		if (tables.Count == 0)
		{
			return;
		}
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			foreach (Table item in tables)
			{
				if (item.NestingLevel != 1)
				{
					continue;
				}
				string name = (sender as RibbonControl).Name;
				if (!(name == "Btn_TableSingleSpace"))
				{
					if (name == "Btn_RemoveLeftIndent")
					{
						item.Range.ParagraphFormat.LeftIndent = 0f;
						item.Range.ParagraphFormat.FirstLineIndent = 0f;
						item.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0f;
					}
				}
				else
				{
					item.Range.ParagraphFormat.SpaceBefore = 0f;
					item.Range.ParagraphFormat.SpaceAfter = 0f;
					item.Range.ParagraphFormat.Space1();
				}
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void SetTableAlignment_Click(object sender, RibbonControlEventArgs e)
	{
		int alignmentType = 0;
		switch ((sender as RibbonButton).Name)
		{
		case "Btn_TableLeft":
			alignmentType = 0;
			break;
		case "Btn_TableCenter":
			alignmentType = 2;
			break;
		case "Btn_TableRight":
			alignmentType = 1;
			break;
		}
		Tables tables = (Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.Tables : Globals.ThisAddIn.Application.Selection.Tables);
		if (tables.Count == 0)
		{
			return;
		}
		foreach (Table item in tables)
		{
			if (item.NestingLevel == 1)
			{
				Globals.ThisAddIn.SetTableShapeAlignment(item, alignmentType, 0f, 1);
			}
		}
	}

	private void GA_NewTables_Click(object sender, RibbonControlEventArgs e)
	{
		switch (GA_NewTables.SelectedItemIndex)
		{
		case 0:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: false, ThreeLineExtra: false, BroadOuterLine: false, TitleRowFilled: false, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: false, Diagonal: false);
			break;
		case 1:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: false, ThreeLineExtra: false, BroadOuterLine: false, TitleRowFilled: true, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: true, Diagonal: false);
			break;
		case 2:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: false, ThreeLineExtra: false, BroadOuterLine: true, TitleRowFilled: false, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: false, Diagonal: false);
			break;
		case 3:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: false, ThreeLineExtra: false, BroadOuterLine: true, TitleRowFilled: true, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: true, Diagonal: false);
			break;
		case 4:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: false, ThreeLineExtra: false, BroadOuterLine: true, TitleRowFilled: false, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: false, Diagonal: true);
			break;
		case 5:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: false, ThreeLineExtra: false, BroadOuterLine: true, TitleRowFilled: true, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: true, Diagonal: true);
			break;
		case 6:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: true, ThreeLineExtra: false, BroadOuterLine: false, TitleRowFilled: false, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: false, Diagonal: false);
			break;
		case 7:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: true, ThreeLineExtra: false, BroadOuterLine: false, TitleRowFilled: true, SummaryRow: true, SummaryColumn: true, SummaryRowFilled: false, Diagonal: false);
			break;
		case 8:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: true, ThreeLineExtra: true, BroadOuterLine: false, TitleRowFilled: false, SummaryRow: true, SummaryColumn: false, SummaryRowFilled: false, Diagonal: false);
			break;
		case 9:
			Globals.ThisAddIn.CreateNewTable(ThreeLine: true, ThreeLineExtra: true, BroadOuterLine: false, TitleRowFilled: true, SummaryRow: true, SummaryColumn: false, SummaryRowFilled: false, Diagonal: false);
			break;
		}
	}

	private void Btn_TableFullWidth_Click(object sender, RibbonControlEventArgs e)
	{
		Tables tables = (Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.Tables : Globals.ThisAddIn.Application.Selection.Tables);
		if (tables.Count == 0)
		{
			return;
		}
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			foreach (Table item in tables)
			{
				if (item.NestingLevel != 1)
				{
					continue;
				}
				if (item.Rows.LeftIndent != 0f && item.Rows.LeftIndent != 9999999f)
				{
					Section section = item.Range.Sections[1];
					float num = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin - item.Rows.LeftIndent;
					if (!section.PageSetup.GutterOnTop)
					{
						num -= section.PageSetup.Gutter;
					}
					item.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
					item.PreferredWidth = num;
				}
				else
				{
					item.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
					item.PreferredWidth = 100f;
				}
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Btn_RepeatTitle_Click(object sender, RibbonControlEventArgs e)
	{
		Tables tables = (Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.Tables : Globals.ThisAddIn.Application.Selection.Tables);
		if (tables.Count == 0)
		{
			return;
		}
		InputForm inputForm = new InputForm("输入要设置为重复标题的行数（阿拉伯数字）：", "1");
		if (inputForm.ShowDialog() != DialogResult.OK)
		{
			return;
		}
		int num;
		try
		{
			num = Convert.ToInt32(inputForm.InputText);
		}
		catch
		{
			MessageBox.Show("请输入阿拉伯数字，最小为1。", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			return;
		}
		if (num <= 0)
		{
			MessageBox.Show("请输入阿拉伯数字，最小为1。", "Word格式助手", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			return;
		}
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			foreach (Table item in tables)
			{
				try
				{
					foreach (Row row2 in item.Rows)
					{
						row2.HeadingFormat = -1;
						if (row2.Index == num)
						{
							break;
						}
					}
				}
				catch
				{
					int row = num;
					if (num < item.Rows.Count)
					{
						for (int i = num + 1; i <= item.Rows.Count; i++)
						{
							bool flag = true;
							Cell next = item.Cell(i, 1).Next;
							while (next != null && next.RowIndex == i)
							{
								if (next.ColumnIndex - next.Previous.ColumnIndex > 1)
								{
									flag = false;
									break;
								}
								next = next.Next;
							}
							if (flag)
							{
								row = i;
								break;
							}
						}
						Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
						object Start = item.Range.Start;
						object End = item.Cell(row, 1).Range.Start - 1;
						activeDocument.Range(ref Start, ref End).Rows.HeadingFormat = -1;
					}
					else
					{
						item.Range.Rows.HeadingFormat = -1;
					}
				}
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void SetPictureAlignment_Click(object sender, RibbonControlEventArgs e)
	{
		int alignmentType = 0;
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		switch ((sender as RibbonButton).Name)
		{
		case "Btn_PictureLeft":
			alignmentType = 0;
			break;
		case "Btn_PictureCenter":
			alignmentType = 2;
			break;
		case "Btn_PictureRight":
			alignmentType = 1;
			break;
		}
		if (Tog_ApplyToAll.Checked)
		{
			if (Globals.ThisAddIn.Application.ActiveDocument.Shapes.Count != 0)
			{
				foreach (Shape shape3 in Globals.ThisAddIn.Application.ActiveDocument.Shapes)
				{
					if ((shape3.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToShapeLinkedPicture) || (shape3.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (shape3.Type == MsoShapeType.msoSmartArt && defaultValue.ApplyToShapeSmartArt) || (shape3.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart))
					{
						Globals.ThisAddIn.SetTableShapeAlignment(shape3, alignmentType, 0f, 3);
					}
					if (shape3.Type == MsoShapeType.msoGroup)
					{
						if (shape3.Anchor.Text != null && defaultValue.ApplyToInlineGroup)
						{
							Globals.ThisAddIn.SetTableShapeAlignment(shape3, alignmentType, 0f, 4);
						}
						else if (shape3.Anchor.Text == null && defaultValue.ApplyToShapeGroup)
						{
							Globals.ThisAddIn.SetTableShapeAlignment(shape3, alignmentType, 0f, 3);
						}
					}
				}
			}
		}
		else if (Globals.ThisAddIn.Application.Selection.ShapeRange.Count != 0)
		{
			foreach (Shape item in Globals.ThisAddIn.Application.Selection.ShapeRange)
			{
				if ((item.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToShapeLinkedPicture) || (item.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (item.Type == MsoShapeType.msoSmartArt && defaultValue.ApplyToShapeSmartArt) || (item.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart))
				{
					Globals.ThisAddIn.SetTableShapeAlignment(item, alignmentType, 0f, 3);
				}
				if (item.Type == MsoShapeType.msoGroup)
				{
					if (item.Anchor.Text != null && defaultValue.ApplyToInlineGroup)
					{
						Globals.ThisAddIn.SetTableShapeAlignment(item, alignmentType, 0f, 4);
					}
					else if (item.Anchor.Text == null && defaultValue.ApplyToShapeGroup)
					{
						Globals.ThisAddIn.SetTableShapeAlignment(item, alignmentType, 0f, 3);
					}
				}
			}
		}
		InlineShapes inlineShapes = (Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.InlineShapes : Globals.ThisAddIn.Application.Selection.InlineShapes);
		if (inlineShapes.Count == 0)
		{
			return;
		}
		foreach (InlineShape item2 in inlineShapes)
		{
			if ((item2.Type == WdInlineShapeType.wdInlineShapeLinkedPicture && defaultValue.ApplyToInlineLinkedPicture) || (item2.Type == WdInlineShapeType.wdInlineShapePicture && defaultValue.ApplyToInlinePicture) || (item2.Type == WdInlineShapeType.wdInlineShapeSmartArt && defaultValue.ApplyToInlineSmartArt) || (item2.Type == WdInlineShapeType.wdInlineShapeChart && defaultValue.ApplyToInlineChart))
			{
				Globals.ThisAddIn.SetTableShapeAlignment(item2, alignmentType, 0f, 2);
			}
		}
	}

	private void SetPictureFormat_Click(object sender, RibbonControlEventArgs e)
	{
		bool flag = false;
		bool flag2 = false;
		bool setSingleSpace = false;
		float num = 1f;
		float num2 = 1f;
		WordFormatHelperDefault defaultValue = Globals.ThisAddIn.defaultValue;
		switch ((sender as RibbonButton).Name)
		{
		case "Btn_SetPictureWidth":
			flag = true;
			if (Ebox_PictureWidth.Text == "")
			{
				return;
			}
			num = Convert.ToSingle(Ebox_PictureWidth.Text.Replace("厘米", ""));
			break;
		case "Btn_SetPictureHeight":
			flag2 = true;
			if (Ebox_PictureHeight.Text == "")
			{
				return;
			}
			num2 = Convert.ToSingle(Ebox_PictureHeight.Text.Replace("厘米", ""));
			break;
		case "Btn_PictureSingleSpace":
			setSingleSpace = true;
			break;
		}
		foreach (InlineShape item in Tog_ApplyToAll.Checked ? Globals.ThisAddIn.Application.ActiveDocument.InlineShapes : Globals.ThisAddIn.Application.Selection.InlineShapes)
		{
			if ((item.Type == WdInlineShapeType.wdInlineShapeChart && defaultValue.ApplyToInlineChart) || (item.Type == WdInlineShapeType.wdInlineShapeLinkedPicture && defaultValue.ApplyToInlineLinkedPicture) || (item.Type == WdInlineShapeType.wdInlineShapePicture && defaultValue.ApplyToInlinePicture) || (item.Type == WdInlineShapeType.wdInlineShapeSmartArt && defaultValue.ApplyToInlineSmartArt))
			{
				ThisAddIn thisAddIn = Globals.ThisAddIn;
				bool sameWidth = flag;
				float pWidth = num;
				bool sameHeight = flag2;
				float pHeight = num2;
				thisAddIn.SetPictureFormat(item, 0, setSingleSpace, sameWidth, pWidth, sameHeight, pHeight);
			}
		}
		if (!(flag || flag2))
		{
			return;
		}
		if (Tog_ApplyToAll.Checked)
		{
			foreach (Shape shape3 in Globals.ThisAddIn.Application.ActiveDocument.Shapes)
			{
				if ((shape3.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToShapeLinkedPicture) || (shape3.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (shape3.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart) || (shape3.Type == MsoShapeType.msoGroup && defaultValue.ApplyToShapeGroup))
				{
					Globals.ThisAddIn.SetPictureFormat(shape3, 1, setSingleSpace: false, flag, num, flag2, num2);
				}
			}
			return;
		}
		foreach (Shape item2 in Globals.ThisAddIn.Application.Selection.ShapeRange)
		{
			if ((item2.Type == MsoShapeType.msoLinkedPicture && defaultValue.ApplyToShapeLinkedPicture) || (item2.Type == MsoShapeType.msoPicture && defaultValue.ApplyToShapePicture) || (item2.Type == MsoShapeType.msoChart && defaultValue.ApplyToShapeChart) || (item2.Type == MsoShapeType.msoGroup && defaultValue.ApplyToShapeGroup))
			{
				Globals.ThisAddIn.SetPictureFormat(item2, 1, setSingleSpace: false, flag, num, flag2, num2);
			}
		}
	}

	private void NumberValidate(object sender, RibbonControlEventArgs e)
	{
		string text = (sender as RibbonEditBox).Text;
		text = text.Replace(" ", "");
		text = text.Replace("厘米", "");
		if (Regex.IsMatch(text, "\\d{1,}\\.{0,1}\\d{0,}$"))
		{
			(sender as RibbonEditBox).Text = ((float)Math.Round(Convert.ToSingle(text), 2)).ToString("0.00") + " 厘米";
			return;
		}
		MessageBox.Show("无效的数值，请输入正确的数值，保留2位小数。\n单位为厘米，可标明或不标明，例如：3.0或者3.0厘米。", "错误");
		(sender as RibbonEditBox).Text = "3.00厘米";
	}

	private void Btn_PunctuationEng2Chn_Click(object sender, RibbonControlEventArgs e)
	{
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (multiSelection == null)
			{
				if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
				{
					return;
				}
				{
					foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
					{
						ThisAddIn.PunctuationWidthSwitch(cell.Range);
					}
					return;
				}
			}
			foreach (Range item in multiSelection)
			{
				ThisAddIn.PunctuationWidthSwitch(item);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Btn_PunctuationChn2Eng_Click(object sender, RibbonControlEventArgs e)
	{
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (multiSelection == null)
			{
				if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
				{
					return;
				}
				{
					foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
					{
						ThisAddIn.PunctuationWidthSwitch(cell.Range, EngToChn: false);
					}
					return;
				}
			}
			foreach (Range item in multiSelection)
			{
				ThisAddIn.PunctuationWidthSwitch(item, EngToChn: false);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Btn_DeleteSpace_Click(object sender, RibbonControlEventArgs e)
	{
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (multiSelection == null)
			{
				if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
				{
					return;
				}
				{
					foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
					{
						ThisAddIn.RemoveWhiteSpace(cell.Range);
					}
					return;
				}
			}
			foreach (Range item in multiSelection)
			{
				ThisAddIn.RemoveWhiteSpace(item);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Ga_InsertDate_Click(object sender, RibbonControlEventArgs e)
	{
		Range range = Globals.ThisAddIn.Application.Selection.Range;
		object Unit = Type.Missing;
		object Count = Type.Missing;
		range.Delete(ref Unit, ref Count);
		string text = new string('\u3000', 4);
		string text2 = DateTime.Today.Year.ToString(" 0000 ");
		string text3 = DateTime.Today.Month.ToString(" 0 ");
		string text4 = DateTime.Today.Day.ToString(" 0 ");
		string text5 = "";
		switch (Ga_InsertDate.SelectedItemIndex)
		{
		case 0:
			text5 = text2 + "年" + text3 + "月" + text4 + "日";
			break;
		case 1:
			text5 = text2 + "年" + text + "月" + text + "日";
			break;
		case 2:
			text5 = text2 + "年" + text3 + "月" + text + "日";
			break;
		case 3:
			text5 = text + "年" + text + "月" + text + "日";
			break;
		}
		Globals.ThisAddIn.Application.Selection.InsertAfter(text5);
		Document activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
		Count = Globals.ThisAddIn.Application.Selection.Start;
		Unit = Globals.ThisAddIn.Application.Selection.Start + text5.Length;
		Range range2 = activeDocument.Range(ref Count, ref Unit);
		Match match = Regex.Match(range2.Text, "\\u3000{4}");
		while (match.Success)
		{
			Document document = range2.Document;
			Unit = range2.Start + match.Index;
			Count = range2.Start + match.Index + match.Value.Length;
			document.Range(ref Unit, ref Count).FormattedText.Underline = WdUnderline.wdUnderlineSingle;
			match = match.NextMatch();
		}
	}

	private void Ga_Bracketed_Click(object sender, RibbonControlEventArgs e)
	{
		if (multiSelection == null)
		{
			return;
		}
		foreach (Range item in multiSelection)
		{
			ThisAddIn.AddBrakets(item, Ga_Bracketed.SelectedItemIndex, ThisAddIn.textFormatSet.RemoveBrackets);
		}
	}

	private void Btn_ToChinese_Click(object sender, RibbonControlEventArgs e)
	{
		if (multiSelection == null)
		{
			return;
		}
		TransToChineseNumbers transToChineseNumbers = new TransToChineseNumbers();
		foreach (Range item in multiSelection)
		{
			string text = transToChineseNumbers.ToChineseNumber(item.Text, ChineseTraditional: true);
			if (!text.StartsWith("转换") && !text.StartsWith("输入"))
			{
				if (text.Contains("點"))
				{
					string[] array = new string[3] { "角", "分", "厘" };
					string text2 = text.Split('點')[1];
					text = text.Split('點')[0];
					text += "元";
					for (int i = 0; i < 3 && i != text2.Length; i++)
					{
						string text3 = text2.Substring(i, 1);
						text = ((!(text3 != "零") || !(text3 != "")) ? (text + text2.Substring(i, 1)) : (text + text3 + array[i]));
					}
					text = text.TrimEnd('零');
					if (text.EndsWith("元"))
					{
						text += "整";
					}
				}
				else
				{
					text += "元整";
				}
			}
			if (text.Contains("萬"))
			{
				text = text.Replace("萬", "万");
			}
			if (text.Contains("億"))
			{
				text = text.Replace("億", "亿");
			}
			item.Text = text;
		}
	}

	private void ParagrahIndetSet_ButtonClick(object sender, RibbonControlEventArgs e)
	{
		string text = (sender as RibbonControl).Name;
		if (text == "Btn_ParagrahIndent")
		{
			string label = (sender as RibbonSplitButton).Label;
			if (!(label == "缩进两字"))
			{
				if (label == "移除缩进")
				{
					text = "Btn_ParagrahNoIndent";
				}
			}
			else
			{
				text = "Btn_ParagrahIndent2Char";
			}
		}
		else
		{
			Btn_ParagrahIndent.Label = (sender as RibbonButton).Label;
			Btn_ParagrahIndent.Image = (sender as RibbonButton).Image;
		}
		bool setIndent = text == "Btn_ParagrahIndent2Char";
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (Globals.ThisAddIn.Application.Selection.Type == WdSelectionType.wdSelectionNormal)
			{
				ThisAddIn.SetIndent2CharOrNot(Globals.ThisAddIn.Application.Selection.Range, setIndent);
			}
			else
			{
				if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
				{
					return;
				}
				{
					foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
					{
						ThisAddIn.SetIndent2CharOrNot(cell.Range, setIndent);
					}
					return;
				}
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Btn_DeleteBlankLine_Click(object sender, RibbonControlEventArgs e)
	{
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (multiSelection != null)
			{
				foreach (Range item in multiSelection)
				{
					ThisAddIn.RemoveSpaceLines(item);
				}
				return;
			}
			if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
			{
				return;
			}
			foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
			{
				ThisAddIn.RemoveSpaceLines(cell.Range);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void Btn_FormatPainter_Click(object sender, RibbonControlEventArgs e)
	{
		ApplyFixFormatPainter();
	}

	public void ApplyFixFormatPainter()
	{
		FixFormatPainterSetting.FixFormat format = ThisAddIn.formatPainter.StoredFormat[ThisAddIn.formatPainter.CurrentID];
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (multiSelection != null)
			{
				foreach (Range item in multiSelection)
				{
					ApplyFixPainter(item, format);
				}
				return;
			}
			if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
			{
				return;
			}
			foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
			{
				ApplyFixPainter(cell.Range, format);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void ApplyFixPainter(Range target, FixFormatPainterSetting.FixFormat format)
	{
		target.Font.Name = format.EngFontName;
		target.Font.NameFarEast = format.ChnFontName;
		target.Font.Size = format.FontSize;
		target.Font.Bold = (format.Bold ? (-1) : 0);
		target.Font.Italic = (format.Italic ? (-1) : 0);
		target.Font.Underline = (format.Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
		if (format.UseColor)
		{
			Color color = Color.FromArgb(format.TextColor);
			target.Font.Color = (WdColor)ThisAddIn.RGB(color.R, color.G, color.B);
		}
		if (format.Shading)
		{
			Color color2 = Color.FromArgb(format.ShadingColor);
			target.Shading.Texture = WdTextureIndex.wdTextureSolid;
			target.Shading.ForegroundPatternColor = (WdColor)ThisAddIn.RGB(color2.R, color2.G, color2.B);
		}
	}

	private void ScriptFormat_Click(object sender, RibbonControlEventArgs e)
	{
		string text;
		bool flag;
		if ((sender as RibbonControl).Name == "Btn_SuperscriptAndSubscript")
		{
			text = (sender as RibbonSplitButton).Label switch
			{
				"平方上标" => "Btn_SquareSuperscript", 
				"立方上标" => "Btn_CubeSuperscript", 
				"数字上标" => "Btn_NumberSuperscript", 
				"数字下标" => "Btn_NumberSubscript", 
				_ => "", 
			};
			flag = false;
		}
		else
		{
			text = (sender as RibbonControl).Name;
			flag = true;
		}
		if (flag)
		{
			ToggleScriptBtn(sender as RibbonButton);
		}
		string text2;
		switch (text)
		{
		case "Btn_SquareSuperscript":
			text2 = "(?<=[mM])2";
			break;
		case "Btn_CubeSuperscript":
			text2 = "(?<=[mM])3";
			break;
		case "Btn_NumberSuperscript":
		case "Btn_NumberSubscript":
			text2 = "[0-9]+";
			break;
		default:
			text2 = "";
			break;
		}
		string scriptText = text2;
		bool superscript = text != "Btn_NumberSubscript";
		try
		{
			Globals.ThisAddIn.Application.ScreenUpdating = false;
			if (multiSelection != null)
			{
				foreach (Range item in multiSelection)
				{
					ThisAddIn.SetSuperscriptOrSubscript(item, scriptText, superscript, useRegex: true);
				}
				return;
			}
			if (Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionColumn && Globals.ThisAddIn.Application.Selection.Type != WdSelectionType.wdSelectionRow)
			{
				return;
			}
			foreach (Cell cell in Globals.ThisAddIn.Application.Selection.Cells)
			{
				ThisAddIn.SetSuperscriptOrSubscript(cell.Range, scriptText, superscript, useRegex: true);
			}
		}
		finally
		{
			Globals.ThisAddIn.Application.ScreenUpdating = true;
		}
	}

	private void ToggleScriptBtn(RibbonButton btn)
	{
		Btn_SuperscriptAndSubscript.Label = btn.Label;
		Btn_SuperscriptAndSubscript.ScreenTip = btn.ScreenTip;
		Btn_SuperscriptAndSubscript.SuperTip = btn.SuperTip;
		Btn_SuperscriptAndSubscript.Image = btn.Image;
	}

	public WordFormatHelperRibbon()
		: base(Globals.Factory.GetRibbonFactory())
	{
		InitializeComponent();
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
		RibbonDialogLauncher dialogLauncher = base.Factory.CreateRibbonDialogLauncher();
		RibbonDropDownItem ribbonDropDownItem = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem2 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem3 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem4 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem5 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem6 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem7 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem8 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem9 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem10 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem11 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem12 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem13 = base.Factory.CreateRibbonDropDownItem();
		RibbonDialogLauncher dialogLauncher2 = base.Factory.CreateRibbonDialogLauncher();
		RibbonDropDownItem ribbonDropDownItem14 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem15 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem16 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem17 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem18 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem19 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem20 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem21 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem22 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem23 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem24 = base.Factory.CreateRibbonDropDownItem();
		RibbonDialogLauncher dialogLauncher3 = base.Factory.CreateRibbonDialogLauncher();
		RibbonDialogLauncher dialogLauncher4 = base.Factory.CreateRibbonDialogLauncher();
		RibbonDialogLauncher dialogLauncher5 = base.Factory.CreateRibbonDialogLauncher();
		RibbonDropDownItem ribbonDropDownItem25 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem26 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem27 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem28 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem29 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem30 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem31 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem32 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem33 = base.Factory.CreateRibbonDropDownItem();
		RibbonDropDownItem ribbonDropDownItem34 = base.Factory.CreateRibbonDropDownItem();
		RibbonDialogLauncher dialogLauncher6 = base.Factory.CreateRibbonDialogLauncher();
		TabAddIns = base.Factory.CreateRibbonTab();
		Tab_WordFormatAssistant = base.Factory.CreateRibbonTab();
		Gp_PageSet = base.Factory.CreateRibbonGroup();
		Ga_FastMargin = base.Factory.CreateRibbonGallery();
		Ga_FastHeaderFooter = base.Factory.CreateRibbonGallery();
		Ga_HeaderLine = base.Factory.CreateRibbonGallery();
		Btn_SetHeaderSingleLine = base.Factory.CreateRibbonButton();
		Btn_SetHeaderThickLine = base.Factory.CreateRibbonButton();
		Btn_SetHeaderDoubleLine = base.Factory.CreateRibbonButton();
		Btn_SetHeaderThinThickLine = base.Factory.CreateRibbonButton();
		Btn_SetHeaderThickThinLine = base.Factory.CreateRibbonButton();
		Btn_DeleteHeaderLine = base.Factory.CreateRibbonButton();
		Ga_FooterLine = base.Factory.CreateRibbonGallery();
		Btn_SetFooterSingleLine = base.Factory.CreateRibbonButton();
		Btn_SetFooterThickLine = base.Factory.CreateRibbonButton();
		Btn_SetFooterDoubleLine = base.Factory.CreateRibbonButton();
		Btn_SetFooterThinThickLine = base.Factory.CreateRibbonButton();
		Btn_SetFooterThickThinLine = base.Factory.CreateRibbonButton();
		Btn_DeleteFooterLine = base.Factory.CreateRibbonButton();
		Gp_TextFormatSet = base.Factory.CreateRibbonGroup();
		Btn_PunctuationEng2Chn = base.Factory.CreateRibbonButton();
		Btn_PunctuationChn2Eng = base.Factory.CreateRibbonButton();
		Btn_DeleteSpace = base.Factory.CreateRibbonButton();
		Ga_InsertDate = base.Factory.CreateRibbonGallery();
		Ga_Bracketed = base.Factory.CreateRibbonGallery();
		Btn_DeleteBlankLine = base.Factory.CreateRibbonButton();
		Btn_ToChinese = base.Factory.CreateRibbonButton();
		Btn_ParagrahIndent = base.Factory.CreateRibbonSplitButton();
		Btn_ParagrahIndent2Char = base.Factory.CreateRibbonButton();
		Btn_ParagrahNoIndent = base.Factory.CreateRibbonButton();
		Btn_FormatPainter = base.Factory.CreateRibbonSplitButton();
		Btn_FormatPainterUI = base.Factory.CreateRibbonButton();
		Btn_SuperscriptAndSubscript = base.Factory.CreateRibbonSplitButton();
		Btn_SquareSuperscript = base.Factory.CreateRibbonButton();
		Btn_CubeSuperscript = base.Factory.CreateRibbonButton();
		Btn_NumberSuperscript = base.Factory.CreateRibbonButton();
		Btn_NumberSubscript = base.Factory.CreateRibbonButton();
		Btn_CustomScript = base.Factory.CreateRibbonButton();
		Gp_List = base.Factory.CreateRibbonGroup();
		Btn_FastListFormat = base.Factory.CreateRibbonButton();
		Btn_ListPunctuation = base.Factory.CreateRibbonSplitButton();
		Btn_ListSemicolonPeriod = base.Factory.CreateRibbonButton();
		Btn_ListCommaPeriod = base.Factory.CreateRibbonButton();
		Btn_ListPeriod = base.Factory.CreateRibbonButton();
		Btn_ListNoPunctuation = base.Factory.CreateRibbonButton();
		Btn_ListStartFromOne = base.Factory.CreateRibbonButton();
		Btn_TransToList = base.Factory.CreateRibbonButton();
		Gp_LevelList = base.Factory.CreateRibbonGroup();
		Ga_FastLevelList = base.Factory.CreateRibbonGallery();
		Btn_Set2LevelList = base.Factory.CreateRibbonButton();
		Btn_Set3LevelList = base.Factory.CreateRibbonButton();
		Btn_Set4LevelList = base.Factory.CreateRibbonButton();
		Btn_Set5LevelList = base.Factory.CreateRibbonButton();
		Gp_TOC = base.Factory.CreateRibbonGroup();
		Btn_TOC = base.Factory.CreateRibbonButton();
		Gp_TablePictureSet = base.Factory.CreateRibbonGroup();
		Tog_ApplyToAll = base.Factory.CreateRibbonToggleButton();
		separator2 = base.Factory.CreateRibbonSeparator();
		Btn_TableLeft = base.Factory.CreateRibbonButton();
		Btn_TableCenter = base.Factory.CreateRibbonButton();
		Btn_TableRight = base.Factory.CreateRibbonButton();
		Btn_FirstRowBold = base.Factory.CreateRibbonButton();
		Btn_FirstColumnBold = base.Factory.CreateRibbonButton();
		Btn_TableThickOutside = base.Factory.CreateRibbonButton();
		Btn_TableSingleSpace = base.Factory.CreateRibbonButton();
		Btn_RemoveUselessLine = base.Factory.CreateRibbonButton();
		Btn_RemoveLeftIndent = base.Factory.CreateRibbonButton();
		Btn_TableFullWidth = base.Factory.CreateRibbonButton();
		Btn_RowHeightFitText = base.Factory.CreateRibbonButton();
		Btn_RepeatTitle = base.Factory.CreateRibbonButton();
		GA_NewTables = base.Factory.CreateRibbonGallery();
		Btn_TableStyle = base.Factory.CreateRibbonButton();
		separator1 = base.Factory.CreateRibbonSeparator();
		Btn_PictureLeft = base.Factory.CreateRibbonButton();
		Btn_PictureCenter = base.Factory.CreateRibbonButton();
		Btn_PictureRight = base.Factory.CreateRibbonButton();
		Btn_SetPictureWidth = base.Factory.CreateRibbonButton();
		Btn_SetPictureHeight = base.Factory.CreateRibbonButton();
		Btn_PictureSingleSpace = base.Factory.CreateRibbonButton();
		Ebox_PictureWidth = base.Factory.CreateRibbonEditBox();
		Ebox_PictureHeight = base.Factory.CreateRibbonEditBox();
		Gp_DocumentSet = base.Factory.CreateRibbonGroup();
		Btn_FastSetStyle = base.Factory.CreateRibbonButton();
		Gp_QRCoder = base.Factory.CreateRibbonGroup();
		Btn_FastQRCoder = base.Factory.CreateRibbonButton();
		Gp_Help = base.Factory.CreateRibbonGroup();
		Ga_Utilities = base.Factory.CreateRibbonGallery();
		Btn_ExportPDF = base.Factory.CreateRibbonButton();
		Btn_OCR = base.Factory.CreateRibbonButton();
		Btn_ExportPng = base.Factory.CreateRibbonButton();
		Ga_SetUnits = base.Factory.CreateRibbonGallery();
		Btn_UnitInch = base.Factory.CreateRibbonButton();
		Btn_UnitCM = base.Factory.CreateRibbonButton();
		Btn_UnitMM = base.Factory.CreateRibbonButton();
		Btn_UnitPt = base.Factory.CreateRibbonButton();
		Btn_UnitPicas = base.Factory.CreateRibbonButton();
		Btn_DefaultValue = base.Factory.CreateRibbonButton();
		GA_AboutAndHelp = base.Factory.CreateRibbonGallery();
		Btn_HelpOnline = base.Factory.CreateRibbonButton();
		Btn_AboutUs = base.Factory.CreateRibbonButton();
		TabAddIns.SuspendLayout();
		Tab_WordFormatAssistant.SuspendLayout();
		Gp_PageSet.SuspendLayout();
		Gp_TextFormatSet.SuspendLayout();
		Gp_List.SuspendLayout();
		Gp_LevelList.SuspendLayout();
		Gp_TOC.SuspendLayout();
		Gp_TablePictureSet.SuspendLayout();
		Gp_DocumentSet.SuspendLayout();
		Gp_QRCoder.SuspendLayout();
		Gp_Help.SuspendLayout();
		SuspendLayout();
		TabAddIns.ControlId.ControlIdType = RibbonControlIdType.Office;
		TabAddIns.Label = "TabAddIns";
		TabAddIns.Name = "TabAddIns";
		Tab_WordFormatAssistant.Groups.Add(Gp_PageSet);
		Tab_WordFormatAssistant.Groups.Add(Gp_TextFormatSet);
		Tab_WordFormatAssistant.Groups.Add(Gp_List);
		Tab_WordFormatAssistant.Groups.Add(Gp_LevelList);
		Tab_WordFormatAssistant.Groups.Add(Gp_TOC);
		Tab_WordFormatAssistant.Groups.Add(Gp_TablePictureSet);
		Tab_WordFormatAssistant.Groups.Add(Gp_DocumentSet);
		Tab_WordFormatAssistant.Groups.Add(Gp_QRCoder);
		Tab_WordFormatAssistant.Groups.Add(Gp_Help);
		Tab_WordFormatAssistant.Label = "Word格式助手";
		Tab_WordFormatAssistant.Name = "Tab_WordFormatAssistant";
		Gp_PageSet.DialogLauncher = dialogLauncher;
		Gp_PageSet.Items.Add(Ga_FastMargin);
		Gp_PageSet.Items.Add(Ga_FastHeaderFooter);
		Gp_PageSet.Items.Add(Ga_HeaderLine);
		Gp_PageSet.Items.Add(Ga_FooterLine);
		Gp_PageSet.Label = "页边距与页眉页脚";
		Gp_PageSet.Name = "Gp_PageSet";
		Gp_PageSet.DialogLauncherClick += ShowFormatHelperUI;
		Ga_FastMargin.ColumnCount = 1;
		Ga_FastMargin.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Ga_FastMargin.Image = Resources.FastMargins;
		Ga_FastMargin.ItemImageSize = new Size(32, 32);
		ribbonDropDownItem.Image = Resources.FastMarginsDefault;
		ribbonDropDownItem.Label = "设置默认页边距";
		ribbonDropDownItem2.Image = Resources.FastMargins20;
		ribbonDropDownItem2.Label = "设置2.0厘米页边距";
		ribbonDropDownItem3.Image = Resources.FastMargins25;
		ribbonDropDownItem3.Label = "设置2.5厘米页边距";
		ribbonDropDownItem4.Image = Resources.FastMargins30;
		ribbonDropDownItem4.Label = "设置3.0厘米页边距";
		ribbonDropDownItem5.Image = Resources.FastMargins35;
		ribbonDropDownItem5.Label = "设置3.5厘米页边距";
		Ga_FastMargin.Items.Add(ribbonDropDownItem);
		Ga_FastMargin.Items.Add(ribbonDropDownItem2);
		Ga_FastMargin.Items.Add(ribbonDropDownItem3);
		Ga_FastMargin.Items.Add(ribbonDropDownItem4);
		Ga_FastMargin.Items.Add(ribbonDropDownItem5);
		Ga_FastMargin.Label = "快速页边距";
		Ga_FastMargin.Name = "Ga_FastMargin";
		Ga_FastMargin.ScreenTip = "快速页边距";
		Ga_FastMargin.ShowImage = true;
		Ga_FastMargin.Click += SetMarginBtn_Click;
		Ga_FastHeaderFooter.ColumnCount = 1;
		Ga_FastHeaderFooter.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Ga_FastHeaderFooter.Image = Resources.FastHeaderFooter;
		Ga_FastHeaderFooter.ItemImageSize = new Size(256, 42);
		ribbonDropDownItem6.Image = Resources.HeaderStyle01;
		ribbonDropDownItem6.Label = "样式1";
		ribbonDropDownItem6.Tag = "";
		ribbonDropDownItem7.Image = Resources.HeaderStyle02;
		ribbonDropDownItem7.Label = "样式2";
		ribbonDropDownItem8.Image = Resources.HeaderStyle03;
		ribbonDropDownItem8.Label = "样式3";
		ribbonDropDownItem9.Image = Resources.HeaderStyle04;
		ribbonDropDownItem9.Label = "样式4";
		ribbonDropDownItem10.Image = Resources.HeaderStyle05;
		ribbonDropDownItem10.Label = "样式5";
		ribbonDropDownItem11.Image = Resources.HeaderStyle06;
		ribbonDropDownItem11.Label = "样式6";
		ribbonDropDownItem12.Image = Resources.HeaderStyle07;
		ribbonDropDownItem12.Label = "样式7";
		ribbonDropDownItem13.Image = Resources.HeaderStyle08;
		ribbonDropDownItem13.Label = "样式8";
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem6);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem7);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem8);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem9);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem10);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem11);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem12);
		Ga_FastHeaderFooter.Items.Add(ribbonDropDownItem13);
		Ga_FastHeaderFooter.Label = "快速页脚样式";
		Ga_FastHeaderFooter.Name = "Ga_FastHeaderFooter";
		Ga_FastHeaderFooter.ScreenTip = "快速页脚样式";
		Ga_FastHeaderFooter.ShowImage = true;
		Ga_FastHeaderFooter.ShowItemLabel = false;
		Ga_FastHeaderFooter.Click += Ga_FastHeaderFooter_Click;
		Ga_HeaderLine.Buttons.Add(Btn_SetHeaderSingleLine);
		Ga_HeaderLine.Buttons.Add(Btn_SetHeaderThickLine);
		Ga_HeaderLine.Buttons.Add(Btn_SetHeaderDoubleLine);
		Ga_HeaderLine.Buttons.Add(Btn_SetHeaderThinThickLine);
		Ga_HeaderLine.Buttons.Add(Btn_SetHeaderThickThinLine);
		Ga_HeaderLine.Buttons.Add(Btn_DeleteHeaderLine);
		Ga_HeaderLine.Image = Resources.FastHeaderLine;
		Ga_HeaderLine.Label = "页眉分隔线";
		Ga_HeaderLine.Name = "Ga_HeaderLine";
		Ga_HeaderLine.ScreenTip = "添加页眉分隔线";
		Ga_HeaderLine.ShowImage = true;
		Ga_HeaderLine.ButtonClick += SetHeaderFooterLine_Click;
		Btn_SetHeaderSingleLine.Image = Resources.Line_Thin;
		Btn_SetHeaderSingleLine.Label = "细实线";
		Btn_SetHeaderSingleLine.Name = "Btn_SetHeaderSingleLine";
		Btn_SetHeaderSingleLine.ShowImage = true;
		Btn_SetHeaderThickLine.Image = Resources.Line_Thick;
		Btn_SetHeaderThickLine.Label = "粗实线";
		Btn_SetHeaderThickLine.Name = "Btn_SetHeaderThickLine";
		Btn_SetHeaderThickLine.ShowImage = true;
		Btn_SetHeaderDoubleLine.Image = Resources.Line_Double;
		Btn_SetHeaderDoubleLine.Label = "双实线";
		Btn_SetHeaderDoubleLine.Name = "Btn_SetHeaderDoubleLine";
		Btn_SetHeaderDoubleLine.ShowImage = true;
		Btn_SetHeaderThinThickLine.Image = Resources.Line_ThinThick;
		Btn_SetHeaderThinThickLine.Label = "细粗复合线";
		Btn_SetHeaderThinThickLine.Name = "Btn_SetHeaderThinThickLine";
		Btn_SetHeaderThinThickLine.ShowImage = true;
		Btn_SetHeaderThickThinLine.Image = Resources.Line_ThickThin;
		Btn_SetHeaderThickThinLine.Label = "粗细复合线";
		Btn_SetHeaderThickThinLine.Name = "Btn_SetHeaderThickThinLine";
		Btn_SetHeaderThickThinLine.ShowImage = true;
		Btn_DeleteHeaderLine.Image = Resources.Remove;
		Btn_DeleteHeaderLine.Label = "删除页眉分隔线";
		Btn_DeleteHeaderLine.Name = "Btn_DeleteHeaderLine";
		Btn_DeleteHeaderLine.ShowImage = true;
		Ga_FooterLine.Buttons.Add(Btn_SetFooterSingleLine);
		Ga_FooterLine.Buttons.Add(Btn_SetFooterThickLine);
		Ga_FooterLine.Buttons.Add(Btn_SetFooterDoubleLine);
		Ga_FooterLine.Buttons.Add(Btn_SetFooterThinThickLine);
		Ga_FooterLine.Buttons.Add(Btn_SetFooterThickThinLine);
		Ga_FooterLine.Buttons.Add(Btn_DeleteFooterLine);
		Ga_FooterLine.Image = Resources.FastFooterLine;
		Ga_FooterLine.Label = "页脚分隔线";
		Ga_FooterLine.Name = "Ga_FooterLine";
		Ga_FooterLine.ScreenTip = "添加页脚分隔线";
		Ga_FooterLine.ShowImage = true;
		Ga_FooterLine.ButtonClick += SetHeaderFooterLine_Click;
		Btn_SetFooterSingleLine.Image = Resources.Line_Thin;
		Btn_SetFooterSingleLine.Label = "细实线";
		Btn_SetFooterSingleLine.Name = "Btn_SetFooterSingleLine";
		Btn_SetFooterSingleLine.ShowImage = true;
		Btn_SetFooterThickLine.Image = Resources.Line_Thick;
		Btn_SetFooterThickLine.Label = "粗实线";
		Btn_SetFooterThickLine.Name = "Btn_SetFooterThickLine";
		Btn_SetFooterThickLine.ShowImage = true;
		Btn_SetFooterDoubleLine.Image = Resources.Line_Double;
		Btn_SetFooterDoubleLine.Label = "双实线";
		Btn_SetFooterDoubleLine.Name = "Btn_SetFooterDoubleLine";
		Btn_SetFooterDoubleLine.ShowImage = true;
		Btn_SetFooterThinThickLine.Image = Resources.Line_ThinThick;
		Btn_SetFooterThinThickLine.Label = "细粗复合线";
		Btn_SetFooterThinThickLine.Name = "Btn_SetFooterThinThickLine";
		Btn_SetFooterThinThickLine.ShowImage = true;
		Btn_SetFooterThickThinLine.Image = Resources.Line_ThickThin;
		Btn_SetFooterThickThinLine.Label = "粗细复核线";
		Btn_SetFooterThickThinLine.Name = "Btn_SetFooterThickThinLine";
		Btn_SetFooterThickThinLine.ShowImage = true;
		Btn_DeleteFooterLine.Image = Resources.Remove;
		Btn_DeleteFooterLine.Label = "删除页脚分隔线";
		Btn_DeleteFooterLine.Name = "Btn_DeleteFooterLine";
		Btn_DeleteFooterLine.ShowImage = true;
		Gp_TextFormatSet.DialogLauncher = dialogLauncher2;
		Gp_TextFormatSet.Items.Add(Btn_PunctuationEng2Chn);
		Gp_TextFormatSet.Items.Add(Btn_PunctuationChn2Eng);
		Gp_TextFormatSet.Items.Add(Btn_DeleteSpace);
		Gp_TextFormatSet.Items.Add(Ga_InsertDate);
		Gp_TextFormatSet.Items.Add(Ga_Bracketed);
		Gp_TextFormatSet.Items.Add(Btn_DeleteBlankLine);
		Gp_TextFormatSet.Items.Add(Btn_ToChinese);
		Gp_TextFormatSet.Items.Add(Btn_ParagrahIndent);
		Gp_TextFormatSet.Items.Add(Btn_FormatPainter);
		Gp_TextFormatSet.Items.Add(Btn_SuperscriptAndSubscript);
		Gp_TextFormatSet.Label = "文本格式";
		Gp_TextFormatSet.Name = "Gp_TextFormatSet";
		Gp_TextFormatSet.DialogLauncherClick += ShowFormatHelperUI;
		Btn_PunctuationEng2Chn.Image = Resources.ChinesePunc;
		Btn_PunctuationEng2Chn.Label = "转中文";
		Btn_PunctuationEng2Chn.Name = "Btn_PunctuationEng2Chn";
		Btn_PunctuationEng2Chn.ScreenTip = "转全角标点符号";
		Btn_PunctuationEng2Chn.ShowImage = true;
		Btn_PunctuationEng2Chn.Click += Btn_PunctuationEng2Chn_Click;
		Btn_PunctuationChn2Eng.Image = Resources.EnglishPunc;
		Btn_PunctuationChn2Eng.Label = "转西文";
		Btn_PunctuationChn2Eng.Name = "Btn_PunctuationChn2Eng";
		Btn_PunctuationChn2Eng.ScreenTip = "转半角标点符号";
		Btn_PunctuationChn2Eng.ShowImage = true;
		Btn_PunctuationChn2Eng.Click += Btn_PunctuationChn2Eng_Click;
		Btn_DeleteSpace.Image = Resources.Remove;
		Btn_DeleteSpace.Label = "删空格";
		Btn_DeleteSpace.Name = "Btn_DeleteSpace";
		Btn_DeleteSpace.OfficeImageId = "Delete";
		Btn_DeleteSpace.ScreenTip = "删空格";
		Btn_DeleteSpace.ShowImage = true;
		Btn_DeleteSpace.Click += Btn_DeleteSpace_Click;
		Ga_InsertDate.ColumnCount = 1;
		Ga_InsertDate.Image = Resources.DateInput;
		Ga_InsertDate.ItemImageSize = new Size(150, 30);
		ribbonDropDownItem14.Image = Resources.DateInput01;
		ribbonDropDownItem14.Label = "Item0";
		ribbonDropDownItem15.Image = Resources.DateInput02;
		ribbonDropDownItem15.Label = "Item1";
		ribbonDropDownItem16.Image = Resources.DateInput03;
		ribbonDropDownItem16.Label = "Item2";
		ribbonDropDownItem17.Image = Resources.DateInput04;
		ribbonDropDownItem17.Label = "Item3";
		Ga_InsertDate.Items.Add(ribbonDropDownItem14);
		Ga_InsertDate.Items.Add(ribbonDropDownItem15);
		Ga_InsertDate.Items.Add(ribbonDropDownItem16);
		Ga_InsertDate.Items.Add(ribbonDropDownItem17);
		Ga_InsertDate.Label = "日期";
		Ga_InsertDate.Name = "Ga_InsertDate";
		Ga_InsertDate.OfficeImageId = "DateAndTimeInsert";
		Ga_InsertDate.ScreenTip = "输入日期标签";
		Ga_InsertDate.ShowImage = true;
		Ga_InsertDate.ShowItemLabel = false;
		Ga_InsertDate.Click += Ga_InsertDate_Click;
		Ga_Bracketed.ColumnCount = 3;
		Ga_Bracketed.Image = Resources.Brackets;
		Ga_Bracketed.ItemImageSize = new Size(32, 32);
		ribbonDropDownItem18.Image = Resources.BracketsStyle01;
		ribbonDropDownItem18.Label = "双引号";
		ribbonDropDownItem18.ScreenTip = "双引号";
		ribbonDropDownItem19.Image = Resources.BracketsStyle02;
		ribbonDropDownItem19.Label = "书名号";
		ribbonDropDownItem19.ScreenTip = "书名号";
		ribbonDropDownItem20.Image = Resources.BracketsStyle03;
		ribbonDropDownItem20.Label = "圆括号";
		ribbonDropDownItem20.ScreenTip = "圆括号";
		ribbonDropDownItem21.Image = Resources.BracketsStyle04;
		ribbonDropDownItem21.Label = "方括号";
		ribbonDropDownItem21.ScreenTip = "方括号";
		ribbonDropDownItem22.Image = Resources.BracketsStyle05;
		ribbonDropDownItem22.Label = "花括号";
		ribbonDropDownItem22.ScreenTip = "花括号";
		ribbonDropDownItem23.Image = Resources.BracketsStyle06;
		ribbonDropDownItem23.Label = "尖括号";
		ribbonDropDownItem23.ScreenTip = "尖括号";
		ribbonDropDownItem24.Image = Resources.BracketsStyle07;
		ribbonDropDownItem24.Label = "壳形括号";
		ribbonDropDownItem24.ScreenTip = "壳形括号";
		Ga_Bracketed.Items.Add(ribbonDropDownItem18);
		Ga_Bracketed.Items.Add(ribbonDropDownItem19);
		Ga_Bracketed.Items.Add(ribbonDropDownItem20);
		Ga_Bracketed.Items.Add(ribbonDropDownItem21);
		Ga_Bracketed.Items.Add(ribbonDropDownItem22);
		Ga_Bracketed.Items.Add(ribbonDropDownItem23);
		Ga_Bracketed.Items.Add(ribbonDropDownItem24);
		Ga_Bracketed.Label = "括号";
		Ga_Bracketed.Name = "Ga_Bracketed";
		Ga_Bracketed.OfficeImageId = "EquationDelimiterGallery";
		Ga_Bracketed.ScreenTip = "添加括号或引号";
		Ga_Bracketed.ShowImage = true;
		Ga_Bracketed.ShowItemLabel = false;
		Ga_Bracketed.Click += Ga_Bracketed_Click;
		Btn_DeleteBlankLine.Image = Resources.RemoveBlankLines;
		Btn_DeleteBlankLine.Label = "删空行";
		Btn_DeleteBlankLine.Name = "Btn_DeleteBlankLine";
		Btn_DeleteBlankLine.ScreenTip = "删空行";
		Btn_DeleteBlankLine.ShowImage = true;
		Btn_DeleteBlankLine.Click += Btn_DeleteBlankLine_Click;
		Btn_ToChinese.Image = Resources.AmountInWords;
		Btn_ToChinese.Label = "金额大写";
		Btn_ToChinese.Name = "Btn_ToChinese";
		Btn_ToChinese.OfficeImageId = "InternationalCurrency";
		Btn_ToChinese.ScreenTip = "金额大写";
		Btn_ToChinese.ShowImage = true;
		Btn_ToChinese.Click += Btn_ToChinese_Click;
		Btn_ParagrahIndent.Image = Resources.ParagrahIndent2Char;
		Btn_ParagrahIndent.Items.Add(Btn_ParagrahIndent2Char);
		Btn_ParagrahIndent.Items.Add(Btn_ParagrahNoIndent);
		Btn_ParagrahIndent.Label = "缩进两字";
		Btn_ParagrahIndent.Name = "Btn_ParagrahIndent";
		Btn_ParagrahIndent.ScreenTip = "首行缩进";
		Btn_ParagrahIndent.Click += ParagrahIndetSet_ButtonClick;
		Btn_ParagrahIndent2Char.Image = Resources.ParagrahIndent2Char;
		Btn_ParagrahIndent2Char.Label = "缩进两字";
		Btn_ParagrahIndent2Char.Name = "Btn_ParagrahIndent2Char";
		Btn_ParagrahIndent2Char.ShowImage = true;
		Btn_ParagrahIndent2Char.Click += ParagrahIndetSet_ButtonClick;
		Btn_ParagrahNoIndent.Image = Resources.ParagrahNoIndent;
		Btn_ParagrahNoIndent.Label = "移除缩进";
		Btn_ParagrahNoIndent.Name = "Btn_ParagrahNoIndent";
		Btn_ParagrahNoIndent.ShowImage = true;
		Btn_ParagrahNoIndent.Click += ParagrahIndetSet_ButtonClick;
		Btn_FormatPainter.Image = Resources.FixFormatPainter;
		Btn_FormatPainter.Items.Add(Btn_FormatPainterUI);
		Btn_FormatPainter.Label = "定格式刷";
		Btn_FormatPainter.Name = "Btn_FormatPainter";
		Btn_FormatPainter.ScreenTip = "固定格式刷";
		Btn_FormatPainter.Click += Btn_FormatPainter_Click;
		Btn_FormatPainterUI.Image = Resources.FixFormatPainter;
		Btn_FormatPainterUI.Label = "格式设置";
		Btn_FormatPainterUI.Name = "Btn_FormatPainterUI";
		Btn_FormatPainterUI.ShowImage = true;
		Btn_FormatPainterUI.Click += ShowFormatHelperUI;
		Btn_SuperscriptAndSubscript.Image = Resources.SquareSuperscript;
		Btn_SuperscriptAndSubscript.Items.Add(Btn_SquareSuperscript);
		Btn_SuperscriptAndSubscript.Items.Add(Btn_CubeSuperscript);
		Btn_SuperscriptAndSubscript.Items.Add(Btn_NumberSuperscript);
		Btn_SuperscriptAndSubscript.Items.Add(Btn_NumberSubscript);
		Btn_SuperscriptAndSubscript.Items.Add(Btn_CustomScript);
		Btn_SuperscriptAndSubscript.Label = "平方上标";
		Btn_SuperscriptAndSubscript.Name = "Btn_SuperscriptAndSubscript";
		Btn_SuperscriptAndSubscript.Click += ScriptFormat_Click;
		Btn_SquareSuperscript.Image = Resources.SquareSuperscript;
		Btn_SquareSuperscript.Label = "平方上标";
		Btn_SquareSuperscript.Name = "Btn_SquareSuperscript";
		Btn_SquareSuperscript.ScreenTip = "平方上标";
		Btn_SquareSuperscript.ShowImage = true;
		Btn_SquareSuperscript.Click += ScriptFormat_Click;
		Btn_CubeSuperscript.Image = Resources.CubeSuperscript;
		Btn_CubeSuperscript.Label = "立方上标";
		Btn_CubeSuperscript.Name = "Btn_CubeSuperscript";
		Btn_CubeSuperscript.ScreenTip = "立方上标";
		Btn_CubeSuperscript.ShowImage = true;
		Btn_CubeSuperscript.Click += ScriptFormat_Click;
		Btn_NumberSuperscript.Image = Resources.NumberSuperscript;
		Btn_NumberSuperscript.Label = "数字上标";
		Btn_NumberSuperscript.Name = "Btn_NumberSuperscript";
		Btn_NumberSuperscript.ScreenTip = "数字上标";
		Btn_NumberSuperscript.ShowImage = true;
		Btn_NumberSuperscript.Click += ScriptFormat_Click;
		Btn_NumberSubscript.Image = Resources.NumberSubscript;
		Btn_NumberSubscript.Label = "数字下标";
		Btn_NumberSubscript.Name = "Btn_NumberSubscript";
		Btn_NumberSubscript.ScreenTip = "数字下标";
		Btn_NumberSubscript.ShowImage = true;
		Btn_NumberSubscript.Click += ScriptFormat_Click;
		Btn_CustomScript.Image = Resources.CustomScript;
		Btn_CustomScript.Label = "自定义";
		Btn_CustomScript.Name = "Btn_CustomScript";
		Btn_CustomScript.ScreenTip = "自定义上下标";
		Btn_CustomScript.ShowImage = true;
		Btn_CustomScript.Click += ShowFormatHelperUI;
		Gp_List.DialogLauncher = dialogLauncher3;
		Gp_List.Items.Add(Btn_FastListFormat);
		Gp_List.Items.Add(Btn_ListPunctuation);
		Gp_List.Items.Add(Btn_ListStartFromOne);
		Gp_List.Items.Add(Btn_TransToList);
		Gp_List.Label = "列表";
		Gp_List.Name = "Gp_List";
		Gp_List.DialogLauncherClick += ShowFormatHelperUI;
		Btn_FastListFormat.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Btn_FastListFormat.Image = Resources.ListFormat;
		Btn_FastListFormat.Label = "快速对齐列表";
		Btn_FastListFormat.Name = "Btn_FastListFormat";
		Btn_FastListFormat.ScreenTip = "快速设置列表";
		Btn_FastListFormat.ShowImage = true;
		Btn_FastListFormat.Click += Btn_FastListFormat_Click;
		Btn_ListPunctuation.Image = Resources.ListPunctuation;
		Btn_ListPunctuation.Items.Add(Btn_ListSemicolonPeriod);
		Btn_ListPunctuation.Items.Add(Btn_ListCommaPeriod);
		Btn_ListPunctuation.Items.Add(Btn_ListPeriod);
		Btn_ListPunctuation.Items.Add(Btn_ListNoPunctuation);
		Btn_ListPunctuation.Label = "分号句号";
		Btn_ListPunctuation.Name = "Btn_ListPunctuation";
		Btn_ListPunctuation.ScreenTip = "列表标点符号";
		Btn_ListPunctuation.Click += ListPunctuation_ButtonClick;
		Btn_ListSemicolonPeriod.Image = Resources.ListPunctuation;
		Btn_ListSemicolonPeriod.Label = "分号句号";
		Btn_ListSemicolonPeriod.Name = "Btn_ListSemicolonPeriod";
		Btn_ListSemicolonPeriod.ShowImage = true;
		Btn_ListSemicolonPeriod.Click += ListPunctuation_ButtonClick;
		Btn_ListCommaPeriod.Image = Resources.ListPunctuation_CommaPeriod;
		Btn_ListCommaPeriod.Label = "逗号句号";
		Btn_ListCommaPeriod.Name = "Btn_ListCommaPeriod";
		Btn_ListCommaPeriod.ShowImage = true;
		Btn_ListCommaPeriod.Click += ListPunctuation_ButtonClick;
		Btn_ListPeriod.Image = Resources.ListPunctuation_AllPeriod;
		Btn_ListPeriod.Label = "全为句号";
		Btn_ListPeriod.Name = "Btn_ListPeriod";
		Btn_ListPeriod.ShowImage = true;
		Btn_ListPeriod.Click += ListPunctuation_ButtonClick;
		Btn_ListNoPunctuation.Image = Resources.ListPunctuation_Remove;
		Btn_ListNoPunctuation.Label = "删除标点";
		Btn_ListNoPunctuation.Name = "Btn_ListNoPunctuation";
		Btn_ListNoPunctuation.ShowImage = true;
		Btn_ListNoPunctuation.Click += ListPunctuation_ButtonClick;
		Btn_ListStartFromOne.Image = Resources.ListStartFromOne;
		Btn_ListStartFromOne.Label = "断开重编";
		Btn_ListStartFromOne.Name = "Btn_ListStartFromOne";
		Btn_ListStartFromOne.ScreenTip = "断开列表重新编号";
		Btn_ListStartFromOne.ShowImage = true;
		Btn_ListStartFromOne.Click += Btn_ListStartFromOne_Click;
		Btn_TransToList.Image = Resources.NormalList;
		Btn_TransToList.Label = "普通列表";
		Btn_TransToList.Name = "Btn_TransToList";
		Btn_TransToList.ScreenTip = "普通列表";
		Btn_TransToList.ShowImage = true;
		Btn_TransToList.Click += Btn_TransToList_Click;
		Gp_LevelList.DialogLauncher = dialogLauncher4;
		Gp_LevelList.Items.Add(Ga_FastLevelList);
		Gp_LevelList.Label = "多级列表";
		Gp_LevelList.Name = "Gp_LevelList";
		Gp_LevelList.DialogLauncherClick += ShowFormatHelperUI;
		Ga_FastLevelList.Buttons.Add(Btn_Set2LevelList);
		Ga_FastLevelList.Buttons.Add(Btn_Set3LevelList);
		Ga_FastLevelList.Buttons.Add(Btn_Set4LevelList);
		Ga_FastLevelList.Buttons.Add(Btn_Set5LevelList);
		Ga_FastLevelList.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Ga_FastLevelList.Image = Resources.LevelListFormat;
		Ga_FastLevelList.Label = "快速多级列表";
		Ga_FastLevelList.Name = "Ga_FastLevelList";
		Ga_FastLevelList.ScreenTip = "快速创建多级列表";
		Ga_FastLevelList.ShowImage = true;
		Ga_FastLevelList.ButtonClick += SetLevelList_Click;
		Btn_Set2LevelList.Image = Resources.LevelStyle_Level2;
		Btn_Set2LevelList.Label = "设置2级多级列表";
		Btn_Set2LevelList.Name = "Btn_Set2LevelList";
		Btn_Set2LevelList.OfficeImageId = "_2";
		Btn_Set2LevelList.ShowImage = true;
		Btn_Set3LevelList.Image = Resources.LevelStyle_Level3;
		Btn_Set3LevelList.Label = "设置3级多级列表";
		Btn_Set3LevelList.Name = "Btn_Set3LevelList";
		Btn_Set3LevelList.OfficeImageId = "_3";
		Btn_Set3LevelList.ShowImage = true;
		Btn_Set4LevelList.Image = Resources.LevelStyle_Level4;
		Btn_Set4LevelList.Label = "设置4级多级列表";
		Btn_Set4LevelList.Name = "Btn_Set4LevelList";
		Btn_Set4LevelList.OfficeImageId = "_4";
		Btn_Set4LevelList.ShowImage = true;
		Btn_Set5LevelList.Image = Resources.LevelStyle_Level5;
		Btn_Set5LevelList.Label = "设置5级多级列表";
		Btn_Set5LevelList.Name = "Btn_Set5LevelList";
		Btn_Set5LevelList.OfficeImageId = "_5";
		Btn_Set5LevelList.ShowImage = true;
		Gp_TOC.Items.Add(Btn_TOC);
		Gp_TOC.Label = "目录";
		Gp_TOC.Name = "Gp_TOC";
		Btn_TOC.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Btn_TOC.Image = Resources.TableOfContents;
		Btn_TOC.Label = "插入自动目录";
		Btn_TOC.Name = "Btn_TOC";
		Btn_TOC.OfficeImageId = "TableOfContentsDialog";
		Btn_TOC.ScreenTip = "插入自动目录";
		Btn_TOC.ShowImage = true;
		Btn_TOC.Click += ShowFormatHelperUI;
		Gp_TablePictureSet.DialogLauncher = dialogLauncher5;
		Gp_TablePictureSet.Items.Add(Tog_ApplyToAll);
		Gp_TablePictureSet.Items.Add(separator2);
		Gp_TablePictureSet.Items.Add(Btn_TableLeft);
		Gp_TablePictureSet.Items.Add(Btn_TableCenter);
		Gp_TablePictureSet.Items.Add(Btn_TableRight);
		Gp_TablePictureSet.Items.Add(Btn_FirstRowBold);
		Gp_TablePictureSet.Items.Add(Btn_FirstColumnBold);
		Gp_TablePictureSet.Items.Add(Btn_TableThickOutside);
		Gp_TablePictureSet.Items.Add(Btn_TableSingleSpace);
		Gp_TablePictureSet.Items.Add(Btn_RemoveUselessLine);
		Gp_TablePictureSet.Items.Add(Btn_RemoveLeftIndent);
		Gp_TablePictureSet.Items.Add(Btn_TableFullWidth);
		Gp_TablePictureSet.Items.Add(Btn_RowHeightFitText);
		Gp_TablePictureSet.Items.Add(Btn_RepeatTitle);
		Gp_TablePictureSet.Items.Add(GA_NewTables);
		Gp_TablePictureSet.Items.Add(separator1);
		Gp_TablePictureSet.Items.Add(Btn_PictureLeft);
		Gp_TablePictureSet.Items.Add(Btn_PictureCenter);
		Gp_TablePictureSet.Items.Add(Btn_PictureRight);
		Gp_TablePictureSet.Items.Add(Btn_SetPictureWidth);
		Gp_TablePictureSet.Items.Add(Btn_SetPictureHeight);
		Gp_TablePictureSet.Items.Add(Btn_PictureSingleSpace);
		Gp_TablePictureSet.Items.Add(Ebox_PictureWidth);
		Gp_TablePictureSet.Items.Add(Ebox_PictureHeight);
		Gp_TablePictureSet.Label = "表格与图片";
		Gp_TablePictureSet.Name = "Gp_TablePictureSet";
		Gp_TablePictureSet.DialogLauncherClick += ShowFormatHelperUI;
		Tog_ApplyToAll.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Tog_ApplyToAll.Image = Resources.Mode_Document_Off;
		Tog_ApplyToAll.Label = "全文模式";
		Tog_ApplyToAll.Name = "Tog_ApplyToAll";
		Tog_ApplyToAll.ScreenTip = "全文模式";
		Tog_ApplyToAll.ShowImage = true;
		Tog_ApplyToAll.Click += Tog_ApplyToAll_Click;
		separator2.Name = "separator2";
		Btn_TableLeft.Image = Resources.TableAlignment_Left;
		Btn_TableLeft.Label = "表左对齐";
		Btn_TableLeft.Name = "Btn_TableLeft";
		Btn_TableLeft.ScreenTip = "表格左对齐";
		Btn_TableLeft.ShowImage = true;
		Btn_TableLeft.Click += SetTableAlignment_Click;
		Btn_TableCenter.Image = Resources.TableAlignment_Middle;
		Btn_TableCenter.Label = "表中对齐";
		Btn_TableCenter.Name = "Btn_TableCenter";
		Btn_TableCenter.ScreenTip = "表格居中对齐";
		Btn_TableCenter.ShowImage = true;
		Btn_TableCenter.Click += SetTableAlignment_Click;
		Btn_TableRight.Image = Resources.TableAlignment_Right;
		Btn_TableRight.Label = "表右对齐";
		Btn_TableRight.Name = "Btn_TableRight";
		Btn_TableRight.ScreenTip = "表格右对齐";
		Btn_TableRight.ShowImage = true;
		Btn_TableRight.Click += SetTableAlignment_Click;
		Btn_FirstRowBold.Image = Resources.TableFormat_FirstRowBold;
		Btn_FirstRowBold.Label = "首行加粗";
		Btn_FirstRowBold.Name = "Btn_FirstRowBold";
		Btn_FirstRowBold.OfficeImageId = "Bold";
		Btn_FirstRowBold.ScreenTip = "表格首行加粗";
		Btn_FirstRowBold.ShowImage = true;
		Btn_FirstRowBold.Click += SetTableFormat_Click;
		Btn_FirstColumnBold.Image = Resources.TableFormat_FirstColBold;
		Btn_FirstColumnBold.Label = "首列加粗";
		Btn_FirstColumnBold.Name = "Btn_FirstColumnBold";
		Btn_FirstColumnBold.OfficeImageId = "Bold";
		Btn_FirstColumnBold.ScreenTip = "表格首列加粗";
		Btn_FirstColumnBold.ShowImage = true;
		Btn_FirstColumnBold.Click += SetTableFormat_Click;
		Btn_TableThickOutside.Image = Resources.TableFormat_OutSideBorder;
		Btn_TableThickOutside.Label = "表格外框";
		Btn_TableThickOutside.Name = "Btn_TableThickOutside";
		Btn_TableThickOutside.OfficeImageId = "BorderThickOutside";
		Btn_TableThickOutside.ScreenTip = "表格外框加粗";
		Btn_TableThickOutside.ShowImage = true;
		Btn_TableThickOutside.Click += SetTableFormat_Click;
		Btn_TableSingleSpace.Image = Resources.SingleSpace;
		Btn_TableSingleSpace.Label = "表单倍行距";
		Btn_TableSingleSpace.Name = "Btn_TableSingleSpace";
		Btn_TableSingleSpace.ScreenTip = "单元格单倍行距";
		Btn_TableSingleSpace.ShowImage = true;
		Btn_TableSingleSpace.Click += SetTableFormatA_Click;
		Btn_RemoveUselessLine.Image = Resources.Remove;
		Btn_RemoveUselessLine.Label = "删除空行";
		Btn_RemoveUselessLine.Name = "Btn_RemoveUselessLine";
		Btn_RemoveUselessLine.OfficeImageId = "Delete";
		Btn_RemoveUselessLine.ScreenTip = "删除单元格空行";
		Btn_RemoveUselessLine.ShowImage = true;
		Btn_RemoveUselessLine.Click += SetTableFormat_Click;
		Btn_RemoveLeftIndent.Image = Resources.TableFormat_RemoveIndent;
		Btn_RemoveLeftIndent.Label = "移除左缩进";
		Btn_RemoveLeftIndent.Name = "Btn_RemoveLeftIndent";
		Btn_RemoveLeftIndent.OfficeImageId = "OutdentClassic";
		Btn_RemoveLeftIndent.ScreenTip = "移除单元格左缩进";
		Btn_RemoveLeftIndent.ShowImage = true;
		Btn_RemoveLeftIndent.Click += SetTableFormatA_Click;
		Btn_TableFullWidth.Image = Resources.TableFullWidth;
		Btn_TableFullWidth.Label = "表格满宽";
		Btn_TableFullWidth.Name = "Btn_TableFullWidth";
		Btn_TableFullWidth.ScreenTip = "表格满宽";
		Btn_TableFullWidth.ShowImage = true;
		Btn_TableFullWidth.Click += Btn_TableFullWidth_Click;
		Btn_RowHeightFitText.Image = Resources.TableRowHeightFit;
		Btn_RowHeightFitText.Label = "行高适配";
		Btn_RowHeightFitText.Name = "Btn_RowHeightFitText";
		Btn_RowHeightFitText.ScreenTip = "表格行高适配内容";
		Btn_RowHeightFitText.ShowImage = true;
		Btn_RowHeightFitText.Click += ShowFormatHelperUI;
		Btn_RepeatTitle.Image = Resources.RepeatTableTitleRow;
		Btn_RepeatTitle.Label = "重复标题";
		Btn_RepeatTitle.Name = "Btn_RepeatTitle";
		Btn_RepeatTitle.ScreenTip = "重复标题行";
		Btn_RepeatTitle.ShowImage = true;
		Btn_RepeatTitle.Click += Btn_RepeatTitle_Click;
		GA_NewTables.Buttons.Add(Btn_TableStyle);
		GA_NewTables.ColumnCount = 1;
		GA_NewTables.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		GA_NewTables.Image = Resources.CreateNewTable;
		GA_NewTables.ItemImageSize = new Size(376, 100);
		ribbonDropDownItem25.Image = Resources.NewTable_01;
		ribbonDropDownItem25.Label = "样式1";
		ribbonDropDownItem26.Image = Resources.NewTable_02;
		ribbonDropDownItem26.Label = "样式2";
		ribbonDropDownItem27.Image = Resources.NewTable_03;
		ribbonDropDownItem27.Label = "样式3";
		ribbonDropDownItem28.Image = Resources.NewTable_04;
		ribbonDropDownItem28.Label = "样式4";
		ribbonDropDownItem29.Image = Resources.NewTable_05;
		ribbonDropDownItem29.Label = "样式5";
		ribbonDropDownItem30.Image = Resources.NewTable_06;
		ribbonDropDownItem30.Label = "样式6";
		ribbonDropDownItem31.Image = Resources.NewTable_07;
		ribbonDropDownItem31.Label = "样式7";
		ribbonDropDownItem32.Image = Resources.NewTable_08;
		ribbonDropDownItem32.Label = "样式8";
		ribbonDropDownItem33.Image = Resources.NewTable_09;
		ribbonDropDownItem33.Label = "样式9";
		ribbonDropDownItem34.Image = Resources.NewTable_10;
		ribbonDropDownItem34.Label = "样式10";
		GA_NewTables.Items.Add(ribbonDropDownItem25);
		GA_NewTables.Items.Add(ribbonDropDownItem26);
		GA_NewTables.Items.Add(ribbonDropDownItem27);
		GA_NewTables.Items.Add(ribbonDropDownItem28);
		GA_NewTables.Items.Add(ribbonDropDownItem29);
		GA_NewTables.Items.Add(ribbonDropDownItem30);
		GA_NewTables.Items.Add(ribbonDropDownItem31);
		GA_NewTables.Items.Add(ribbonDropDownItem32);
		GA_NewTables.Items.Add(ribbonDropDownItem33);
		GA_NewTables.Items.Add(ribbonDropDownItem34);
		GA_NewTables.Label = "新建表格";
		GA_NewTables.Name = "GA_NewTables";
		GA_NewTables.OfficeImageId = "TableStyleNew";
		GA_NewTables.ScreenTip = "新建表格";
		GA_NewTables.ShowImage = true;
		GA_NewTables.ShowItemLabel = false;
		GA_NewTables.ButtonClick += ShowFormatHelperUI;
		GA_NewTables.Click += GA_NewTables_Click;
		Btn_TableStyle.Image = Resources.TableStyle;
		Btn_TableStyle.Label = "表格样式设置";
		Btn_TableStyle.Name = "Btn_TableStyle";
		Btn_TableStyle.OfficeImageId = "TableStyleNew";
		Btn_TableStyle.ScreenTip = "样式设置";
		Btn_TableStyle.ShowImage = true;
		separator1.Name = "separator1";
		Btn_PictureLeft.Image = Resources.PictureAlignment_Left;
		Btn_PictureLeft.Label = "图左对齐";
		Btn_PictureLeft.Name = "Btn_PictureLeft";
		Btn_PictureLeft.ScreenTip = "图片左对齐";
		Btn_PictureLeft.ShowImage = true;
		Btn_PictureLeft.Click += SetPictureAlignment_Click;
		Btn_PictureCenter.Image = Resources.PictureAlignment_Middle;
		Btn_PictureCenter.Label = "图中对齐";
		Btn_PictureCenter.Name = "Btn_PictureCenter";
		Btn_PictureCenter.ScreenTip = "图片居中对齐";
		Btn_PictureCenter.ShowImage = true;
		Btn_PictureCenter.Click += SetPictureAlignment_Click;
		Btn_PictureRight.Image = Resources.PictureAlignment_Right;
		Btn_PictureRight.Label = "图右对齐";
		Btn_PictureRight.Name = "Btn_PictureRight";
		Btn_PictureRight.ScreenTip = "图片右对齐";
		Btn_PictureRight.ShowImage = true;
		Btn_PictureRight.Click += SetPictureAlignment_Click;
		Btn_SetPictureWidth.Image = Resources.PictureWidth;
		Btn_SetPictureWidth.Label = "设置图宽";
		Btn_SetPictureWidth.Name = "Btn_SetPictureWidth";
		Btn_SetPictureWidth.OfficeImageId = "SizeToControlWidth";
		Btn_SetPictureWidth.ScreenTip = "设置图片宽度";
		Btn_SetPictureWidth.ShowImage = true;
		Btn_SetPictureWidth.Click += SetPictureFormat_Click;
		Btn_SetPictureHeight.Image = Resources.PictureHeight;
		Btn_SetPictureHeight.Label = "设置图高";
		Btn_SetPictureHeight.Name = "Btn_SetPictureHeight";
		Btn_SetPictureHeight.OfficeImageId = "SizeToControlHeight";
		Btn_SetPictureHeight.ScreenTip = "设置图片高度";
		Btn_SetPictureHeight.ShowImage = true;
		Btn_SetPictureHeight.Click += SetPictureFormat_Click;
		Btn_PictureSingleSpace.Image = Resources.SingleSpace;
		Btn_PictureSingleSpace.Label = "图单倍行距";
		Btn_PictureSingleSpace.Name = "Btn_PictureSingleSpace";
		Btn_PictureSingleSpace.ScreenTip = "图片单倍行距";
		Btn_PictureSingleSpace.ShowImage = true;
		Btn_PictureSingleSpace.Click += SetPictureFormat_Click;
		Ebox_PictureWidth.Label = "图宽";
		Ebox_PictureWidth.MaxLength = 10;
		Ebox_PictureWidth.Name = "Ebox_PictureWidth";
		Ebox_PictureWidth.ScreenTip = "图片宽度";
		Ebox_PictureWidth.Text = null;
		Ebox_PictureWidth.TextChanged += NumberValidate;
		Ebox_PictureHeight.Label = "图高";
		Ebox_PictureHeight.Name = "Ebox_PictureHeight";
		Ebox_PictureHeight.ScreenTip = "图片高度";
		Ebox_PictureHeight.Text = null;
		Ebox_PictureHeight.TextChanged += NumberValidate;
		Gp_DocumentSet.Items.Add(Btn_FastSetStyle);
		Gp_DocumentSet.Label = "样式";
		Gp_DocumentSet.Name = "Gp_DocumentSet";
		Btn_FastSetStyle.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Btn_FastSetStyle.Image = Resources.Styles;
		Btn_FastSetStyle.Label = "快速设置文档样式";
		Btn_FastSetStyle.Name = "Btn_FastSetStyle";
		Btn_FastSetStyle.OfficeImageId = "StylesModifyStyle";
		Btn_FastSetStyle.ScreenTip = "设置文档样式";
		Btn_FastSetStyle.ShowImage = true;
		Btn_FastSetStyle.Click += ShowFormatHelperUI;
		Gp_QRCoder.DialogLauncher = dialogLauncher6;
		Gp_QRCoder.Items.Add(Btn_FastQRCoder);
		Gp_QRCoder.Label = "二维码/条码";
		Gp_QRCoder.Name = "Gp_QRCoder";
		Gp_QRCoder.DialogLauncherClick += ShowFormatHelperUI;
		Btn_FastQRCoder.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Btn_FastQRCoder.Image = Resources.QRCode;
		Btn_FastQRCoder.Label = "创建二维码";
		Btn_FastQRCoder.Name = "Btn_FastQRCoder";
		Btn_FastQRCoder.ScreenTip = "创建二维码";
		Btn_FastQRCoder.ShowImage = true;
		Btn_FastQRCoder.Click += Btn_FastQRCoder_Click;
		Gp_Help.Items.Add(Ga_Utilities);
		Gp_Help.Items.Add(Ga_SetUnits);
		Gp_Help.Items.Add(Btn_DefaultValue);
		Gp_Help.Items.Add(GA_AboutAndHelp);
		Gp_Help.Label = "设置及说明";
		Gp_Help.Name = "Gp_Help";
		Ga_Utilities.Buttons.Add(Btn_ExportPDF);
		Ga_Utilities.Buttons.Add(Btn_OCR);
		Ga_Utilities.Buttons.Add(Btn_ExportPng);
		Ga_Utilities.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Ga_Utilities.Image = Resources.Utilities;
		Ga_Utilities.Label = "工具箱";
		Ga_Utilities.Name = "Ga_Utilities";
		Ga_Utilities.ScreenTip = "工具箱";
		Ga_Utilities.ShowImage = true;
		Ga_Utilities.ButtonClick += Ga_Utilities_ButtonClick;
		Btn_ExportPDF.Image = Resources.ExportPDF;
		Btn_ExportPDF.Label = "批量转存PDF";
		Btn_ExportPDF.Name = "Btn_ExportPDF";
		Btn_ExportPDF.ScreenTip = "批量转存PDF";
		Btn_ExportPDF.ShowImage = true;
		Btn_OCR.Image = Resources.OCRTools;
		Btn_OCR.Label = "文字识别工具";
		Btn_OCR.Name = "Btn_OCR";
		Btn_OCR.ScreenTip = "文字识别工具";
		Btn_OCR.ShowImage = true;
		Btn_ExportPng.Image = Resources.ExportToImage;
		Btn_ExportPng.Label = "导出图片";
		Btn_ExportPng.Name = "Btn_ExportPng";
		Btn_ExportPng.ScreenTip = "导出图片";
		Btn_ExportPng.ShowImage = true;
		Ga_SetUnits.Buttons.Add(Btn_UnitInch);
		Ga_SetUnits.Buttons.Add(Btn_UnitCM);
		Ga_SetUnits.Buttons.Add(Btn_UnitMM);
		Ga_SetUnits.Buttons.Add(Btn_UnitPt);
		Ga_SetUnits.Buttons.Add(Btn_UnitPicas);
		Ga_SetUnits.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Ga_SetUnits.Image = Resources.MeasureUnits;
		Ga_SetUnits.Label = "测量单位";
		Ga_SetUnits.Name = "Ga_SetUnits";
		Ga_SetUnits.OfficeImageId = "ShowRuler";
		Ga_SetUnits.ScreenTip = "测量单位";
		Ga_SetUnits.ShowImage = true;
		Ga_SetUnits.ButtonClick += Ga_SetUnits_ButtonClick;
		Btn_UnitInch.Label = "英寸(Inches)";
		Btn_UnitInch.Name = "Btn_UnitInch";
		Btn_UnitCM.Label = "厘米(cm)";
		Btn_UnitCM.Name = "Btn_UnitCM";
		Btn_UnitMM.Label = "毫米(mm)";
		Btn_UnitMM.Name = "Btn_UnitMM";
		Btn_UnitPt.Label = "磅(pt)";
		Btn_UnitPt.Name = "Btn_UnitPt";
		Btn_UnitPicas.Label = "派卡(Picas)";
		Btn_UnitPicas.Name = "Btn_UnitPicas";
		Btn_DefaultValue.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Btn_DefaultValue.Image = Resources.Settings;
		Btn_DefaultValue.Label = "设置";
		Btn_DefaultValue.Name = "Btn_DefaultValue";
		Btn_DefaultValue.ScreenTip = "设置";
		Btn_DefaultValue.ShowImage = true;
		Btn_DefaultValue.Click += Btn_DefaultValue_Click;
		GA_AboutAndHelp.Buttons.Add(Btn_HelpOnline);
		GA_AboutAndHelp.Buttons.Add(Btn_AboutUs);
		GA_AboutAndHelp.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		GA_AboutAndHelp.Image = Resources.Help;
		GA_AboutAndHelp.Label = "使用说明";
		GA_AboutAndHelp.Name = "GA_AboutAndHelp";
		GA_AboutAndHelp.OfficeImageId = "Help";
		GA_AboutAndHelp.ScreenTip = "使用说明";
		GA_AboutAndHelp.ShowImage = true;
		GA_AboutAndHelp.ButtonClick += GA_AboutAndHelp_ButtonClick;
		Btn_HelpOnline.Image = Resources.HelpDocOnline;
		Btn_HelpOnline.Label = "在线帮助文档";
		Btn_HelpOnline.Name = "Btn_HelpOnline";
		Btn_HelpOnline.OfficeImageId = "CLViewDialogHelpID";
		Btn_HelpOnline.ShowImage = true;
		Btn_AboutUs.Image = Resources.Infomation;
		Btn_AboutUs.Label = "关于与打赏";
		Btn_AboutUs.Name = "Btn_AboutUs";
		Btn_AboutUs.OfficeImageId = "Info";
		Btn_AboutUs.ShowImage = true;
		base.Name = "WordFormatHelperRibbon";
		base.RibbonType = "Microsoft.Word.Document";
		base.Tabs.Add(TabAddIns);
		base.Tabs.Add(Tab_WordFormatAssistant);
		base.Load += WordFormatHelperRibbon_Load;
		TabAddIns.ResumeLayout(performLayout: false);
		TabAddIns.PerformLayout();
		Tab_WordFormatAssistant.ResumeLayout(performLayout: false);
		Tab_WordFormatAssistant.PerformLayout();
		Gp_PageSet.ResumeLayout(performLayout: false);
		Gp_PageSet.PerformLayout();
		Gp_TextFormatSet.ResumeLayout(performLayout: false);
		Gp_TextFormatSet.PerformLayout();
		Gp_List.ResumeLayout(performLayout: false);
		Gp_List.PerformLayout();
		Gp_LevelList.ResumeLayout(performLayout: false);
		Gp_LevelList.PerformLayout();
		Gp_TOC.ResumeLayout(performLayout: false);
		Gp_TOC.PerformLayout();
		Gp_TablePictureSet.ResumeLayout(performLayout: false);
		Gp_TablePictureSet.PerformLayout();
		Gp_DocumentSet.ResumeLayout(performLayout: false);
		Gp_DocumentSet.PerformLayout();
		Gp_QRCoder.ResumeLayout(performLayout: false);
		Gp_QRCoder.PerformLayout();
		Gp_Help.ResumeLayout(performLayout: false);
		Gp_Help.PerformLayout();
		ResumeLayout(performLayout: false);
	}
}
