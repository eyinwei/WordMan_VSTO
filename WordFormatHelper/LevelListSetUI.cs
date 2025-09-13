using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordFormatHelper{

public class LevelListSetUI : UserControl
{
	private bool userChanged = true;

	private IContainer components;

	private ComboBox Cmb_LstTLevels;

	private Label label6;

	private Label label5;

	private Label label3;

	private NumericUpDownWithUnit UnifiedIndent;

	private Button BtnApplyFastSet;

	private CheckBox ChkNumber;

	private CheckBox ChkText;

	private NumericUpDownWithUnit AfterFristIndent;

	private Label label4;

	private NumericUpDownWithUnit LevelFristIndent;

	private CheckBox ChkAfterNumber;

	private CheckBox ChkAutoIndent;

	private Label label2;

	private Label Lab_LinkedStyle;

	private Button GetCurrentListTemplate;

	private Button CreateListTemplate;

	private Label Lab_AfterNumIndent;

	private Label Lab_TextIndent;

	private Label Lab_NumIndent;

	private Label Lab_NumberFormat;

	private Label Lab_NumberStyle;

	private Label label1;

	private GroupBox groupBox1;

	private CheckBox Chk_LinkToTitles;

	private CheckBox Chk_UnlinkToTitles;

	protected override CreateParams CreateParams
	{
		get
		{
			CreateParams obj = base.CreateParams;
			obj.ExStyle |= 33554432;
			return obj;
		}
	}

	public LevelListSetUI()
	{
		InitializeComponent();
		Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
		Section first = application.Selection.Sections.First;
		float num = application.PointsToCentimeters(first.PageSetup.PageWidth);
		application.PointsToCentimeters(first.PageSetup.PageHeight);
		for (int i = 1; i <= 9; i++)
		{
			string name = "LabLevel" + i;
			string name2 = "CmbNumStyle" + i;
			string name3 = "TextBoxNumFormat" + i;
			string name4 = "TxtBoxNumIndent" + i;
			string name5 = "TxtBoxTextIndent" + i;
			string name6 = "TxtBoxAfterNumIndent" + i;
			string name7 = "CmbLinkedStyle" + i;
			Label label = new Label
			{
				Text = "第" + i + "级"
			};
			ComboBox comboBox = new ComboBox();
			TextBox textBox = new TextBox();
			NumericUpDownWithUnit numericUpDownWithUnit = new NumericUpDownWithUnit();
			NumericUpDownWithUnit numericUpDownWithUnit2 = new NumericUpDownWithUnit();
			NumericUpDownWithUnit numericUpDownWithUnit3 = new NumericUpDownWithUnit();
			ComboBox comboBox2 = new ComboBox();
			label.Name = name;
			label.AutoSize = false;
			label.Size = new Size(50, 30);
			label.Location = new Point(10, 40 + (i - 1) * 35);
			label.TextAlign = ContentAlignment.MiddleCenter;
			label.Visible = false;
			base.Controls.Add(label);
			comboBox.Name = name2;
			comboBox.AutoSize = false;
			comboBox.Items.AddRange(new object[10] { "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,貳,叁...", "甲,乙,丙...", "正规编号" });
			comboBox.Size = new Size(100, 30);
			comboBox.Location = new Point(70, 40 + (i - 1) * 35);
			comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
			comboBox.Visible = false;
			base.Controls.Add(comboBox);
			comboBox.TabIndex = 1 + (i - 1) * 6;
			textBox.Name = name3;
			textBox.AutoSize = false;
			textBox.Size = new Size(100, 30);
			textBox.Location = new Point(180, 40 + (i - 1) * 35);
			textBox.Visible = false;
			base.Controls.Add(textBox);
			textBox.TabIndex = 2 + (i - 1) * 6;
			numericUpDownWithUnit.Name = name4;
			numericUpDownWithUnit.AutoSize = false;
			numericUpDownWithUnit.Label = "厘米";
			numericUpDownWithUnit.Increment = 0.01m;
			numericUpDownWithUnit.DecimalPlaces = 2;
			numericUpDownWithUnit.Maximum = (decimal)num;
			numericUpDownWithUnit.Size = new Size(100, 30);
			numericUpDownWithUnit.Location = new Point(290, 40 + (i - 1) * 35);
			numericUpDownWithUnit.Visible = false;
			base.Controls.Add(numericUpDownWithUnit);
			numericUpDownWithUnit.TabIndex = 3 + (i - 1) * 6;
			numericUpDownWithUnit2.Name = name5;
			numericUpDownWithUnit2.AutoSize = false;
			numericUpDownWithUnit2.Label = "厘米";
			numericUpDownWithUnit2.Increment = 0.01m;
			numericUpDownWithUnit2.DecimalPlaces = 2;
			numericUpDownWithUnit2.Maximum = (decimal)num;
			numericUpDownWithUnit2.Size = new Size(100, 30);
			numericUpDownWithUnit2.Location = new Point(400, 40 + (i - 1) * 35);
			numericUpDownWithUnit2.Visible = false;
			base.Controls.Add(numericUpDownWithUnit2);
			numericUpDownWithUnit2.TabIndex = 4 + (i - 1) * 6;
			numericUpDownWithUnit3.Name = name6;
			numericUpDownWithUnit3.AutoSize = false;
			numericUpDownWithUnit3.Label = "厘米";
			numericUpDownWithUnit3.Increment = 0.01m;
			numericUpDownWithUnit3.DecimalPlaces = 2;
			numericUpDownWithUnit3.Maximum = (decimal)num;
			numericUpDownWithUnit3.Size = new Size(100, 30);
			numericUpDownWithUnit3.Location = new Point(510, 40 + (i - 1) * 35);
			numericUpDownWithUnit3.Visible = false;
			base.Controls.Add(numericUpDownWithUnit3);
			numericUpDownWithUnit3.TabIndex = 5 + (i - 1) * 6;
			comboBox2.Name = name7;
			comboBox2.AutoSize = false;
			comboBox2.Size = new Size(100, 30);
			comboBox2.Location = new Point(620, 40 + (i - 1) * 35);
			comboBox2.Items.AddRange(new object[10] { "无", "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7", "标题 8", "标题 9" });
			comboBox2.Visible = false;
			base.Controls.Add(comboBox2);
			comboBox2.TabIndex = 6 + (i - 1) * 6;
			comboBox2.Leave += LinkStyle_Leaved;
		}
		UnifiedIndent.Maximum = (decimal)num;
		LevelFristIndent.Maximum = (decimal)num;
		AfterFristIndent.Maximum = (decimal)num;
	}

	private void Cmb_LstTLevels_TextChanged(object sender, EventArgs e)
	{
		if (Cmb_LstTLevels.Text != "")
		{
			ShowLevelClass(Convert.ToInt16(Cmb_LstTLevels.Text));
		}
	}

	private void ShowLevelClass(int iLevel)
	{
		string text = "";
		for (int i = 1; i <= 9; i++)
		{
			string key = "CmbNumStyle" + i;
			string key2 = "TextBoxNumFormat" + i;
			string key3 = "TxtBoxNumIndent" + i;
			string key4 = "TxtBoxTextIndent" + i;
			string key5 = "TxtBoxAfterNumIndent" + i;
			string key6 = "CmbLinkedStyle" + i;
			string key7 = "LabLevel" + i;
			Label label = base.Controls.Find(key7, searchAllChildren: false)[0] as Label;
			ComboBox comboBox = base.Controls.Find(key, searchAllChildren: false)[0] as ComboBox;
			TextBox textBox = base.Controls.Find(key2, searchAllChildren: false)[0] as TextBox;
			NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find(key3, searchAllChildren: false)[0] as NumericUpDownWithUnit;
			NumericUpDownWithUnit numericUpDownWithUnit2 = base.Controls.Find(key4, searchAllChildren: false)[0] as NumericUpDownWithUnit;
			NumericUpDownWithUnit numericUpDownWithUnit3 = base.Controls.Find(key5, searchAllChildren: false)[0] as NumericUpDownWithUnit;
			ComboBox comboBox2 = base.Controls.Find(key6, searchAllChildren: false)[0] as ComboBox;
			if (i > iLevel)
			{
				label.Visible = false;
				comboBox.Visible = false;
				textBox.Visible = false;
				numericUpDownWithUnit.Visible = false;
				numericUpDownWithUnit2.Visible = false;
				numericUpDownWithUnit3.Visible = false;
				comboBox2.Visible = false;
				continue;
			}
			label.Visible = true;
			comboBox.Visible = true;
			comboBox.SelectedIndex = ((comboBox.SelectedIndex != -1) ? comboBox.SelectedIndex : 0);
			textBox.Visible = true;
			text = text + "%" + i + ".";
			if (textBox.Text == "")
			{
				textBox.Text = text.Substring(0, text.Length - 1);
			}
			numericUpDownWithUnit.Visible = true;
			numericUpDownWithUnit2.Visible = true;
			numericUpDownWithUnit3.Visible = true;
			comboBox2.Visible = true;
			comboBox2.SelectedIndex = ((comboBox2.SelectedIndex != -1) ? comboBox2.SelectedIndex : 0);
		}
		ChkAfterNumber.Enabled = true;
		Chk_LinkToTitles.Enabled = true;
		Chk_UnlinkToTitles.Enabled = true;
		ChkAutoIndent.Enabled = true;
		ChkNumber.Enabled = true;
		ChkText.Enabled = true;
		BtnApplyFastSet.Enabled = true;
	}

	private void GetCurrentListTemplate_Click(object sender, EventArgs e)
	{
		int num = 1;
		ListTemplate listTemplate = Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListTemplate;
		if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType != WdListType.wdListOutlineNumbering)
		{
			MessageBox.Show("当前位置无多级列表或列表不属于大纲级别！", "提醒");
			return;
		}
		foreach (ListLevel listLevel3 in listTemplate.ListLevels)
		{
			if (listLevel3.NumberFormat == "")
			{
				break;
			}
			num = listLevel3.Index;
		}
		Cmb_LstTLevels.Text = num.ToString();
		for (int i = 1; i <= num; i++)
		{
			ListLevel listLevel2 = listTemplate.ListLevels[i];
			(base.Controls.Find("CmbNumStyle" + i, searchAllChildren: false)[0] as ComboBox).SelectedIndex = Globals.ThisAddIn.LevelNumStyle.IndexOf(listLevel2.NumberStyle);
			(base.Controls.Find("TextBoxNumFormat" + i, searchAllChildren: false)[0] as TextBox).Text = listLevel2.NumberFormat.ToString();
			(base.Controls.Find("TxtBoxNumIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit).Value = (decimal)Globals.ThisAddIn.Application.PointsToCentimeters(listLevel2.NumberPosition);
			(base.Controls.Find("TxtBoxTextIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit).Value = (decimal)Globals.ThisAddIn.Application.PointsToCentimeters(listLevel2.TextPosition);
			NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find("TxtBoxAfterNumIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit;
			if (listLevel2.TabPosition != 9999999f)
			{
				numericUpDownWithUnit.Value = (decimal)Globals.ThisAddIn.Application.PointsToCentimeters(listLevel2.TabPosition);
			}
			else
			{
				numericUpDownWithUnit.Value = 0m;
			}
			(base.Controls.Find("CmbLinkedStyle" + i, searchAllChildren: false)[0] as ComboBox).Text = ((listLevel2.LinkedStyle == "") ? "无" : listLevel2.LinkedStyle);
		}
	}

	private void CreateListTemplate_Click(object sender, EventArgs e)
	{
		int num = Convert.ToInt16(Cmb_LstTLevels.Text);
		int[] array = new int[num];
		string[] array2 = new string[num];
		string[] array3 = new string[num];
		float[] array4 = new float[num];
		float[] array5 = new float[num];
		float[] array6 = new float[num];
		for (int i = 0; i < num; i++)
		{
			ComboBox comboBox = base.Controls.Find("CmbNumStyle" + (i + 1), searchAllChildren: false)[0] as ComboBox;
			array[i] = comboBox.SelectedIndex;
			TextBox textBox = base.Controls.Find("TextBoxNumFormat" + (i + 1), searchAllChildren: false)[0] as TextBox;
			if (!textBox.Text.Contains("%" + (i + 1)))
			{
				MessageBox.Show("错误：第" + (i + 1) + "级编号格式未包含本级别的编号！");
				return;
			}
			array2[i] = textBox.Text;
			ComboBox comboBox2 = base.Controls.Find("CmbLinkedStyle" + (i + 1), searchAllChildren: false)[0] as ComboBox;
			if (i == 0)
			{
				array3[i] = comboBox2.Text;
			}
			else
			{
				if (array3.Contains(comboBox2.Text) && comboBox2.Text != "无")
				{
					MessageBox.Show("错误：第" + (i + 1) + "级链接样式出现重复！");
					return;
				}
				array3[i] = comboBox2.Text;
			}
			NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find("TxtBoxNumIndent" + (i + 1), searchAllChildren: false)[0] as NumericUpDownWithUnit;
			array4[i] = (float)numericUpDownWithUnit.Value;
			numericUpDownWithUnit = base.Controls.Find("TxtBoxTextIndent" + (i + 1), searchAllChildren: false)[0] as NumericUpDownWithUnit;
			array5[i] = (float)numericUpDownWithUnit.Value;
			numericUpDownWithUnit = base.Controls.Find("TxtBoxAfterNumIndent" + (i + 1), searchAllChildren: false)[0] as NumericUpDownWithUnit;
			if (numericUpDownWithUnit.Value != 0m)
			{
				array6[i] = (float)numericUpDownWithUnit.Value;
			}
		}
		Globals.ThisAddIn.CreateListTemplate(num, array, array2, array4, array5, array6, array3);
	}

	private void ChkNumber_CheckedChanged(object sender, EventArgs e)
	{
		if (ChkNumber.Checked || ChkText.Checked || ChkAfterNumber.Checked)
		{
			UnifiedIndent.Enabled = true;
		}
		else
		{
			UnifiedIndent.Enabled = false;
		}
	}

	private void ChkAutoIndent_CheckedChanged(object sender, EventArgs e)
	{
		if (ChkAutoIndent.Checked)
		{
			LevelFristIndent.Enabled = true;
			AfterFristIndent.Enabled = true;
		}
		else
		{
			LevelFristIndent.Enabled = false;
			AfterFristIndent.Enabled = false;
		}
	}

	private void BtnApplyFastSet_Click(object sender, EventArgs e)
	{
		for (int i = 1; i <= Convert.ToInt16(Cmb_LstTLevels.Text); i++)
		{
			if (ChkNumber.Checked)
			{
				NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find("TxtBoxNumIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit;
				numericUpDownWithUnit.Value = UnifiedIndent.Value;
			}
			if (ChkText.Checked)
			{
				NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find("TxtBoxTextIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit;
				numericUpDownWithUnit.Value = UnifiedIndent.Value;
			}
			if (ChkAfterNumber.Checked)
			{
				NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find("TxtBoxAfterNumIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit;
				numericUpDownWithUnit.Value = UnifiedIndent.Value;
			}
			if (ChkAutoIndent.Checked)
			{
				NumericUpDownWithUnit numericUpDownWithUnit = base.Controls.Find("TxtBoxNumIndent" + i, searchAllChildren: false)[0] as NumericUpDownWithUnit;
				if (i == 1)
				{
					numericUpDownWithUnit.Value = LevelFristIndent.Value;
				}
				else
				{
					numericUpDownWithUnit.Value = LevelFristIndent.Value + AfterFristIndent.Value * (decimal)(i - 1);
				}
			}
			if (Chk_LinkToTitles.Checked)
			{
				(base.Controls.Find("CmbLinkedStyle" + i, searchAllChildren: false)[0] as ComboBox).SelectedIndex = i;
			}
			else if (Chk_UnlinkToTitles.Checked)
			{
				(base.Controls.Find("CmbLinkedStyle" + i, searchAllChildren: false)[0] as ComboBox).SelectedIndex = 0;
			}
		}
		ChkNumber.Checked = false;
		ChkText.Checked = false;
		ChkAfterNumber.Checked = false;
		ChkAutoIndent.Checked = false;
		Chk_LinkToTitles.Checked = false;
		Chk_UnlinkToTitles.Checked = false;
	}

	private void LinkStyle_Leaved(object sender, EventArgs e)
	{
		string text = (sender as ComboBox).Text;
		if (Regex.IsMatch(text, "^[ ]{1,}$") || text == "")
		{
			(sender as ComboBox).SelectedIndex = 0;
		}
	}

	private void Chk_LinkToTitles_CheckedChanged(object sender, EventArgs e)
	{
		if (userChanged && Chk_LinkToTitles.Checked)
		{
			userChanged = false;
			if (Chk_UnlinkToTitles.Checked)
			{
				Chk_UnlinkToTitles.Checked = false;
			}
			userChanged = true;
		}
	}

	private void Chk_UnlinkToTitles_CheckedChanged(object sender, EventArgs e)
	{
		if (userChanged && Chk_UnlinkToTitles.Checked)
		{
			userChanged = false;
			if (Chk_LinkToTitles.Checked)
			{
				Chk_LinkToTitles.Checked = false;
			}
			userChanged = true;
		}
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
		this.Cmb_LstTLevels = new System.Windows.Forms.ComboBox();
		this.label6 = new System.Windows.Forms.Label();
		this.label5 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.BtnApplyFastSet = new System.Windows.Forms.Button();
		this.ChkNumber = new System.Windows.Forms.CheckBox();
		this.ChkText = new System.Windows.Forms.CheckBox();
		this.label4 = new System.Windows.Forms.Label();
		this.ChkAfterNumber = new System.Windows.Forms.CheckBox();
		this.ChkAutoIndent = new System.Windows.Forms.CheckBox();
		this.label2 = new System.Windows.Forms.Label();
		this.Lab_LinkedStyle = new System.Windows.Forms.Label();
		this.GetCurrentListTemplate = new System.Windows.Forms.Button();
		this.CreateListTemplate = new System.Windows.Forms.Button();
		this.Lab_AfterNumIndent = new System.Windows.Forms.Label();
		this.Lab_TextIndent = new System.Windows.Forms.Label();
		this.Lab_NumIndent = new System.Windows.Forms.Label();
		this.Lab_NumberFormat = new System.Windows.Forms.Label();
		this.Lab_NumberStyle = new System.Windows.Forms.Label();
		this.label1 = new System.Windows.Forms.Label();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.LevelFristIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.UnifiedIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.AfterFristIndent = new WordFormatHelper.NumericUpDownWithUnit();
		this.Chk_LinkToTitles = new System.Windows.Forms.CheckBox();
		this.Chk_UnlinkToTitles = new System.Windows.Forms.CheckBox();
		this.groupBox1.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.LevelFristIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.UnifiedIndent).BeginInit();
		((System.ComponentModel.ISupportInitialize)this.AfterFristIndent).BeginInit();
		base.SuspendLayout();
		this.Cmb_LstTLevels.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.Cmb_LstTLevels.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.Cmb_LstTLevels.FormattingEnabled = true;
		this.Cmb_LstTLevels.ImeMode = System.Windows.Forms.ImeMode.NoControl;
		this.Cmb_LstTLevels.Items.AddRange(new object[9] { "1", "2", "3", "4", "5", "6", "7", "8", "9" });
		this.Cmb_LstTLevels.Location = new System.Drawing.Point(114, 366);
		this.Cmb_LstTLevels.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
		this.Cmb_LstTLevels.MaxDropDownItems = 9;
		this.Cmb_LstTLevels.Name = "Cmb_LstTLevels";
		this.Cmb_LstTLevels.Size = new System.Drawing.Size(119, 28);
		this.Cmb_LstTLevels.TabIndex = 66;
		this.Cmb_LstTLevels.SelectedIndexChanged += new System.EventHandler(Cmb_LstTLevels_TextChanged);
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(6, 248);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(136, 20);
		this.label6.TabIndex = 11;
		this.label6.Text = "3. 快速链接标题样式";
		this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(6, 143);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(178, 20);
		this.label5.TabIndex = 10;
		this.label5.Text = "2. 快速设定递进的编号缩进";
		this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(6, 22);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(192, 20);
		this.label3.TabIndex = 9;
		this.label3.Text = "1. 快速为每级设定统一的缩进";
		this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.BtnApplyFastSet.Enabled = false;
		this.BtnApplyFastSet.Location = new System.Drawing.Point(139, 355);
		this.BtnApplyFastSet.Name = "BtnApplyFastSet";
		this.BtnApplyFastSet.Size = new System.Drawing.Size(118, 30);
		this.BtnApplyFastSet.TabIndex = 8;
		this.BtnApplyFastSet.Text = "应用以上设置";
		this.BtnApplyFastSet.UseVisualStyleBackColor = true;
		this.BtnApplyFastSet.Click += new System.EventHandler(BtnApplyFastSet_Click);
		this.ChkNumber.AutoSize = true;
		this.ChkNumber.Enabled = false;
		this.ChkNumber.Location = new System.Drawing.Point(24, 47);
		this.ChkNumber.Name = "ChkNumber";
		this.ChkNumber.Size = new System.Drawing.Size(84, 24);
		this.ChkNumber.TabIndex = 0;
		this.ChkNumber.Text = "编号缩进";
		this.ChkNumber.UseVisualStyleBackColor = true;
		this.ChkNumber.CheckedChanged += new System.EventHandler(ChkNumber_CheckedChanged);
		this.ChkText.AutoSize = true;
		this.ChkText.Enabled = false;
		this.ChkText.Location = new System.Drawing.Point(114, 47);
		this.ChkText.Name = "ChkText";
		this.ChkText.Size = new System.Drawing.Size(84, 24);
		this.ChkText.TabIndex = 1;
		this.ChkText.Text = "文本缩进";
		this.ChkText.UseVisualStyleBackColor = true;
		this.ChkText.CheckedChanged += new System.EventHandler(ChkNumber_CheckedChanged);
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(21, 208);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(121, 20);
		this.label4.TabIndex = 8;
		this.label4.Text = "以下每级递进缩进";
		this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.ChkAfterNumber.AutoSize = true;
		this.ChkAfterNumber.Enabled = false;
		this.ChkAfterNumber.Location = new System.Drawing.Point(24, 76);
		this.ChkAfterNumber.Name = "ChkAfterNumber";
		this.ChkAfterNumber.Size = new System.Drawing.Size(98, 24);
		this.ChkAfterNumber.TabIndex = 2;
		this.ChkAfterNumber.Text = "编号后缩进";
		this.ChkAfterNumber.UseVisualStyleBackColor = true;
		this.ChkAfterNumber.CheckedChanged += new System.EventHandler(ChkNumber_CheckedChanged);
		this.ChkAutoIndent.AutoSize = true;
		this.ChkAutoIndent.Enabled = false;
		this.ChkAutoIndent.Location = new System.Drawing.Point(24, 176);
		this.ChkAutoIndent.Name = "ChkAutoIndent";
		this.ChkAutoIndent.Size = new System.Drawing.Size(126, 24);
		this.ChkAutoIndent.TabIndex = 4;
		this.ChkAutoIndent.Text = "一级编号缩进为";
		this.ChkAutoIndent.UseVisualStyleBackColor = true;
		this.ChkAutoIndent.CheckedChanged += new System.EventHandler(ChkAutoIndent_CheckedChanged);
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(20, 105);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(65, 20);
		this.label2.TabIndex = 3;
		this.label2.Text = "统一为：";
		this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
		this.Lab_LinkedStyle.Location = new System.Drawing.Point(620, 10);
		this.Lab_LinkedStyle.Name = "Lab_LinkedStyle";
		this.Lab_LinkedStyle.Size = new System.Drawing.Size(100, 20);
		this.Lab_LinkedStyle.TabIndex = 76;
		this.Lab_LinkedStyle.Text = "链接样式";
		this.Lab_LinkedStyle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.GetCurrentListTemplate.Location = new System.Drawing.Point(414, 364);
		this.GetCurrentListTemplate.Name = "GetCurrentListTemplate";
		this.GetCurrentListTemplate.Size = new System.Drawing.Size(150, 30);
		this.GetCurrentListTemplate.TabIndex = 69;
		this.GetCurrentListTemplate.Text = "载入当前多级列表";
		this.GetCurrentListTemplate.UseVisualStyleBackColor = true;
		this.GetCurrentListTemplate.Click += new System.EventHandler(GetCurrentListTemplate_Click);
		this.CreateListTemplate.Location = new System.Drawing.Point(570, 364);
		this.CreateListTemplate.Name = "CreateListTemplate";
		this.CreateListTemplate.Size = new System.Drawing.Size(150, 30);
		this.CreateListTemplate.TabIndex = 70;
		this.CreateListTemplate.Text = "设置多级列表";
		this.CreateListTemplate.UseVisualStyleBackColor = true;
		this.CreateListTemplate.Click += new System.EventHandler(CreateListTemplate_Click);
		this.Lab_AfterNumIndent.Location = new System.Drawing.Point(510, 10);
		this.Lab_AfterNumIndent.Name = "Lab_AfterNumIndent";
		this.Lab_AfterNumIndent.Size = new System.Drawing.Size(100, 20);
		this.Lab_AfterNumIndent.TabIndex = 75;
		this.Lab_AfterNumIndent.Text = "编号后缩进";
		this.Lab_AfterNumIndent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.Lab_TextIndent.Location = new System.Drawing.Point(400, 10);
		this.Lab_TextIndent.Name = "Lab_TextIndent";
		this.Lab_TextIndent.Size = new System.Drawing.Size(100, 20);
		this.Lab_TextIndent.TabIndex = 74;
		this.Lab_TextIndent.Text = "文本缩进";
		this.Lab_TextIndent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.Lab_NumIndent.Location = new System.Drawing.Point(290, 10);
		this.Lab_NumIndent.Name = "Lab_NumIndent";
		this.Lab_NumIndent.Size = new System.Drawing.Size(100, 20);
		this.Lab_NumIndent.TabIndex = 68;
		this.Lab_NumIndent.Text = "编号缩进";
		this.Lab_NumIndent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.Lab_NumberFormat.Location = new System.Drawing.Point(180, 10);
		this.Lab_NumberFormat.Name = "Lab_NumberFormat";
		this.Lab_NumberFormat.Size = new System.Drawing.Size(100, 20);
		this.Lab_NumberFormat.TabIndex = 73;
		this.Lab_NumberFormat.Text = "编号格式";
		this.Lab_NumberFormat.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.Lab_NumberStyle.Location = new System.Drawing.Point(70, 10);
		this.Lab_NumberStyle.Name = "Lab_NumberStyle";
		this.Lab_NumberStyle.Size = new System.Drawing.Size(100, 20);
		this.Lab_NumberStyle.TabIndex = 71;
		this.Lab_NumberStyle.Text = "编号样式";
		this.Lab_NumberStyle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
		this.label1.AutoSize = true;
		this.label1.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		this.label1.Location = new System.Drawing.Point(10, 370);
		this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(107, 20);
		this.label1.TabIndex = 67;
		this.label1.Text = "设计列表级数：";
		this.groupBox1.Controls.Add(this.Chk_UnlinkToTitles);
		this.groupBox1.Controls.Add(this.Chk_LinkToTitles);
		this.groupBox1.Controls.Add(this.LevelFristIndent);
		this.groupBox1.Controls.Add(this.label6);
		this.groupBox1.Controls.Add(this.label3);
		this.groupBox1.Controls.Add(this.BtnApplyFastSet);
		this.groupBox1.Controls.Add(this.ChkNumber);
		this.groupBox1.Controls.Add(this.ChkText);
		this.groupBox1.Controls.Add(this.label5);
		this.groupBox1.Controls.Add(this.ChkAfterNumber);
		this.groupBox1.Controls.Add(this.UnifiedIndent);
		this.groupBox1.Controls.Add(this.label2);
		this.groupBox1.Controls.Add(this.ChkAutoIndent);
		this.groupBox1.Controls.Add(this.label4);
		this.groupBox1.Controls.Add(this.AfterFristIndent);
		this.groupBox1.Location = new System.Drawing.Point(728, 5);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(263, 392);
		this.groupBox1.TabIndex = 77;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "快捷设置";
		this.LevelFristIndent.DecimalPlaces = 2;
		this.LevelFristIndent.Enabled = false;
		this.LevelFristIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.LevelFristIndent.Label = "厘米";
		this.LevelFristIndent.Location = new System.Drawing.Point(157, 174);
		this.LevelFristIndent.Name = "LevelFristIndent";
		this.LevelFristIndent.Size = new System.Drawing.Size(100, 26);
		this.LevelFristIndent.TabIndex = 5;
		this.UnifiedIndent.DecimalPlaces = 2;
		this.UnifiedIndent.Enabled = false;
		this.UnifiedIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.UnifiedIndent.Label = "厘米";
		this.UnifiedIndent.Location = new System.Drawing.Point(157, 102);
		this.UnifiedIndent.Name = "UnifiedIndent";
		this.UnifiedIndent.Size = new System.Drawing.Size(100, 26);
		this.UnifiedIndent.TabIndex = 3;
		this.AfterFristIndent.DecimalPlaces = 2;
		this.AfterFristIndent.Enabled = false;
		this.AfterFristIndent.Increment = new decimal(new int[4] { 1, 0, 0, 131072 });
		this.AfterFristIndent.Label = "厘米";
		this.AfterFristIndent.Location = new System.Drawing.Point(157, 206);
		this.AfterFristIndent.Name = "AfterFristIndent";
		this.AfterFristIndent.Size = new System.Drawing.Size(100, 26);
		this.AfterFristIndent.TabIndex = 6;
		this.Chk_LinkToTitles.AutoSize = true;
		this.Chk_LinkToTitles.Enabled = false;
		this.Chk_LinkToTitles.Location = new System.Drawing.Point(24, 274);
		this.Chk_LinkToTitles.Name = "Chk_LinkToTitles";
		this.Chk_LinkToTitles.Size = new System.Drawing.Size(154, 24);
		this.Chk_LinkToTitles.TabIndex = 14;
		this.Chk_LinkToTitles.Text = "链接到对应标题样式";
		this.Chk_LinkToTitles.UseVisualStyleBackColor = true;
		this.Chk_LinkToTitles.CheckedChanged += new System.EventHandler(Chk_LinkToTitles_CheckedChanged);
		this.Chk_UnlinkToTitles.AutoSize = true;
		this.Chk_UnlinkToTitles.Enabled = false;
		this.Chk_UnlinkToTitles.Location = new System.Drawing.Point(24, 301);
		this.Chk_UnlinkToTitles.Name = "Chk_UnlinkToTitles";
		this.Chk_UnlinkToTitles.Size = new System.Drawing.Size(168, 24);
		this.Chk_UnlinkToTitles.TabIndex = 15;
		this.Chk_UnlinkToTitles.Text = "取消对应标题样式链接";
		this.Chk_UnlinkToTitles.UseVisualStyleBackColor = true;
		this.Chk_UnlinkToTitles.CheckedChanged += new System.EventHandler(Chk_UnlinkToTitles_CheckedChanged);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.Controls.Add(this.groupBox1);
		base.Controls.Add(this.Cmb_LstTLevels);
		base.Controls.Add(this.Lab_LinkedStyle);
		base.Controls.Add(this.GetCurrentListTemplate);
		base.Controls.Add(this.CreateListTemplate);
		base.Controls.Add(this.Lab_AfterNumIndent);
		base.Controls.Add(this.Lab_TextIndent);
		base.Controls.Add(this.Lab_NumIndent);
		base.Controls.Add(this.Lab_NumberFormat);
		base.Controls.Add(this.Lab_NumberStyle);
		base.Controls.Add(this.label1);
		this.DoubleBuffered = true;
		this.Font = new System.Drawing.Font("微软雅黑", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 134);
		base.Name = "LevelListSetUI";
		base.Size = new System.Drawing.Size(1000, 405);
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.LevelFristIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.UnifiedIndent).EndInit();
		((System.ComponentModel.ISupportInitialize)this.AfterFristIndent).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}