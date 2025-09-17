// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.BuildInStyleSeletor
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using WordFormatHelper.Properties;

public class BuildInStyleSeletor : Form
{
	private IContainer components;

	private CheckedListBox Lst_BuildInStyles;

	private Button Btn_SelectComfrim;

	public int[] SelectedIndices { get; private set; } = Array.Empty<int>();

	public BuildInStyleSeletor(string[] buildInStyleNames, int[] checkedIndex)
	{
		InitializeComponent();
		base.Icon = Resources.WAIcon;
		Lst_BuildInStyles.Items.Clear();
		Lst_BuildInStyles.Items.AddRange(buildInStyleNames);
		if (checkedIndex.Length != 0)
		{
			foreach (int index in checkedIndex)
			{
				Lst_BuildInStyles.SetItemChecked(index, value: true);
			}
		}
		base.DialogResult = DialogResult.Cancel;
	}

	private void Btn_SelectComfrim_Click(object sender, EventArgs e)
	{
		if (Lst_BuildInStyles.CheckedIndices.Count > 0)
		{
			List<int> list = new List<int>();
			foreach (int checkedIndex in Lst_BuildInStyles.CheckedIndices)
			{
				list.Add(checkedIndex);
			}
			SelectedIndices = list.ToArray();
			base.DialogResult = DialogResult.OK;
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
		this.Lst_BuildInStyles = new System.Windows.Forms.CheckedListBox();
		this.Btn_SelectComfrim = new System.Windows.Forms.Button();
		base.SuspendLayout();
		this.Lst_BuildInStyles.CheckOnClick = true;
		this.Lst_BuildInStyles.Dock = System.Windows.Forms.DockStyle.Top;
		this.Lst_BuildInStyles.FormattingEnabled = true;
		this.Lst_BuildInStyles.Location = new System.Drawing.Point(0, 0);
		this.Lst_BuildInStyles.Name = "Lst_BuildInStyles";
		this.Lst_BuildInStyles.Size = new System.Drawing.Size(284, 424);
		this.Lst_BuildInStyles.TabIndex = 0;
		this.Btn_SelectComfrim.Dock = System.Windows.Forms.DockStyle.Bottom;
		this.Btn_SelectComfrim.Location = new System.Drawing.Point(0, 431);
		this.Btn_SelectComfrim.Name = "Btn_SelectComfrim";
		this.Btn_SelectComfrim.Size = new System.Drawing.Size(284, 30);
		this.Btn_SelectComfrim.TabIndex = 1;
		this.Btn_SelectComfrim.Text = "确定";
		this.Btn_SelectComfrim.UseVisualStyleBackColor = true;
		this.Btn_SelectComfrim.Click += new System.EventHandler(Btn_SelectComfrim_Click);
		base.AutoScaleDimensions = new System.Drawing.SizeF(96f, 96f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
		this.BackColor = System.Drawing.Color.AliceBlue;
		base.ClientSize = new System.Drawing.Size(284, 461);
		base.Controls.Add(this.Btn_SelectComfrim);
		base.Controls.Add(this.Lst_BuildInStyles);
		this.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
		base.Margin = new System.Windows.Forms.Padding(4);
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "BuildInStyleSeletor";
		base.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
		this.Text = "Word内置样式选择";
		base.ResumeLayout(false);
	}
}
