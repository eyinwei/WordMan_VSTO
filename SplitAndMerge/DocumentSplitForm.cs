using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;

namespace WordMan_VSTO
{
    public partial class DocumentSplitForm : Form
    {
        public SplitMode SplitMode { get; private set; }
        public List<PageRange> PageRanges { get; private set; }
        public Action<int, int, string> ProgressCallback { get; private set; }

        private RadioButton rbPageByPage;
        private RadioButton rbCustomRanges;
        private TextBox txtPageRanges;
        private Label lblPageRanges;
        private StandardButton btnOK;
        private StandardButton btnCancel;
        private Label lblInstructions;
        private ProgressBar progressBar;
        private Label lblProgress;

        public DocumentSplitForm()
        {
            InitializeComponent();
            PageRanges = new List<PageRange>();
        }

        private void InitializeComponent()
        {
            this.rbPageByPage = new System.Windows.Forms.RadioButton();
            this.rbCustomRanges = new System.Windows.Forms.RadioButton();
            this.txtPageRanges = new System.Windows.Forms.TextBox();
            this.lblPageRanges = new System.Windows.Forms.Label();
            this.btnOK = new StandardButton(StandardButton.ButtonType.Primary, "确定", new Size(75, 35), new Point(139, 235));
            this.btnCancel = new StandardButton(StandardButton.ButtonType.Secondary, "取消", new Size(80, 35), new Point(235, 235));
            this.lblInstructions = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // rbPageByPage
            // 
            this.rbPageByPage.AutoSize = true;
            this.rbPageByPage.Checked = true;
            this.rbPageByPage.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.rbPageByPage.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(10)))), ((int)(((byte)(10)))));
            this.rbPageByPage.Location = new System.Drawing.Point(15, 48);
            this.rbPageByPage.Name = "rbPageByPage";
            this.rbPageByPage.Size = new System.Drawing.Size(74, 21);
            this.rbPageByPage.TabIndex = 1;
            this.rbPageByPage.TabStop = true;
            this.rbPageByPage.Text = "逐页拆分";
            this.rbPageByPage.UseVisualStyleBackColor = true;
            this.rbPageByPage.CheckedChanged += new System.EventHandler(this.rbPageByPage_CheckedChanged);
            // 
            // rbCustomRanges
            // 
            this.rbCustomRanges.AutoSize = true;
            this.rbCustomRanges.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.rbCustomRanges.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(10)))), ((int)(((byte)(10)))));
            this.rbCustomRanges.Location = new System.Drawing.Point(15, 78);
            this.rbCustomRanges.Name = "rbCustomRanges";
            this.rbCustomRanges.Size = new System.Drawing.Size(86, 21);
            this.rbCustomRanges.TabIndex = 2;
            this.rbCustomRanges.Text = "自定义范围";
            this.rbCustomRanges.UseVisualStyleBackColor = true;
            this.rbCustomRanges.CheckedChanged += new System.EventHandler(this.rbCustomRanges_CheckedChanged);
            // 
            // txtPageRanges
            // 
            this.txtPageRanges.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.txtPageRanges.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPageRanges.Enabled = false;
            this.txtPageRanges.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtPageRanges.Location = new System.Drawing.Point(15, 128);
            this.txtPageRanges.Name = "txtPageRanges";
            this.txtPageRanges.Size = new System.Drawing.Size(300, 23);
            this.txtPageRanges.TabIndex = 4;
            this.txtPageRanges.Text = "1-2,3-5,7";
            // 
            // lblPageRanges
            // 
            this.lblPageRanges.AutoSize = true;
            this.lblPageRanges.Enabled = false;
            this.lblPageRanges.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblPageRanges.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(10)))), ((int)(((byte)(10)))));
            this.lblPageRanges.Location = new System.Drawing.Point(15, 108);
            this.lblPageRanges.Name = "lblPageRanges";
            this.lblPageRanges.Size = new System.Drawing.Size(167, 17);
            this.lblPageRanges.TabIndex = 3;
            this.lblPageRanges.Text = "页码范围（如：1-2,3-5,7）：";
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Name = "btnOK";
            this.btnOK.TabIndex = 10;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.TabIndex = 11;
            // 
            // lblInstructions
            // 
            this.lblInstructions.AutoSize = true;
            this.lblInstructions.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.lblInstructions.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(10)))), ((int)(((byte)(10)))));
            this.lblInstructions.Location = new System.Drawing.Point(15, 18);
            this.lblInstructions.Name = "lblInstructions";
            this.lblInstructions.Size = new System.Drawing.Size(149, 19);
            this.lblInstructions.TabIndex = 0;
            this.lblInstructions.Text = "请选择文档拆分方式：";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(15, 180);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(300, 20);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 8;
            this.progressBar.Visible = false;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblProgress.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(10)))), ((int)(((byte)(10)))));
            this.lblProgress.Location = new System.Drawing.Point(15, 205);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(0, 17);
            this.lblProgress.TabIndex = 9;
            this.lblProgress.Visible = false;
            // 
            // DocumentSplitForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(330, 285);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.txtPageRanges);
            this.Controls.Add(this.lblPageRanges);
            this.Controls.Add(this.rbCustomRanges);
            this.Controls.Add(this.rbPageByPage);
            this.Controls.Add(this.lblInstructions);
            this.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DocumentSplitForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "文档拆分";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void rbPageByPage_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPageByPage.Checked)
            {
                lblPageRanges.Enabled = false;
                txtPageRanges.Enabled = false;
                SplitMode = SplitMode.PageByPage;
            }
        }

        private void rbCustomRanges_CheckedChanged(object sender, EventArgs e)
        {
            if (rbCustomRanges.Checked)
            {
                lblPageRanges.Enabled = true;
                txtPageRanges.Enabled = true;
                SplitMode = SplitMode.CustomRanges;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rbCustomRanges.Checked)
            {
                if (string.IsNullOrWhiteSpace(txtPageRanges.Text))
                {
                    MessageBox.Show("请输入页码范围。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    PageRanges = ParsePageRanges(txtPageRanges.Text);
                    if (PageRanges.Count == 0)
                    {
                        MessageBox.Show("页码范围格式不正确，请使用格式如：1-2,3-5,7", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"页码范围解析失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                // 逐页拆分，不需要解析页码范围
                PageRanges = new List<PageRange>();
            }

            // 设置进度回调
            ProgressCallback = UpdateProgress;

            DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// 更新进度显示
        /// </summary>
        private void UpdateProgress(int current, int total, string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int, int, string>(UpdateProgress), current, total, message);
                return;
            }

            // 显示进度条
            if (!progressBar.Visible)
            {
                progressBar.Visible = true;
                lblProgress.Visible = true;
                progressBar.Maximum = total;
            }

            // 更新进度
            progressBar.Value = Math.Min(current, total);
            lblProgress.Text = message;

            // 刷新界面
            Application.DoEvents();
        }

        /// <summary>
        /// 解析页码范围字符串
        /// </summary>
        private List<PageRange> ParsePageRanges(string pageRangesText)
        {
            var ranges = new List<PageRange>();
            var parts = pageRangesText.Split(',');

            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;

                // 匹配单个页码或页码范围
                var match = Regex.Match(trimmed, @"^(\d+)(?:-(\d+))?$");
                if (match.Success)
                {
                    int startPage = int.Parse(match.Groups[1].Value);
                    int endPage = match.Groups[2].Success ? int.Parse(match.Groups[2].Value) : startPage;

                    if (startPage <= endPage)
                    {
                        ranges.Add(new PageRange(startPage, endPage));
                    }
                }
            }

            return ranges;
        }
    }
}
