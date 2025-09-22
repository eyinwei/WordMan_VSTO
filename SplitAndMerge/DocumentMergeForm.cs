using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WordMan_VSTO;

namespace WordMan_VSTO.SplitAndMerge
{
    public partial class DocumentMergeForm : Form
    {
        public List<string> SelectedFiles { get; private set; }
        public MergeOptions MergeOptions { get; private set; }

        private ListBox lstFiles;
        private StandardButton btnMoveUp;
        private StandardButton btnMoveDown;
        private StandardButton btnRemove;
        private StandardButton btnAdd;
        private StandardButton btnOK;
        private StandardButton btnCancel;
        private Label lblFiles;
        private Label lblBreakType;
        private RadioButton rdoPageBreak;
        private RadioButton rdoSectionBreak;

        public DocumentMergeForm(List<string> filePaths)
        {
            InitializeComponent();
            SelectedFiles = new List<string>(filePaths);
            MergeOptions = new MergeOptions();
            LoadFiles();
        }

        private void InitializeComponent()
        {
            this.lstFiles = new System.Windows.Forms.ListBox();
            this.btnMoveUp = new StandardButton(StandardButton.ButtonType.Small, "上移", new Size(50, 35), new Point(318, 35));
            this.btnMoveDown = new StandardButton(StandardButton.ButtonType.Small, "下移", new Size(50, 35), new Point(318, 75));
            this.btnRemove = new StandardButton(StandardButton.ButtonType.Small, "移除", new Size(50, 35), new Point(318, 115));
            this.btnAdd = new StandardButton(StandardButton.ButtonType.Small, "添加", new Size(50, 35), new Point(318, 155));
            this.btnOK = new StandardButton(StandardButton.ButtonType.Primary, "确定", new Size(75, 35), new Point(211, 240));
            this.btnCancel = new StandardButton(StandardButton.ButtonType.Secondary, "取消", new Size(75, 35), new Point(292, 240));
            this.lblFiles = new System.Windows.Forms.Label();
            this.lblBreakType = new System.Windows.Forms.Label();
            this.rdoPageBreak = new System.Windows.Forms.RadioButton();
            this.rdoSectionBreak = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // lstFiles
            // 
            this.lstFiles.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lstFiles.FormattingEnabled = true;
            this.lstFiles.ItemHeight = 19;
            this.lstFiles.Location = new System.Drawing.Point(12, 35);
            this.lstFiles.Name = "lstFiles";
            this.lstFiles.Size = new System.Drawing.Size(300, 156);
            this.lstFiles.TabIndex = 1;
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.TabIndex = 2;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.TabIndex = 3;
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.TabIndex = 4;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Name = "btnOK";
            this.btnOK.TabIndex = 9;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.TabIndex = 10;
            // 
            // lblFiles
            // 
            this.lblFiles.AutoSize = true;
            this.lblFiles.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblFiles.Location = new System.Drawing.Point(12, 9);
            this.lblFiles.Name = "lblFiles";
            this.lblFiles.Size = new System.Drawing.Size(135, 20);
            this.lblFiles.TabIndex = 0;
            this.lblFiles.Text = "选择要合并的文档：";
            // 
            // lblBreakType
            // 
            this.lblBreakType.AutoSize = true;
            this.lblBreakType.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBreakType.Location = new System.Drawing.Point(12, 205);
            this.lblBreakType.Name = "lblBreakType";
            this.lblBreakType.Size = new System.Drawing.Size(107, 20);
            this.lblBreakType.TabIndex = 6;
            this.lblBreakType.Text = "文档分隔方式：";
            // 
            // rdoPageBreak
            // 
            this.rdoPageBreak.AutoSize = true;
            this.rdoPageBreak.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rdoPageBreak.Location = new System.Drawing.Point(140, 205);
            this.rdoPageBreak.Name = "rdoPageBreak";
            this.rdoPageBreak.Size = new System.Drawing.Size(69, 24);
            this.rdoPageBreak.TabIndex = 7;
            this.rdoPageBreak.Text = "分页符";
            this.rdoPageBreak.UseVisualStyleBackColor = true;
            // 
            // rdoSectionBreak
            // 
            this.rdoSectionBreak.AutoSize = true;
            this.rdoSectionBreak.Checked = true;
            this.rdoSectionBreak.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rdoSectionBreak.Location = new System.Drawing.Point(230, 205);
            this.rdoSectionBreak.Name = "rdoSectionBreak";
            this.rdoSectionBreak.Size = new System.Drawing.Size(69, 24);
            this.rdoSectionBreak.TabIndex = 8;
            this.rdoSectionBreak.TabStop = true;
            this.rdoSectionBreak.Text = "分节符";
            this.rdoSectionBreak.UseVisualStyleBackColor = true;
            // 
            // DocumentMergeForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(380, 290);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.rdoSectionBreak);
            this.Controls.Add(this.rdoPageBreak);
            this.Controls.Add(this.lblBreakType);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnMoveDown);
            this.Controls.Add(this.btnMoveUp);
            this.Controls.Add(this.lstFiles);
            this.Controls.Add(this.lblFiles);
            this.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DocumentMergeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "文档合并";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void LoadFiles()
        {
            lstFiles.Items.Clear();
            foreach (var file in SelectedFiles)
            {
                lstFiles.Items.Add(Path.GetFileName(file));
            }
        }

        private void btnMoveUp_Click(object sender, EventArgs e)
        {
            if (lstFiles.SelectedIndex > 0)
            {
                int selectedIndex = lstFiles.SelectedIndex;
                var selectedFile = SelectedFiles[selectedIndex];
                
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(selectedIndex - 1, selectedFile);
                
                LoadFiles();
                lstFiles.SelectedIndex = selectedIndex - 1;
            }
        }

        private void btnMoveDown_Click(object sender, EventArgs e)
        {
            if (lstFiles.SelectedIndex >= 0 && lstFiles.SelectedIndex < SelectedFiles.Count - 1)
            {
                int selectedIndex = lstFiles.SelectedIndex;
                var selectedFile = SelectedFiles[selectedIndex];
                
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(selectedIndex + 1, selectedFile);
                
                LoadFiles();
                lstFiles.SelectedIndex = selectedIndex + 1;
            }
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (lstFiles.SelectedIndex >= 0)
            {
                int selectedIndex = lstFiles.SelectedIndex;
                SelectedFiles.RemoveAt(selectedIndex);
                LoadFiles();
                
                if (SelectedFiles.Count > 0)
                {
                    if (selectedIndex >= SelectedFiles.Count)
                        selectedIndex = SelectedFiles.Count - 1;
                    lstFiles.SelectedIndex = selectedIndex;
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择要添加的Word文档",
                Filter = "Word文档 (*.docx)|*.docx|Word文档 (*.doc)|*.doc|所有文件 (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var filePath in openFileDialog.FileNames)
                {
                    if (!SelectedFiles.Contains(filePath))
                    {
                        SelectedFiles.Add(filePath);
                    }
                }
                LoadFiles();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (SelectedFiles.Count < 2)
            {
                MessageBox.Show("请至少选择2个文档进行合并。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 使用默认合并选项（所有选项都启用）
            MergeOptions.AddPageBreaks = true;
            MergeOptions.PreserveFormatting = true;
            MergeOptions.CopyStyles = true;
            MergeOptions.UseSectionBreak = rdoSectionBreak.Checked;

            DialogResult = DialogResult.OK;
        }
    }
}
