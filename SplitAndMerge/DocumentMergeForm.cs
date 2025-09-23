using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WordMan;

namespace WordMan.SplitAndMerge
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
            var font = new Font("微软雅黑", 10F);
            
            // 创建控件
            lstFiles = new ListBox { Font = font, Location = new Point(12, 35), Size = new Size(300, 156) };
            btnMoveUp = new StandardButton(StandardButton.ButtonType.Small, "上移", new Size(50, 35), new Point(318, 35));
            btnMoveDown = new StandardButton(StandardButton.ButtonType.Small, "下移", new Size(50, 35), new Point(318, 75));
            btnRemove = new StandardButton(StandardButton.ButtonType.Small, "移除", new Size(50, 35), new Point(318, 115));
            btnAdd = new StandardButton(StandardButton.ButtonType.Small, "添加", new Size(50, 35), new Point(318, 155));
            btnOK = new StandardButton(StandardButton.ButtonType.Primary, "确定", new Size(75, 35), new Point(211, 240));
            btnCancel = new StandardButton(StandardButton.ButtonType.Secondary, "取消", new Size(75, 35), new Point(292, 240));
            lblFiles = new Label { AutoSize = true, Font = font, Location = new Point(12, 9), Text = "选择要合并的文档：" };
            lblBreakType = new Label { AutoSize = true, Font = font, Location = new Point(12, 205), Text = "文档分隔方式：" };
            rdoPageBreak = new RadioButton { AutoSize = true, Font = font, Location = new Point(140, 205), Text = "分页符" };
            rdoSectionBreak = new RadioButton { AutoSize = true, Checked = true, Font = font, Location = new Point(230, 205), Text = "分节符" };

            // 设置事件
            btnMoveUp.Click += btnMoveUp_Click;
            btnMoveDown.Click += btnMoveDown_Click;
            btnRemove.Click += btnRemove_Click;
            btnAdd.Click += btnAdd_Click;
            btnOK.Click += btnOK_Click;

            // 设置窗体属性
            AcceptButton = btnOK;
            CancelButton = btnCancel;
            ClientSize = new Size(380, 290);
            Font = font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "文档合并";

            // 添加控件
            Controls.AddRange(new Control[] { 
                btnCancel, btnOK, rdoSectionBreak, rdoPageBreak, lblBreakType,
                btnAdd, btnRemove, btnMoveDown, btnMoveUp, lstFiles, lblFiles 
            });
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
            MoveFile(-1);
        }

        private void btnMoveDown_Click(object sender, EventArgs e)
        {
            MoveFile(1);
        }

        private void MoveFile(int direction)
        {
            if (lstFiles.SelectedIndex < 0) return;
            
            var selectedIndex = lstFiles.SelectedIndex;
            var newIndex = selectedIndex + direction;
            
            if (newIndex >= 0 && newIndex < SelectedFiles.Count)
            {
                var selectedFile = SelectedFiles[selectedIndex];
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(newIndex, selectedFile);
                
                LoadFiles();
                lstFiles.SelectedIndex = newIndex;
            }
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (lstFiles.SelectedIndex >= 0)
            {
                var selectedIndex = lstFiles.SelectedIndex;
                SelectedFiles.RemoveAt(selectedIndex);
                LoadFiles();
                
                if (SelectedFiles.Count > 0)
                {
                    var newIndex = Math.Min(selectedIndex, SelectedFiles.Count - 1);
                    lstFiles.SelectedIndex = newIndex;
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
                foreach (var filePath in openFileDialog.FileNames.Where(fp => !SelectedFiles.Contains(fp)))
                {
                    SelectedFiles.Add(filePath);
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

            MergeOptions.AddPageBreaks = true;
            MergeOptions.PreserveFormatting = true;
            MergeOptions.CopyStyles = true;
            MergeOptions.UseSectionBreak = rdoSectionBreak.Checked;

            DialogResult = DialogResult.OK;
        }
    }
}
