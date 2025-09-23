using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式选择窗体
    /// </summary>
    public partial class StyleSelectionForm : Form
    {
        private CheckedListBox checkedListBox1;
        private Button btnOK;
        private Button btnCancel;
        private Label label1;
        
        public List<string> SelectedStyles { get; private set; }
        private List<string> AvailableStyles { get; set; }

        public StyleSelectionForm()
        {
            InitializeComponent();
            SelectedStyles = new List<string>();
        }

        /// <summary>
        /// 初始化可用样式列表
        /// </summary>
        public void InitializeStyles(List<string> styles)
        {
            AvailableStyles = styles ?? new List<string>();
            checkedListBox1.Items.Clear();
            
            foreach (var style in AvailableStyles)
            {
                checkedListBox1.Items.Add(style, false);
            }
        }

        /// <summary>
        /// 设置已选中的样式
        /// </summary>
        public void SetSelectedStyles(List<string> selectedStyles)
        {
            if (selectedStyles == null) return;
            
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                string styleName = checkedListBox1.Items[i].ToString();
                checkedListBox1.SetItemChecked(i, selectedStyles.Contains(styleName));
            }
        }

        private void InitializeComponent()
        {
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(10, 25);
            this.checkedListBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(309, 356);
            this.checkedListBox1.TabIndex = 1;
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(138, 389);
            this.btnOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(86, 25);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(233, 389);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(86, 25);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "选择要加载的样式：";
            // 
            // StyleSelectionForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(329, 425);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "StyleSelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "样式选择";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedStyles.Clear();
            
            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                SelectedStyles.Add(checkedListBox1.CheckedItems[i].ToString());
            }
        }
    }
}
