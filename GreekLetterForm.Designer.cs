namespace WordMan_VSTO
{
    partial class GreekLetterForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.chkItalic = new System.Windows.Forms.CheckBox();
            this.chkUppercase = new System.Windows.Forms.CheckBox();
            this.chkBold = new System.Windows.Forms.CheckBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.tableLetters = new System.Windows.Forms.TableLayoutPanel();
            this.SuspendLayout();
            // 
            // chkItalic
            // 
            this.chkItalic.AutoSize = true;
            this.chkItalic.BackColor = System.Drawing.SystemColors.Control;
            this.chkItalic.Cursor = System.Windows.Forms.Cursors.Hand;
            this.chkItalic.Font = new System.Drawing.Font("微软雅黑", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkItalic.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.chkItalic.Location = new System.Drawing.Point(28, 408);
            this.chkItalic.Name = "chkItalic";
            this.chkItalic.Size = new System.Drawing.Size(101, 45);
            this.chkItalic.TabIndex = 0;
            this.chkItalic.Text = "斜体";
            this.chkItalic.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkItalic.UseVisualStyleBackColor = false;
            // 
            // chkUppercase
            // 
            this.chkUppercase.AutoSize = true;
            this.chkUppercase.Cursor = System.Windows.Forms.Cursors.Hand;
            this.chkUppercase.Font = new System.Drawing.Font("微软雅黑", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkUppercase.Location = new System.Drawing.Point(249, 408);
            this.chkUppercase.Name = "chkUppercase";
            this.chkUppercase.Size = new System.Drawing.Size(101, 45);
            this.chkUppercase.TabIndex = 1;
            this.chkUppercase.Text = "大写";
            this.chkUppercase.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkUppercase.UseVisualStyleBackColor = true;
            // 
            // chkBold
            // 
            this.chkBold.AutoSize = true;
            this.chkBold.Cursor = System.Windows.Forms.Cursors.Hand;
            this.chkBold.Font = new System.Drawing.Font("微软雅黑", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkBold.Location = new System.Drawing.Point(470, 408);
            this.chkBold.Name = "chkBold";
            this.chkBold.Size = new System.Drawing.Size(101, 45);
            this.chkBold.TabIndex = 2;
            this.chkBold.Text = "粗体";
            this.chkBold.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkBold.UseVisualStyleBackColor = true;
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.Color.RosyBrown;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnClose.Font = new System.Drawing.Font("微软雅黑", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnClose.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnClose.Location = new System.Drawing.Point(180, 483);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(250, 50);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // tableLetters
            // 
            this.tableLetters.AccessibleRole = System.Windows.Forms.AccessibleRole.Grip;
            this.tableLetters.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble;
            this.tableLetters.ColumnCount = 5;
            this.tableLetters.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLetters.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tableLetters.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
            this.tableLetters.Location = new System.Drawing.Point(0, 0);
            this.tableLetters.Name = "tableLetters";
            this.tableLetters.RowCount = 5;
            this.tableLetters.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLetters.Size = new System.Drawing.Size(584, 369);
            this.tableLetters.TabIndex = 4;
            // 
            // GreekLetterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 561);
            this.Controls.Add(this.tableLetters);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.chkBold);
            this.Controls.Add(this.chkUppercase);
            this.Controls.Add(this.chkItalic);
            this.Name = "GreekLetterForm";
            this.Opacity = 0.9D;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "GreekLetterForm";
            this.Load += new System.EventHandler(this.GreekLetterForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkItalic;
        private System.Windows.Forms.CheckBox chkUppercase;
        private System.Windows.Forms.CheckBox chkBold;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TableLayoutPanel tableLetters;
    }
}