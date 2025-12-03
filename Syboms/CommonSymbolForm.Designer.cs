namespace WordMan
{
    partial class CommonSymbolForm
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
            this.components = new System.ComponentModel.Container();
            this.btnClose = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageSymbols = new System.Windows.Forms.TabPage();
            this.tabPageMath = new System.Windows.Forms.TabPage();
            this.tabPageNumbers = new System.Windows.Forms.TabPage();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tabPageExtend = new System.Windows.Forms.TabPage();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
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
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageSymbols);
            this.tabControl1.Controls.Add(this.tabPageMath);
            this.tabControl1.Controls.Add(this.tabPageNumbers);
            this.tabControl1.Controls.Add(this.tabPageExtend);
            this.tabControl1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tabControl1.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(584, 477);
            this.tabControl1.TabIndex = 5;
            // 
            // tabPageSymbols
            // 
            this.tabPageSymbols.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageSymbols.Location = new System.Drawing.Point(4, 36);
            this.tabPageSymbols.Name = "tabPageSymbols";
            this.tabPageSymbols.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSymbols.Size = new System.Drawing.Size(576, 437);
            this.tabPageSymbols.TabIndex = 0;
            this.tabPageSymbols.Text = "符号";
            this.tabPageSymbols.UseVisualStyleBackColor = true;
            // 
            // tabPageMath
            // 
            this.tabPageMath.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageMath.Location = new System.Drawing.Point(4, 36);
            this.tabPageMath.Name = "tabPageMath";
            this.tabPageMath.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMath.Size = new System.Drawing.Size(576, 437);
            this.tabPageMath.TabIndex = 1;
            this.tabPageMath.Text = "数学";
            this.tabPageMath.UseVisualStyleBackColor = true;
            // 
            // tabPageNumbers
            // 
            this.tabPageNumbers.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageNumbers.Location = new System.Drawing.Point(4, 36);
            this.tabPageNumbers.Name = "tabPageNumbers";
            this.tabPageNumbers.Size = new System.Drawing.Size(576, 437);
            this.tabPageNumbers.TabIndex = 2;
            this.tabPageNumbers.Text = "序号";
            this.tabPageNumbers.UseVisualStyleBackColor = true;
            // 
            // tabPageExtend
            // 
            this.tabPageExtend.Font = new System.Drawing.Font("Times New Roman", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageExtend.Location = new System.Drawing.Point(4, 36);
            this.tabPageExtend.Name = "tabPageExtend";
            this.tabPageExtend.Size = new System.Drawing.Size(576, 437);
            this.tabPageExtend.TabIndex = 3;
            this.tabPageExtend.Text = "扩展符号";
            this.tabPageExtend.UseVisualStyleBackColor = true;
            // 
            // CommonSymbolForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 561);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnClose);
            this.CancelButton = this.btnClose;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "CommonSymbolForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CommonSymbolForm";
            this.Load += new System.EventHandler(this.CommonSymbolForm_Load);
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageSymbols;
        private System.Windows.Forms.TabPage tabPageMath;
        private System.Windows.Forms.TabPage tabPageNumbers;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TabPage tabPageExtend;
    }
}