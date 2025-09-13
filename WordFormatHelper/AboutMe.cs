using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace WordFormatHelper{

public class AboutMe : Form
{
	private IContainer components;

	private PictureBox Pic_QR;

	private Label label1;

	private Label Lab_Version;

	private Label label3;

	private Label label4;

	public AboutMe()
	{
		InitializeComponent();
		Lab_Version.Text = "版本：" + Assembly.GetExecutingAssembly().GetName().Version.ToString();
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
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WordFormatHelper.AboutMe));
		this.Pic_QR = new System.Windows.Forms.PictureBox();
		this.label1 = new System.Windows.Forms.Label();
		this.Lab_Version = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label4 = new System.Windows.Forms.Label();
		((System.ComponentModel.ISupportInitialize)this.Pic_QR).BeginInit();
		base.SuspendLayout();
		this.Pic_QR.Image = (System.Drawing.Image)resources.GetObject("Pic_QR.Image");
		this.Pic_QR.Location = new System.Drawing.Point(369, 12);
		this.Pic_QR.Name = "Pic_QR";
		this.Pic_QR.Size = new System.Drawing.Size(260, 260);
		this.Pic_QR.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
		this.Pic_QR.TabIndex = 0;
		this.Pic_QR.TabStop = false;
		this.label1.AutoSize = true;
		this.label1.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		this.label1.Location = new System.Drawing.Point(12, 21);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(171, 18);
		this.label1.TabIndex = 1;
		this.label1.Text = "插件名称：Word格式助手";
		this.Lab_Version.AutoSize = true;
		this.Lab_Version.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		this.Lab_Version.Location = new System.Drawing.Point(12, 59);
		this.Lab_Version.Name = "Lab_Version";
		this.Lab_Version.Size = new System.Drawing.Size(78, 18);
		this.Lab_Version.TabIndex = 2;
		this.Lab_Version.Text = "版本：V1.1";
		this.label3.AutoSize = true;
		this.label3.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		this.label3.Location = new System.Drawing.Point(12, 97);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(109, 18);
		this.label3.TabIndex = 3;
		this.label3.Text = "作者：YinDong";
		this.label4.Font = new System.Drawing.Font("Microsoft JhengHei UI", 10.5f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		this.label4.Location = new System.Drawing.Point(12, 135);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(357, 137);
		this.label4.TabIndex = 4;
		this.label4.Text = "关于本插件：本插件为作者业余兴趣编写，不含任何商业性质，本插件免费使用，供大家学习交流。\r\n\r\n如果在您使用中觉得本插件对您有一定的帮助，也欢迎您给作者些许鼓励！\r\n\r\n扫描右侧二维码，可以给作者打赏！";
		base.AutoScaleDimensions = new System.Drawing.SizeF(96f, 96f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
		this.BackColor = System.Drawing.Color.White;
		base.ClientSize = new System.Drawing.Size(638, 275);
		base.Controls.Add(this.label4);
		base.Controls.Add(this.label3);
		base.Controls.Add(this.Lab_Version);
		base.Controls.Add(this.label1);
		base.Controls.Add(this.Pic_QR);
		base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = "AboutMe";
		this.Text = "关于Word格式助手";
		((System.ComponentModel.ISupportInitialize)this.Pic_QR).EndInit();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
}