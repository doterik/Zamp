using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace OutlookAddin
{
	/// <summary>
	/// Summary description for MessageForm.
	/// </summary>
	public class MessageForm : System.Windows.Forms.Form
	{
		public System.Windows.Forms.RichTextBox tbMessage;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label9;
		public System.Windows.Forms.Label lbFrom;
		public System.Windows.Forms.Label lbTo;
		public System.Windows.Forms.Label lbCC;
		public System.Windows.Forms.Label lbSubject;
		public System.Windows.Forms.Label lbSentOn;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MessageForm()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
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
			this.tbMessage = new System.Windows.Forms.RichTextBox();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.lbFrom = new System.Windows.Forms.Label();
			this.lbTo = new System.Windows.Forms.Label();
			this.lbCC = new System.Windows.Forms.Label();
			this.lbSubject = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.lbSentOn = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// tbMessage
			// 
			this.tbMessage.BackColor = System.Drawing.SystemColors.Window;
			this.tbMessage.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.tbMessage.Location = new System.Drawing.Point(0, 112);
			this.tbMessage.Name = "tbMessage";
			this.tbMessage.ReadOnly = true;
			this.tbMessage.Size = new System.Drawing.Size(504, 200);
			this.tbMessage.TabIndex = 0;
			this.tbMessage.Text = "";
			// 
			// splitter1
			// 
			this.splitter1.BackColor = System.Drawing.SystemColors.ControlDark;
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.splitter1.Enabled = false;
			this.splitter1.Location = new System.Drawing.Point(0, 111);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(504, 1);
			this.splitter1.TabIndex = 1;
			this.splitter1.TabStop = false;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 23);
			this.label1.TabIndex = 2;
			this.label1.Text = "From:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(12, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 23);
			this.label2.TabIndex = 3;
			this.label2.Text = "To:";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(12, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 23);
			this.label3.TabIndex = 4;
			this.label3.Text = "CC:";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(12, 80);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(48, 23);
			this.label4.TabIndex = 5;
			this.label4.Text = "Subject:";
			// 
			// lbFrom
			// 
			this.lbFrom.Location = new System.Drawing.Point(64, 8);
			this.lbFrom.Name = "lbFrom";
			this.lbFrom.Size = new System.Drawing.Size(236, 23);
			this.lbFrom.TabIndex = 6;
			// 
			// lbTo
			// 
			this.lbTo.Location = new System.Drawing.Point(64, 32);
			this.lbTo.Name = "lbTo";
			this.lbTo.Size = new System.Drawing.Size(432, 23);
			this.lbTo.TabIndex = 7;
			// 
			// lbCC
			// 
			this.lbCC.Location = new System.Drawing.Point(64, 56);
			this.lbCC.Name = "lbCC";
			this.lbCC.Size = new System.Drawing.Size(432, 23);
			this.lbCC.TabIndex = 8;
			// 
			// lbSubject
			// 
			this.lbSubject.Location = new System.Drawing.Point(64, 80);
			this.lbSubject.Name = "lbSubject";
			this.lbSubject.Size = new System.Drawing.Size(432, 23);
			this.lbSubject.TabIndex = 9;
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(306, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 23);
			this.label9.TabIndex = 10;
			this.label9.Text = "Sent:";
			// 
			// lbSentOn
			// 
			this.lbSentOn.Location = new System.Drawing.Point(352, 8);
			this.lbSentOn.Name = "lbSentOn";
			this.lbSentOn.Size = new System.Drawing.Size(144, 23);
			this.lbSentOn.TabIndex = 11;
			// 
			// MessageForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 312);
			this.Controls.Add(this.lbSentOn);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.lbSubject);
			this.Controls.Add(this.lbCC);
			this.Controls.Add(this.lbTo);
			this.Controls.Add(this.lbFrom);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.tbMessage);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(520, 350);
			this.Name = "MessageForm";
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Message";
			this.ResumeLayout(false);

		}
		#endregion
	}
}
