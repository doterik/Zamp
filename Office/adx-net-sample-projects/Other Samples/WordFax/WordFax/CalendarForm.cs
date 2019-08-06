using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace WordFax
{
	/// <summary>
	/// Summary description for CalendarForm.
	/// </summary>
	public class CalendarForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Button btnApply;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.MonthCalendar Calendar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CalendarForm()
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
			this.Calendar = new System.Windows.Forms.MonthCalendar();
			this.btnApply = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.SuspendLayout();
			// 
			// Calendar
			// 
			this.Calendar.Dock = System.Windows.Forms.DockStyle.Top;
			this.Calendar.Location = new System.Drawing.Point(0, 0);
			this.Calendar.Name = "Calendar";
			this.Calendar.TabIndex = 0;
			// 
			// btnApply
			// 
			this.btnApply.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnApply.Location = new System.Drawing.Point(12, 169);
			this.btnApply.Name = "btnApply";
			this.btnApply.Size = new System.Drawing.Size(72, 23);
			this.btnApply.TabIndex = 1;
			this.btnApply.Text = "Apply";
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(145, 169);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(72, 23);
			this.btnCancel.TabIndex = 2;
			this.btnCancel.Text = "Cancel";
			// 
			// splitter1
			// 
			this.splitter1.BackColor = System.Drawing.SystemColors.ControlDark;
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
			this.splitter1.Location = new System.Drawing.Point(0, 162);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(229, 1);
			this.splitter1.TabIndex = 3;
			this.splitter1.TabStop = false;
			// 
			// CalendarForm
			// 
			this.AcceptButton = this.btnApply;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(229, 201);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnApply);
			this.Controls.Add(this.Calendar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "CalendarForm";
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Calendar";
			this.ResumeLayout(false);

		}
		#endregion
	}
}
