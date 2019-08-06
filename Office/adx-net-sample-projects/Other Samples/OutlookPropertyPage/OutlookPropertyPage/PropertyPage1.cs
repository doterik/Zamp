using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Proxies;

namespace OutlookPropertyPage
{
	/// <summary>
	/// Add-in Express Outlook Option Page
	/// </summary>
	[GuidAttribute("F9771703-1161-4570-BFAA-F9960B7BED3E"), ProgId("OutlookPropertyPage2005.PropertyPage1")]
	public class PropertyPage1 : AddinExpress.MSO.ADXOlPropertyPage
	{
		private Label label1;
		private TextBox textBox1;

		public PropertyPage1()
		{
			// This call is required by the Component Designer
			InitializeComponent();
		}

		#region Component Designer generated code
		/// <summary>
		/// Required by designer
		/// </summary>
		private System.ComponentModel.Container components = null;

		/// <summary>
		/// Required by designer - do not modify
		/// the following method
		/// </summary>
		private void InitializeComponent()
		{
			this.label1 = new System.Windows.Forms.Label();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(56, 57);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(190, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Caption of the Explorer button";
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(59, 73);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(232, 20);
			this.textBox1.TabIndex = 1;
			this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
			// 
			// PropertyPage1
			// 
			this.Controls.Add(this.textBox1);
			this.Controls.Add(this.label1);
			this.Name = "PropertyPage1";
			this.Size = new System.Drawing.Size(413, 358);
			this.Load += new System.EventHandler(this.PropertyPage1_Load);
			this.Apply += new System.EventHandler(this.PropertyPage1_Apply);
			this.Dirty += new AddinExpress.MSO.ADXDirty_EventHandler(this.PropertyPage1_Dirty);
			this.ResumeLayout(false);
			this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// Clean up any resources being used
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

		private bool loaded = false;

		private void PropertyPage1_Load(object sender, EventArgs e)
		{
			textBox1.Text = AddinModule.CurrentInstance.Caption;
			loaded = true;
		}

		private void PropertyPage1_Apply(object sender, EventArgs e)
		{
			AddinModule.CurrentInstance.Caption = textBox1.Text;
		}

		private void PropertyPage1_Dirty(object sender, AddinExpress.MSO.ADXDirtyEventArgs e)
		{
			e.Dirty = true;
		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{
			if (loaded)
				this.OnStatusChange();
		}
	}
}

