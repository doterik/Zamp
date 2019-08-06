using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;

namespace SendMail
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.RichTextBox txbMail;
		private System.Windows.Forms.TextBox txbTO;
		private System.Windows.Forms.TextBox txbCC;
		private System.Windows.Forms.TextBox txbBCC;
		private System.Windows.Forms.TextBox txbSubject;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ToolBar toolBar1;
		private System.Windows.Forms.ToolBarButton tbtnSend;
		private System.Windows.Forms.ToolBarButton tbtnMode;
		private System.Windows.Forms.Splitter splitter2;
		private AddinExpress.Outlook.SecurityManager securityManager1;
		private System.ComponentModel.IContainer components;

		public Form1()
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
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.panel1 = new System.Windows.Forms.Panel();
			this.splitter2 = new System.Windows.Forms.Splitter();
			this.toolBar1 = new System.Windows.Forms.ToolBar();
			this.tbtnSend = new System.Windows.Forms.ToolBarButton();
			this.tbtnMode = new System.Windows.Forms.ToolBarButton();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txbSubject = new System.Windows.Forms.TextBox();
			this.txbBCC = new System.Windows.Forms.TextBox();
			this.txbCC = new System.Windows.Forms.TextBox();
			this.txbTO = new System.Windows.Forms.TextBox();
			this.txbMail = new System.Windows.Forms.RichTextBox();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.securityManager1 = new AddinExpress.Outlook.SecurityManager();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.splitter2);
			this.panel1.Controls.Add(this.toolBar1);
			this.panel1.Controls.Add(this.label4);
			this.panel1.Controls.Add(this.label3);
			this.panel1.Controls.Add(this.label2);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Controls.Add(this.txbSubject);
			this.panel1.Controls.Add(this.txbBCC);
			this.panel1.Controls.Add(this.txbCC);
			this.panel1.Controls.Add(this.txbTO);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(538, 168);
			this.panel1.TabIndex = 0;
			// 
			// splitter2
			// 
			this.splitter2.BackColor = System.Drawing.SystemColors.ControlDark;
			this.splitter2.Dock = System.Windows.Forms.DockStyle.Top;
			this.splitter2.Enabled = false;
			this.splitter2.Location = new System.Drawing.Point(0, 42);
			this.splitter2.Name = "splitter2";
			this.splitter2.Size = new System.Drawing.Size(538, 1);
			this.splitter2.TabIndex = 13;
			this.splitter2.TabStop = false;
			// 
			// toolBar1
			// 
			this.toolBar1.Appearance = System.Windows.Forms.ToolBarAppearance.Flat;
			this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.tbtnSend,
            this.tbtnMode});
			this.toolBar1.DropDownArrows = true;
			this.toolBar1.ImageList = this.imageList1;
			this.toolBar1.Location = new System.Drawing.Point(0, 0);
			this.toolBar1.Name = "toolBar1";
			this.toolBar1.ShowToolTips = true;
			this.toolBar1.Size = new System.Drawing.Size(538, 42);
			this.toolBar1.TabIndex = 12;
			this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
			// 
			// tbtnSend
			// 
			this.tbtnSend.ImageIndex = 0;
			this.tbtnSend.Name = "tbtnSend";
			this.tbtnSend.Text = "Send";
			this.tbtnSend.ToolTipText = "Send mail";
			// 
			// tbtnMode
			// 
			this.tbtnMode.ImageIndex = 1;
			this.tbtnMode.Name = "tbtnMode";
			this.tbtnMode.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton;
			this.tbtnMode.Text = "Security";
			this.tbtnMode.ToolTipText = "Security mode";
			// 
			// imageList1
			// 
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Fuchsia;
			this.imageList1.Images.SetKeyName(0, "");
			this.imageList1.Images.SetKeyName(1, "");
			this.imageList1.Images.SetKeyName(2, "");
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 128);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(48, 23);
			this.label4.TabIndex = 10;
			this.label4.Text = "Subject:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 23);
			this.label3.TabIndex = 9;
			this.label3.Text = "Bcc:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 23);
			this.label2.TabIndex = 8;
			this.label2.Text = "Cc:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 23);
			this.label1.TabIndex = 7;
			this.label1.Text = "To:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txbSubject
			// 
			this.txbSubject.Location = new System.Drawing.Point(72, 128);
			this.txbSubject.Name = "txbSubject";
			this.txbSubject.Size = new System.Drawing.Size(448, 20);
			this.txbSubject.TabIndex = 6;
			// 
			// txbBCC
			// 
			this.txbBCC.Location = new System.Drawing.Point(72, 104);
			this.txbBCC.Name = "txbBCC";
			this.txbBCC.Size = new System.Drawing.Size(448, 20);
			this.txbBCC.TabIndex = 5;
			// 
			// txbCC
			// 
			this.txbCC.Location = new System.Drawing.Point(72, 80);
			this.txbCC.Name = "txbCC";
			this.txbCC.Size = new System.Drawing.Size(448, 20);
			this.txbCC.TabIndex = 4;
			// 
			// txbTO
			// 
			this.txbTO.Location = new System.Drawing.Point(72, 56);
			this.txbTO.Name = "txbTO";
			this.txbTO.Size = new System.Drawing.Size(448, 20);
			this.txbTO.TabIndex = 3;
			// 
			// txbMail
			// 
			this.txbMail.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txbMail.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.txbMail.Location = new System.Drawing.Point(0, 168);
			this.txbMail.Name = "txbMail";
			this.txbMail.Size = new System.Drawing.Size(538, 216);
			this.txbMail.TabIndex = 1;
			this.txbMail.Text = "";
			// 
			// splitter1
			// 
			this.splitter1.BackColor = System.Drawing.SystemColors.ControlDark;
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.splitter1.Enabled = false;
			this.splitter1.Location = new System.Drawing.Point(0, 167);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(538, 1);
			this.splitter1.TabIndex = 2;
			this.splitter1.TabStop = false;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(538, 384);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.txbMail);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "Form1";
			this.Text = "Send Mail";
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.Run(new Form1());
		}

		private void toolBar1_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.Text)
			{
				case "Send":
					Outlook._Application outlookApp = null;
					Outlook._MailItem newMail = null;
					Outlook.Recipient recipient = null;
					Outlook.Recipients recipients = null;
					Type olType = Type.GetTypeFromProgID("Outlook.Application", false);
					if (olType != null)
					{
						try
						{
							outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook._Application;
						}
						catch { }
						try
						{
							if (outlookApp == null)
								outlookApp = Activator.CreateInstance(olType) as Outlook._Application;
						}
						catch { }
						if (outlookApp != null)
						{
							if (tbtnMode.Pushed)
							{
								securityManager1.ConnectTo(outlookApp);
								securityManager1.DisableOOMWarnings = true;
							}
							try
							{
								newMail = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
								if (newMail != null)
								{
									try
									{
										recipients = newMail.Recipients;
										if (recipients != null)
											try
											{
												if (txbTO.Text != string.Empty)
												{
													recipient = recipients.Add(txbTO.Text);
													if (recipient != null)
														try
														{
															recipient.Type = (int)Outlook.OlMailRecipientType.olTo;
															recipient.Resolve();
														}
														finally { Marshal.ReleaseComObject(recipient); }
												}
												if (txbCC.Text != string.Empty)
												{
													recipient = recipients.Add(txbCC.Text);
													if (recipient != null)
														try
														{
															recipient.Type = (int)Outlook.OlMailRecipientType.olCC;
															recipient.Resolve();
														}
														finally { Marshal.ReleaseComObject(recipient); }
												}
												if (txbBCC.Text != string.Empty)
												{
													recipient = recipients.Add(txbBCC.Text);
													if (recipient != null)
														try
														{
															recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
															recipient.Resolve();
														}
														finally { Marshal.ReleaseComObject(recipient); }
												}
												newMail.Subject = txbSubject.Text;
												newMail.Body = txbMail.Text;
												newMail.Send();
											}
											finally { Marshal.ReleaseComObject(recipients); }

										Marshal.ReleaseComObject(newMail);
										newMail = null;

										MessageBox.Show("The message has been sent successfully.", "Send Mail", MessageBoxButtons.OK, MessageBoxIcon.Information);
									}
									catch (Exception err)
									{
										if (newMail != null)
											Marshal.ReleaseComObject(newMail);
										MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
								}
							}
							finally
							{
								if (tbtnMode.Pushed)
								{
									securityManager1.DisableOOMWarnings = false;
									securityManager1.Disconnect(outlookApp);
								}
								if (outlookApp != null)
									Marshal.ReleaseComObject(outlookApp);
							}
						}
					}
					break;
				case "Security":
					if (tbtnMode.Pushed)
						tbtnMode.ImageIndex = 2;
					else
						tbtnMode.ImageIndex = 1;
					break;
			}
		}
	}
}
