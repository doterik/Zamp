using System;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;

namespace OutlookAddin
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("37537AC4-FEE5-448d-A2FE-4B74B0110EDC"), ProgId("OutlookAddin.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		private System.Windows.Forms.ImageList imageList1;
		private AddinExpress.MSO.ADXOlExplorerCommandBar cmbSecurityManager;
		private AddinExpress.MSO.ADXCommandBarButton btnMode;
		private AddinExpress.MSO.ADXCommandBarButton btnContacts;
		private AddinExpress.MSO.ADXCommandBarButton btnMessage;
		private AddinExpress.Outlook.SecurityManager securityManager1;
		private AddinExpress.MSO.ADXRibbonTab RibbonTabSecurityManager;
		private AddinExpress.MSO.ADXRibbonGroup RibbonGroupSecurityManager;
		private AddinExpress.MSO.ADXRibbonButton RibbonButtonMode;
		private AddinExpress.MSO.ADXRibbonButton RibbonButtonContacts;
		private AddinExpress.MSO.ADXRibbonButton RibbonButtonMessage;
		private AddinExpress.MSO.ADXRibbonButtonGroup adxRibbonButtonGroup1;
		private AddinExpress.MSO.ADXOutlookAppEvents adxOutlookEvents;

		public AddinModule()
		{
			Application.EnableVisualStyles();
			InitializeComponent();
		}

		// Add-in Module Implementation

		#region Component Designer generated code
		/// <summary>
		/// Required by designer
		/// </summary>
		private System.ComponentModel.IContainer components;

		/// <summary>
		/// Required by designer support - do not modify
		/// the following method
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.cmbSecurityManager = new AddinExpress.MSO.ADXOlExplorerCommandBar(this.components);
			this.btnMode = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.btnContacts = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.btnMessage = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxOutlookEvents = new AddinExpress.MSO.ADXOutlookAppEvents(this.components);
			this.securityManager1 = new AddinExpress.Outlook.SecurityManager();
			this.RibbonTabSecurityManager = new AddinExpress.MSO.ADXRibbonTab(this.components);
			this.RibbonGroupSecurityManager = new AddinExpress.MSO.ADXRibbonGroup(this.components);
			this.RibbonButtonMode = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.RibbonButtonContacts = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.RibbonButtonMessage = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonButtonGroup1 = new AddinExpress.MSO.ADXRibbonButtonGroup(this.components);
			// 
			// imageList1
			// 
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			this.imageList1.Images.SetKeyName(0, "");
			this.imageList1.Images.SetKeyName(1, "");
			this.imageList1.Images.SetKeyName(2, "");
			this.imageList1.Images.SetKeyName(3, "");
			// 
			// cmbSecurityManager
			// 
			this.cmbSecurityManager.CommandBarName = "Security Manager .NET";
			this.cmbSecurityManager.CommandBarTag = "19CD1F5B-F746-4146-A614-459AAD87FB38";
			this.cmbSecurityManager.Controls.Add(this.btnMode);
			this.cmbSecurityManager.Controls.Add(this.btnContacts);
			this.cmbSecurityManager.Controls.Add(this.btnMessage);
			this.cmbSecurityManager.Temporary = true;
			this.cmbSecurityManager.UpdateCounter = 13;
			// 
			// btnMode
			// 
			this.btnMode.Caption = "Security (ON)";
			this.btnMode.ControlTag = "947ca615-7c35-4a08-a0e6-3814e12a7b94";
			this.btnMode.Image = 0;
			this.btnMode.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.btnMode.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
			this.btnMode.Tag = "10";
			this.btnMode.Temporary = true;
			this.btnMode.TooltipText = "Switch security";
			this.btnMode.UpdateCounter = 22;
			this.btnMode.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.CommonButtonClick);
			// 
			// btnContacts
			// 
			this.btnContacts.BeginGroup = true;
			this.btnContacts.Caption = "Enum Contacts";
			this.btnContacts.ControlTag = "82cd1e32-41e8-4358-a8fe-7ed9bc5ddbc3";
			this.btnContacts.Image = 2;
			this.btnContacts.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.btnContacts.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIcon;
			this.btnContacts.Tag = "20";
			this.btnContacts.Temporary = true;
			this.btnContacts.TooltipText = "Show Contacts";
			this.btnContacts.UpdateCounter = 23;
			this.btnContacts.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.CommonButtonClick);
			// 
			// btnMessage
			// 
			this.btnMessage.Caption = "Get Message Details";
			this.btnMessage.ControlTag = "a13aa433-87a9-4a12-9dcd-de1c52f3df75";
			this.btnMessage.Image = 3;
			this.btnMessage.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.btnMessage.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIcon;
			this.btnMessage.Tag = "30";
			this.btnMessage.Temporary = true;
			this.btnMessage.TooltipText = "Message details";
			this.btnMessage.UpdateCounter = 20;
			this.btnMessage.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.CommonButtonClick);
			// 
			// adxOutlookEvents
			// 
			this.adxOutlookEvents.ExplorerSelectionChange += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerSelectionChange);
			this.adxOutlookEvents.ExplorerFolderSwitch += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerFolderSwitch);
			// 
			// RibbonTabSecurityManager
			// 
			this.RibbonTabSecurityManager.Caption = "Security Manager .NET";
			this.RibbonTabSecurityManager.Controls.Add(this.RibbonGroupSecurityManager);
			this.RibbonTabSecurityManager.Id = "adxRibbonTab_b467dcde14a448aca4fcf7223c261918";
			this.RibbonTabSecurityManager.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// RibbonGroupSecurityManager
			// 
			this.RibbonGroupSecurityManager.Caption = "Security Manager .NET";
			this.RibbonGroupSecurityManager.Controls.Add(this.RibbonButtonMode);
			this.RibbonGroupSecurityManager.Controls.Add(this.adxRibbonButtonGroup1);
			this.RibbonGroupSecurityManager.Id = "adxRibbonGroup_b4b3e8c77aa84009aedd8b8c28cd1675";
			this.RibbonGroupSecurityManager.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.RibbonGroupSecurityManager.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// RibbonButtonMode
			// 
			this.RibbonButtonMode.Caption = "Security (ON)";
			this.RibbonButtonMode.Id = "adxRibbonButton_f63fa01bdd234417a3c712edceaff55e";
			this.RibbonButtonMode.Image = 0;
			this.RibbonButtonMode.ImageList = this.imageList1;
			this.RibbonButtonMode.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.RibbonButtonMode.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.RibbonButtonMode.ToggleButton = true;
			this.RibbonButtonMode.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibbonButtonMode_OnClick);
			// 
			// RibbonButtonContacts
			// 
			this.RibbonButtonContacts.Caption = "Show Contacts";
			this.RibbonButtonContacts.Id = "adxRibbonButton_3c514a06998b4ca89c5bbec1a2ea1da7";
			this.RibbonButtonContacts.Image = 2;
			this.RibbonButtonContacts.ImageList = this.imageList1;
			this.RibbonButtonContacts.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.RibbonButtonContacts.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.RibbonButtonContacts.ScreenTip = "Show Contacts";
			this.RibbonButtonContacts.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibbonButtonContacts_OnClick);
			// 
			// RibbonButtonMessage
			// 
			this.RibbonButtonMessage.Caption = "Message details";
			this.RibbonButtonMessage.Id = "adxRibbonButton_da6bc0612f0f41ddbf3e8fbc9b041e9c";
			this.RibbonButtonMessage.Image = 3;
			this.RibbonButtonMessage.ImageList = this.imageList1;
			this.RibbonButtonMessage.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.RibbonButtonMessage.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.RibbonButtonMessage.ScreenTip = "Message details";
			this.RibbonButtonMessage.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibbonButtonMessage_OnClick);
			// 
			// adxRibbonButtonGroup1
			// 
			this.adxRibbonButtonGroup1.Controls.Add(this.RibbonButtonContacts);
			this.adxRibbonButtonGroup1.Controls.Add(this.RibbonButtonMessage);
			this.adxRibbonButtonGroup1.Id = "adxRibbonButtonGroup_778acc6d2f4d4b41ba12bc3152d48502";
			this.adxRibbonButtonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// AddinModule
			// 
			this.AddinName = "Security Manager Sample (2005)";
			this.Images = this.imageList1;
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);

		}
		#endregion

		#region Add-in Express automatic code

		// Required by Add-in Express - do not modify
		// the methods within this region

		public override System.ComponentModel.IContainer GetContainer()
		{
			return components;
		}

		[ComRegisterFunctionAttribute]
		public static void AddinRegister(Type t)
		{
			AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
		}

		[ComUnregisterFunctionAttribute]
		public static void AddinUnregister(Type t)
		{
			AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
		}

		public override void UninstallControls()
		{
			base.UninstallControls();
		}

		#endregion

		public Outlook._Application OutlookApp
		{
			get
			{
				return (HostApplication as Outlook._Application);
			}
		}

		private void AddinModule_AddinInitialize(object sender, EventArgs e)
		{
			// Outlook 2010
			if (this.HostMajorVersion >= 14)
				cmbSecurityManager.UseForRibbon = false;
		}

		private void CommonButtonClick(object sender)
		{
			AddinExpress.MSO.ADXCommandBarButton button = sender as AddinExpress.MSO.ADXCommandBarButton;
			switch (button.Tag)
			{
				case "10":
					if (button.State == AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonUp)
					{
						button.State = AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonDown;
						button.Image = 1;
						button.Caption = "Security (OFF)";
					}
					else
					{
						button.State = AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonUp;
						button.Image = 0;
						button.Caption = "Security (ON)";
					}
					break;
				case "20":
					DoEnumContacts();
					break;
				case "30":
					DoGetInfo();
					break;
			}
		}

		private void RibbonButtonMode_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			if (RibbonButtonMode.Pressed)
			{
				btnMode.State = AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonDown;
				RibbonButtonMode.Image = 1;
				RibbonButtonMode.Caption = "Security (OFF)";
			}
			else
			{
				btnMode.State = AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonUp;
				RibbonButtonMode.Image = 0;
				RibbonButtonMode.Caption = "Security (ON)";
			}
		}

		private void RibbonButtonContacts_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			DoEnumContacts();
		}

		private void RibbonButtonMessage_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			DoGetInfo();
		}

		private void DoGetInfo()
		{
			if (btnMode.State == AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonDown)
				securityManager1.DisableOOMWarnings = true;

			try
			{
				Outlook._Explorer activeExplorer = OutlookApp.ActiveExplorer();
				if (activeExplorer != null)
					try
					{
						object selectedItem = null;
						Outlook.Selection selection = null;
						try
						{
							selection = activeExplorer.Selection;
							if (selection != null)
								try
								{
									if (selection.Count > 0)
										selectedItem = selection.Item(1);
								}
								finally { Marshal.ReleaseComObject(selection); }
						}
						catch { }
						if (selectedItem != null)
							try
							{
								if (selectedItem is Outlook.MailItem)
								{
									Outlook.MailItem mail = selectedItem as Outlook.MailItem;
									MessageForm frmMessage = new MessageForm();
									frmMessage.lbFrom.Text = mail.SenderName;
									frmMessage.lbSentOn.Text = mail.SentOn.ToString();
									frmMessage.lbTo.Text = mail.To;
									frmMessage.lbCC.Text = mail.CC;
									frmMessage.lbSubject.Text = mail.Subject;
									frmMessage.tbMessage.Text = mail.Body;
									frmMessage.ShowDialog();
									frmMessage.Dispose();
								}
							}
							finally { Marshal.ReleaseComObject(selectedItem); }
					}
					finally { Marshal.ReleaseComObject(activeExplorer); }
			}
			finally
			{
				if (btnMode.State == AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonDown)
					securityManager1.DisableOOMWarnings = false;
			}
		}

		private void DoEnumContacts()
		{
			if (btnMode.State == AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonDown)
				securityManager1.DisableOOMWarnings = true;

			try
			{
				Outlook.NameSpace namespace_ = OutlookApp.GetNamespace("MAPI");
				if (namespace_ != null)
					try
					{
						Outlook.MAPIFolder folder = namespace_.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
						if (folder != null)
							try
							{
								ContactsForm frmContacts = new ContactsForm();
								try
								{
									Outlook.Items items = folder.Items;
									if (items != null)
										try
										{
											for (int i = 1; i <= items.Count; i++)
											{
												Outlook.ContactItem contact = items.Item(i) as Outlook.ContactItem;
												if (contact != null)
													try
													{
														if (!string.IsNullOrEmpty(contact.Email1DisplayName))
														{
															DataRow newRow = frmContacts.dsContacts.Tables["Contacts"].NewRow();
															newRow["DisplayName"] = contact.Email1DisplayName;
															newRow["Address"] = contact.Email1Address;
															newRow["AddressType"] = contact.Email1AddressType;
															frmContacts.dsContacts.Tables["Contacts"].Rows.Add(newRow);
														}
														if (!string.IsNullOrEmpty(contact.Email2DisplayName))
														{
															DataRow newRow = frmContacts.dsContacts.Tables["Contacts"].NewRow();
															newRow["DisplayName"] = contact.Email2DisplayName;
															newRow["Address"] = contact.Email2Address;
															newRow["AddressType"] = contact.Email2AddressType;
															frmContacts.dsContacts.Tables["Contacts"].Rows.Add(newRow);
														}
														if (!string.IsNullOrEmpty(contact.Email3DisplayName))
														{
															DataRow newRow = frmContacts.dsContacts.Tables["Contacts"].NewRow();
															newRow["DisplayName"] = contact.Email3DisplayName;
															newRow["Address"] = contact.Email3Address;
															newRow["AddressType"] = contact.Email3AddressType;
															frmContacts.dsContacts.Tables["Contacts"].Rows.Add(newRow);
														}
													}
													finally { Marshal.ReleaseComObject(contact); }
											}
										}
										finally { Marshal.ReleaseComObject(items); }
								}
								catch { }
								frmContacts.ShowDialog();
								frmContacts.Dispose();
							}
							finally { Marshal.ReleaseComObject(folder); }
					}
					finally { Marshal.ReleaseComObject(namespace_); }
			}
			finally
			{
				if (btnMode.State == AddinExpress.MSO.ADXMsoButtonState.adxMsoButtonDown)
					securityManager1.DisableOOMWarnings = false;
			}
		}

		private void adxOutlookEvents_ExplorerSelectionChange(object sender, object explorer)
		{
			int selectedItemClass = -1;
			Outlook._Explorer activeExplorer = explorer as Outlook._Explorer;
			if (activeExplorer != null)
			{
				try
				{
					Outlook.Selection selection = activeExplorer.Selection;
					if (selection != null)
						try
						{
							if (selection.Count > 0)
							{
								object selectedItem = selection.Item(1);
								if (selectedItem != null)
									try
									{
										selectedItemClass = Convert.ToInt32(selectedItem.GetType().InvokeMember("Class", BindingFlags.GetProperty, null, selectedItem, null));
									}
									finally { Marshal.ReleaseComObject(selectedItem); }
							}
						}
						finally { Marshal.ReleaseComObject(selection); }
				}
				catch { }
				btnMessage.Enabled = (selectedItemClass == (int)Outlook.OlObjectClass.olMail);
			}
		}

		private void adxOutlookEvents_ExplorerFolderSwitch(object sender, object explorer)
		{
			Outlook._Explorer activeExplorer = OutlookApp.ActiveExplorer();
			if (activeExplorer != null)
				try
				{
					Outlook.MAPIFolder folder = activeExplorer.CurrentFolder;
					if (folder != null)
						try
						{
							Outlook.Items items = folder.Items;
							if (items != null)
								try
								{
									btnMessage.Enabled = (items.Count != 0);
								}
								finally { Marshal.ReleaseComObject(items); }
						}
						finally { Marshal.ReleaseComObject(folder); }
				}
				finally { Marshal.ReleaseComObject(activeExplorer); }
		}
	}
}
