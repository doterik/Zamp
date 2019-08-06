using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace OutlookPropertyPage
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("4DDDBB23-ADCC-462E-A047-A3DA46030C60"), ProgId("OutlookPropertyPage.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			System.Windows.Forms.Application.EnableVisualStyles();
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXOlExplorerCommandBar adxOlExplorerCommandBar1;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton1;
		private AddinExpress.MSO.ADXRibbonTab adxRibbonTab1;
		private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton1;
		private AddinExpress.MSO.ADXOlFolderPage adxOlFolderPage1;

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
			this.adxOlExplorerCommandBar1 = new AddinExpress.MSO.ADXOlExplorerCommandBar(this.components);
			this.adxCommandBarButton1 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
			this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
			this.adxRibbonButton1 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxOlFolderPage1 = new AddinExpress.MSO.ADXOlFolderPage(this.components);
			// 
			// adxOlExplorerCommandBar1
			// 
			this.adxOlExplorerCommandBar1.CommandBarName = "Outlook Command Bar";
			this.adxOlExplorerCommandBar1.CommandBarTag = "4b1bbeb1-c133-4763-8dc5-9e0250661bd6";
			this.adxOlExplorerCommandBar1.Controls.Add(this.adxCommandBarButton1);
			this.adxOlExplorerCommandBar1.Temporary = true;
			this.adxOlExplorerCommandBar1.UpdateCounter = 4;
			// 
			// adxCommandBarButton1
			// 
			this.adxCommandBarButton1.Caption = "Default Caption";
			this.adxCommandBarButton1.ControlTag = "e379ad2f-0161-4be6-8552-fd3a31bdae25";
			this.adxCommandBarButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton1.Temporary = true;
			this.adxCommandBarButton1.UpdateCounter = 1;
			// 
			// adxRibbonTab1
			// 
			this.adxRibbonTab1.Caption = "Outlook Ribbon Tab";
			this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
			this.adxRibbonTab1.Id = "adxRibbonTab_4b98cc75d99a442093ba931e295075eb";
			this.adxRibbonTab1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonGroup1
			// 
			this.adxRibbonGroup1.Caption = "Ribbon Group";
			this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton1);
			this.adxRibbonGroup1.Id = "adxRibbonGroup_b55f44de4dd54e229da99ac832767972";
			this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton1
			// 
			this.adxRibbonButton1.Caption = "Default Caption";
			this.adxRibbonButton1.Id = "adxRibbonButton_7c1b07c0714c43d7aa3ee4d1298f035f";
			this.adxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxOlFolderPage1
			// 
			this.adxOlFolderPage1.FolderName = "ROOT_FOLDER_NAME\\Inbox";
			this.adxOlFolderPage1.PageControl = null;
			this.adxOlFolderPage1.PageTitle = "My Page";
			this.adxOlFolderPage1.PageType = "OutlookPropertyPage.PropertyPage1";
			// 
			// AddinModule
			// 
			this.AddinName = "OutlookPropertyPage";
			this.FolderPages.Add(this.adxOlFolderPage1);
			this.PageTitle = "My Page";
			this.PageType = "OutlookPropertyPage.PropertyPage1";
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);

		}
		#endregion

		#region Add-in Express automatic code

		// Required by Add-in Express - do not modify
		// the methods within this region

		public override System.ComponentModel.IContainer GetContainer()
		{
			if (components == null)
				components = new System.ComponentModel.Container();
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

		public static new AddinModule CurrentInstance
		{
			get
			{
				return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
			}
		}

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
				adxOlExplorerCommandBar1.UseForRibbon = false;

			// set the FolderName property
			Outlook._NameSpace ns = OutlookApp.GetNamespace("MAPI");
			if (ns != null)
				try
				{
					Outlook.MAPIFolder inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
					if (inbox != null)
						try
						{
							Outlook.MAPIFolder store = inbox.Parent as Outlook.MAPIFolder;
							if (store != null)
								try
								{
									this.adxOlFolderPage1.FolderName = store.Name + @"\" + inbox.Name;
								}
								finally { Marshal.ReleaseComObject(store); }
						}
						finally { Marshal.ReleaseComObject(inbox); }
				}
				finally { Marshal.ReleaseComObject(ns); }
		}

		public string Caption
		{
			get
			{
				return adxCommandBarButton1.Caption;
			}
			set
			{
				adxRibbonButton1.Caption = value;
				adxCommandBarButton1.Caption = value;
			}
		}
	}
}

