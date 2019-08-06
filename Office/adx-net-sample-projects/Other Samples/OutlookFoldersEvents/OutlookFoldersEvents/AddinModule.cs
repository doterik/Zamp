using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace OutlookFoldersEvents
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("EE7F063B-A576-4BA2-B10E-BD2B65DE423C"), ProgId("OutlookFoldersEvents.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		private AddinExpress.MSO.ADXOlExplorerCommandBar adxOlExplorerCommandBar1;
		private AddinExpress.MSO.ADXCommandBarComboBox adxCommandBarComboBox1;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton2;
		private AddinExpress.MSO.ADXRibbonTab adxRibbonTab1;
		private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
		private AddinExpress.MSO.ADXRibbonDropDown adxRibbonDropDown1;
		private AddinExpress.MSO.ADXRibbonBox adxRibbonBox1;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton1;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton2;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem1;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem2;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem3;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem4;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem5;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem6;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem7;
		private AddinExpress.MSO.ADXRibbonItem adxRibbonItem8;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton1;

		public AddinModule()
		{
			InitializeComponent();
		}

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
			this.adxCommandBarComboBox1 = new AddinExpress.MSO.ADXCommandBarComboBox(this.components);
			this.adxCommandBarButton1 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxCommandBarButton2 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
			this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
			this.adxRibbonDropDown1 = new AddinExpress.MSO.ADXRibbonDropDown(this.components);
			this.adxRibbonItem1 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem2 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem3 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem4 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem5 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem6 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem7 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonItem8 = new AddinExpress.MSO.ADXRibbonItem(this.components);
			this.adxRibbonBox1 = new AddinExpress.MSO.ADXRibbonBox(this.components);
			this.adxRibbonButton1 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonButton2 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			// 
			// adxOlExplorerCommandBar1
			// 
			this.adxOlExplorerCommandBar1.CommandBarName = "Outlook Folders Bar";
			this.adxOlExplorerCommandBar1.CommandBarTag = "c6c9aa26-48fa-49e0-b81d-001027cce250";
			this.adxOlExplorerCommandBar1.Controls.Add(this.adxCommandBarComboBox1);
			this.adxOlExplorerCommandBar1.Controls.Add(this.adxCommandBarButton1);
			this.adxOlExplorerCommandBar1.Controls.Add(this.adxCommandBarButton2);
			this.adxOlExplorerCommandBar1.Temporary = true;
			this.adxOlExplorerCommandBar1.UpdateCounter = 5;
			// 
			// adxCommandBarComboBox1
			// 
			this.adxCommandBarComboBox1.Caption = "Connect To";
			this.adxCommandBarComboBox1.ControlTag = "70caeebe-6e34-4b44-96cb-4229e2da3e75";
			this.adxCommandBarComboBox1.DropDownWidth = 150;
			this.adxCommandBarComboBox1.Items.AddRange(new string[] {
            "Deleted Items",
            "Drafts",
            "Inbox",
            "Outbox",
            "Sent Items",
            "Contacts",
            "Tasks",
            "Calendar"});
			this.adxCommandBarComboBox1.ListIndex = 3;
			this.adxCommandBarComboBox1.Style = AddinExpress.MSO.ADXMsoComboStyle.adxMsoComboLabel;
			this.adxCommandBarComboBox1.Temporary = true;
			this.adxCommandBarComboBox1.UpdateCounter = 7;
			this.adxCommandBarComboBox1.Change += new AddinExpress.MSO.ADXChange_EventHandler(this.adxCommandBarComboBox1_Change);
			// 
			// adxCommandBarButton1
			// 
			this.adxCommandBarButton1.BeginGroup = true;
			this.adxCommandBarButton1.Caption = "Disconnect";
			this.adxCommandBarButton1.ControlTag = "0dbb15c8-0e06-4754-af09-8d1995ecefda";
			this.adxCommandBarButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton1.Temporary = true;
			this.adxCommandBarButton1.UpdateCounter = 5;
			this.adxCommandBarButton1.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton1_Click);
			// 
			// adxCommandBarButton2
			// 
			this.adxCommandBarButton2.Caption = "Reconnect";
			this.adxCommandBarButton2.ControlTag = "3f6277c2-50a6-4d4c-a9a2-c03ea609efb8";
			this.adxCommandBarButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton2.Temporary = true;
			this.adxCommandBarButton2.UpdateCounter = 2;
			this.adxCommandBarButton2.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton2_Click);
			// 
			// adxRibbonTab1
			// 
			this.adxRibbonTab1.Caption = "Outlook Folder Items Events";
			this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
			this.adxRibbonTab1.Id = "adxRibbonTab_3751245a58444aa3b24bdc029eddb1af";
			this.adxRibbonTab1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonGroup1
			// 
			this.adxRibbonGroup1.Caption = "Ribbon Group";
			this.adxRibbonGroup1.Controls.Add(this.adxRibbonDropDown1);
			this.adxRibbonGroup1.Controls.Add(this.adxRibbonBox1);
			this.adxRibbonGroup1.Id = "adxRibbonGroup_bcaf1dfb60b240f9a32ec9719c5621fd";
			this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonDropDown1
			// 
			this.adxRibbonDropDown1.Caption = "Connect To";
			this.adxRibbonDropDown1.Id = "adxRibbonDropDown_2dee8f8bc48040e4bfd22966dbdfa2f7";
			this.adxRibbonDropDown1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem1);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem2);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem3);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem4);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem5);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem6);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem7);
			this.adxRibbonDropDown1.Items.Add(this.adxRibbonItem8);
			this.adxRibbonDropDown1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.adxRibbonDropDown1.OnAction += new AddinExpress.MSO.ADXRibbonOnActionSelected_EventHandler(this.adxRibbonDropDown1_OnAction);
			// 
			// adxRibbonItem1
			// 
			this.adxRibbonItem1.Caption = "Deleted Items";
			this.adxRibbonItem1.Id = "adxRibbonItem_4de9d06feafe4e28a46dffae59707f8b";
			this.adxRibbonItem1.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem2
			// 
			this.adxRibbonItem2.Caption = "Drafts";
			this.adxRibbonItem2.Id = "adxRibbonItem_431f27c95e4d44638827ecbd6acc4128";
			this.adxRibbonItem2.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem3
			// 
			this.adxRibbonItem3.Caption = "Inbox";
			this.adxRibbonItem3.Id = "adxRibbonItem_69cbbb1fd89a4175b516b1345a80dbb7";
			this.adxRibbonItem3.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem4
			// 
			this.adxRibbonItem4.Caption = "Outbox";
			this.adxRibbonItem4.Id = "adxRibbonItem_6752aaa56ee749aba8a6d4ce213c6d55";
			this.adxRibbonItem4.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem5
			// 
			this.adxRibbonItem5.Caption = "Sent Items";
			this.adxRibbonItem5.Id = "adxRibbonItem_4abc0594332b47dcbb1f2e42bc8983b2";
			this.adxRibbonItem5.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem6
			// 
			this.adxRibbonItem6.Caption = "Contacts";
			this.adxRibbonItem6.Id = "adxRibbonItem_8e6371826d2048dd94ae12fa5041d7c0";
			this.adxRibbonItem6.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem7
			// 
			this.adxRibbonItem7.Caption = "Tasks";
			this.adxRibbonItem7.Id = "adxRibbonItem_ce60fc3c1e6948fea79359bb1a2c5db0";
			this.adxRibbonItem7.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonItem8
			// 
			this.adxRibbonItem8.Caption = "Calendar";
			this.adxRibbonItem8.Id = "adxRibbonItem_9f4a5beca6cd4830a81dd7c92c8f38b2";
			this.adxRibbonItem8.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonBox1
			// 
			this.adxRibbonBox1.Controls.Add(this.adxRibbonButton1);
			this.adxRibbonBox1.Controls.Add(this.adxRibbonButton2);
			this.adxRibbonBox1.Id = "adxRibbonBox_2e8760c1eda7458caa399da1f31f4687";
			this.adxRibbonBox1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton1
			// 
			this.adxRibbonButton1.Caption = "Disconnect";
			this.adxRibbonButton1.Id = "adxRibbonButton_2ed5445e9b6240cbb13f7b11f6d801d3";
			this.adxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.adxRibbonButton1.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton1_OnClick);
			// 
			// adxRibbonButton2
			// 
			this.adxRibbonButton2.Caption = "Reconnect";
			this.adxRibbonButton2.Id = "adxRibbonButton_502118594adb49cc95998be1147f7ba9";
			this.adxRibbonButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.adxRibbonButton2.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton2_OnClick);
			// 
			// AddinModule
			// 
			this.AddinName = "Outlook Folders Events Example";
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.AddinBeginShutdown += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinBeginShutdown);
			this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);
			this.AddinStartupComplete += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinStartupComplete);

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

		private OutlookFoldersEvents eventObject = null;

		private void AddinModule_AddinInitialize(object sender, EventArgs e)
		{
			// Outlook 2010
			if (this.HostMajorVersion >= 14)
				adxOlExplorerCommandBar1.UseForRibbon = false;
		}

		private void AddinModule_AddinStartupComplete(object sender, System.EventArgs e)
		{
			eventObject = new OutlookFoldersEvents(this);
			if (adxCommandBarComboBox1.ListIndex > 0)
				ConnectToFolderByIndex(adxCommandBarComboBox1.ListIndex);
		}

		private void AddinModule_AddinBeginShutdown(object sender, System.EventArgs e)
		{
			if (eventObject != null)
				eventObject.Dispose();
		}

		private void adxCommandBarComboBox1_Change(object sender)
		{
			if (adxCommandBarComboBox1.ListIndex > 0)
				ConnectToFolderByIndex(adxCommandBarComboBox1.ListIndex);
		}

		private void adxCommandBarButton1_Click(object sender)
		{
			if (eventObject != null)
				eventObject.RemoveConnection();
		}

		private void adxCommandBarButton2_Click(object sender)
		{
			if (adxCommandBarComboBox1.ListIndex > 0)
				ConnectToFolderByIndex(adxCommandBarComboBox1.ListIndex);
		}

		private void adxRibbonDropDown1_OnAction(object sender, AddinExpress.MSO.IRibbonControl Control, string selectedId, int selectedIndex)
		{
			if (selectedIndex >= 0)
				ConnectToFolderByIndex(selectedIndex + 1);
		}

		private void adxRibbonButton1_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			if (eventObject != null)
				eventObject.RemoveConnection();
		}

		private void adxRibbonButton2_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			if (adxRibbonDropDown1.SelectedItemIndex >= 0)
				ConnectToFolderByIndex(adxRibbonDropDown1.SelectedItemIndex + 1);
		}

		private void ConnectToFolderByIndex(int index)
		{
			if (eventObject != null)
			{
				Outlook.OlDefaultFolders folderType = Outlook.OlDefaultFolders.olFolderInbox;
				switch (index)
				{
					case 1:
						folderType = Outlook.OlDefaultFolders.olFolderDeletedItems;
						break;
					case 2:
						folderType = Outlook.OlDefaultFolders.olFolderDrafts;
						break;
					case 3:
						folderType = Outlook.OlDefaultFolders.olFolderInbox;
						break;
					case 4:
						folderType = Outlook.OlDefaultFolders.olFolderOutbox;
						break;
					case 5:
						folderType = Outlook.OlDefaultFolders.olFolderSentMail;
						break;
					case 6:
						folderType = Outlook.OlDefaultFolders.olFolderContacts;
						break;
					case 7:
						folderType = Outlook.OlDefaultFolders.olFolderTasks;
						break;
					case 8:
						folderType = Outlook.OlDefaultFolders.olFolderCalendar;
						break;
				}
				Outlook._NameSpace ns = OutlookApp.GetNamespace("MAPI");
				if (ns != null)
					try
					{
						Outlook.MAPIFolder folder = ns.GetDefaultFolder(folderType);
						eventObject.ConnectTo(folder, true);
					}
					finally { Marshal.ReleaseComObject(ns); }
			}
		}
	}
}
