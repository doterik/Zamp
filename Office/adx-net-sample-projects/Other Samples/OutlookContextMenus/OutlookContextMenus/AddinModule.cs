using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace OutlookContextMenus
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("51E44FAF-07A1-461B-B54A-BFB8A0E35143"), ProgId("OutlookContextMenus.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXContextMenu ItemsContextMenu;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton1;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton2;
		private AddinExpress.MSO.ADXCommandBarPopup adxCommandBarPopup1;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton3;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton4;
		private AddinExpress.MSO.ADXContextMenu FoldersContextMenu;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton5;
		private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton6;
		private AddinExpress.MSO.ADXRibbonContextMenu ItemsRibbonContextMenu;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton1;
		private AddinExpress.MSO.ADXRibbonSplitButton adxRibbonSplitButton1;
		private AddinExpress.MSO.ADXRibbonMenu adxRibbonMenu1;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton2;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton3;
		private AddinExpress.MSO.ADXRibbonMenuSeparator adxRibbonMenuSeparator1;
		private AddinExpress.MSO.ADXRibbonCheckBox adxRibbonCheckBox1;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton4;
		private AddinExpress.MSO.ADXRibbonContextMenu FoldersRibbonContextMenu;
		private System.Windows.Forms.ImageList imageList1616;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton5;
		private AddinExpress.MSO.ADXRibbonSplitButton adxRibbonSplitButton2;
		private AddinExpress.MSO.ADXRibbonMenu adxRibbonMenu2;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton6;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton7;
		private AddinExpress.MSO.ADXRibbonMenuSeparator adxRibbonMenuSeparator2;
		private AddinExpress.MSO.ADXRibbonCheckBox adxRibbonCheckBox2;
		private AddinExpress.MSO.ADXRibbonButton adxRibbonButton8;
		private AddinExpress.MSO.ADXRibbonMenuSeparator adxRibbonMenuSeparator3;
		private AddinExpress.MSO.ADXRibbonMenuSeparator adxRibbonMenuSeparator4;

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
			this.ItemsContextMenu = new AddinExpress.MSO.ADXContextMenu(this.components);
			this.adxCommandBarButton1 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxCommandBarButton2 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxCommandBarPopup1 = new AddinExpress.MSO.ADXCommandBarPopup(this.components);
			this.adxCommandBarButton3 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxCommandBarButton4 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.FoldersContextMenu = new AddinExpress.MSO.ADXContextMenu(this.components);
			this.adxCommandBarButton5 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxCommandBarButton6 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.ItemsRibbonContextMenu = new AddinExpress.MSO.ADXRibbonContextMenu(this.components);
			this.adxRibbonButton1 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonSplitButton1 = new AddinExpress.MSO.ADXRibbonSplitButton(this.components);
			this.adxRibbonMenu1 = new AddinExpress.MSO.ADXRibbonMenu(this.components);
			this.adxRibbonButton2 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonButton3 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonMenuSeparator1 = new AddinExpress.MSO.ADXRibbonMenuSeparator(this.components);
			this.adxRibbonCheckBox1 = new AddinExpress.MSO.ADXRibbonCheckBox(this.components);
			this.adxRibbonButton4 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.FoldersRibbonContextMenu = new AddinExpress.MSO.ADXRibbonContextMenu(this.components);
			this.adxRibbonButton5 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonSplitButton2 = new AddinExpress.MSO.ADXRibbonSplitButton(this.components);
			this.adxRibbonMenu2 = new AddinExpress.MSO.ADXRibbonMenu(this.components);
			this.adxRibbonButton6 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonButton7 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.adxRibbonMenuSeparator2 = new AddinExpress.MSO.ADXRibbonMenuSeparator(this.components);
			this.adxRibbonCheckBox2 = new AddinExpress.MSO.ADXRibbonCheckBox(this.components);
			this.adxRibbonButton8 = new AddinExpress.MSO.ADXRibbonButton(this.components);
			this.imageList1616 = new System.Windows.Forms.ImageList(this.components);
			this.adxRibbonMenuSeparator3 = new AddinExpress.MSO.ADXRibbonMenuSeparator(this.components);
			this.adxRibbonMenuSeparator4 = new AddinExpress.MSO.ADXRibbonMenuSeparator(this.components);
			// 
			// ItemsContextMenu
			// 
			this.ItemsContextMenu.CommandBarName = "Context Menu";
			this.ItemsContextMenu.CommandBarTag = "460b64c6-ffc3-4a63-9618-1cec5377eddc";
			this.ItemsContextMenu.Controls.Add(this.adxCommandBarButton1);
			this.ItemsContextMenu.Controls.Add(this.adxCommandBarButton2);
			this.ItemsContextMenu.Controls.Add(this.adxCommandBarPopup1);
			this.ItemsContextMenu.SupportedApp = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.ItemsContextMenu.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.ItemsContextMenu.Temporary = true;
			this.ItemsContextMenu.UpdateCounter = 4;
			// 
			// adxCommandBarButton1
			// 
			this.adxCommandBarButton1.BeginGroup = true;
			this.adxCommandBarButton1.Caption = "Command Bar Button1";
			this.adxCommandBarButton1.ControlTag = "32ce3d11-41c4-4086-9fd0-45d35625eeee";
			this.adxCommandBarButton1.Image = 0;
			this.adxCommandBarButton1.ImageList = this.imageList1616;
			this.adxCommandBarButton1.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.adxCommandBarButton1.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
			this.adxCommandBarButton1.Temporary = true;
			this.adxCommandBarButton1.UpdateCounter = 9;
			this.adxCommandBarButton1.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.Button_Click);
			// 
			// adxCommandBarButton2
			// 
			this.adxCommandBarButton2.Caption = "Command Bar Button2";
			this.adxCommandBarButton2.ControlTag = "f7b53240-0bc5-4d08-936a-6f0c26da56f1";
			this.adxCommandBarButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton2.Temporary = true;
			this.adxCommandBarButton2.UpdateCounter = 1;
			// 
			// adxCommandBarPopup1
			// 
			this.adxCommandBarPopup1.Caption = "Command Bar Popup1";
			this.adxCommandBarPopup1.Controls.Add(this.adxCommandBarButton3);
			this.adxCommandBarPopup1.Controls.Add(this.adxCommandBarButton4);
			this.adxCommandBarPopup1.ControlTag = "eed5c104-8fab-4ff8-8680-4adf540e32b6";
			this.adxCommandBarPopup1.Temporary = true;
			this.adxCommandBarPopup1.UpdateCounter = 1;
			// 
			// adxCommandBarButton3
			// 
			this.adxCommandBarButton3.Caption = "Command Bar Button3";
			this.adxCommandBarButton3.ControlTag = "0347549b-6a83-4619-baf3-ae6f8a8ada95";
			this.adxCommandBarButton3.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton3.Temporary = true;
			this.adxCommandBarButton3.UpdateCounter = 1;
			// 
			// adxCommandBarButton4
			// 
			this.adxCommandBarButton4.Caption = "Command Bar Button4";
			this.adxCommandBarButton4.ControlTag = "49a7fa33-7f5a-47a0-81f6-5b01560b6313";
			this.adxCommandBarButton4.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton4.Temporary = true;
			this.adxCommandBarButton4.UpdateCounter = 1;
			// 
			// FoldersContextMenu
			// 
			this.FoldersContextMenu.CommandBarName = "Folder Context Menu";
			this.FoldersContextMenu.CommandBarTag = "8c09195d-985b-4f59-8e28-e81b7b0da8c2";
			this.FoldersContextMenu.Controls.Add(this.adxCommandBarButton5);
			this.FoldersContextMenu.Controls.Add(this.adxCommandBarButton6);
			this.FoldersContextMenu.SupportedApp = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.FoldersContextMenu.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.FoldersContextMenu.Temporary = true;
			this.FoldersContextMenu.UpdateCounter = 5;
			// 
			// adxCommandBarButton5
			// 
			this.adxCommandBarButton5.BeginGroup = true;
			this.adxCommandBarButton5.Caption = "Command Bar Button1";
			this.adxCommandBarButton5.ControlTag = "1ce01f9a-e8aa-45ee-866f-b99de3b9e655";
			this.adxCommandBarButton5.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxCommandBarButton5.Temporary = true;
			this.adxCommandBarButton5.UpdateCounter = 4;
			this.adxCommandBarButton5.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.Button_Click);
			// 
			// adxCommandBarButton6
			// 
			this.adxCommandBarButton6.Caption = "Command Bar Button2";
			this.adxCommandBarButton6.ControlTag = "ef5c3894-16ac-41c3-b9b0-3d74757b16c8";
			this.adxCommandBarButton6.Image = 0;
			this.adxCommandBarButton6.ImageList = this.imageList1616;
			this.adxCommandBarButton6.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.adxCommandBarButton6.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
			this.adxCommandBarButton6.Temporary = true;
			this.adxCommandBarButton6.UpdateCounter = 5;
			// 
			// ItemsRibbonContextMenu
			// 
			this.ItemsRibbonContextMenu.ContextMenuNames.AddRange(new string[] {
            "Outlook.Explorer.ContextMenuMailItem",
            "Outlook.Explorer.ContextMenuContactItem",
            "Outlook.Explorer.ContextMenuJournalItem",
            "Outlook.Explorer.ContextMenuNoteItem",
            "Outlook.Explorer.ContextMenuTaskItem",
            "Outlook.Explorer.ContextMenuCalendarItem"});
			this.ItemsRibbonContextMenu.Controls.Add(this.adxRibbonMenuSeparator3);
			this.ItemsRibbonContextMenu.Controls.Add(this.adxRibbonButton1);
			this.ItemsRibbonContextMenu.Controls.Add(this.adxRibbonSplitButton1);
			this.ItemsRibbonContextMenu.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton1
			// 
			this.adxRibbonButton1.Caption = "Ribbon Button1";
			this.adxRibbonButton1.Id = "adxRibbonButton_a164c0251225431bafdbbec0eb3ed971";
			this.adxRibbonButton1.Image = 0;
			this.adxRibbonButton1.ImageList = this.imageList1616;
			this.adxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.adxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.adxRibbonButton1.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibbonButton_Click);
			// 
			// adxRibbonSplitButton1
			// 
			this.adxRibbonSplitButton1.Caption = "Ribbon SplitButton1";
			this.adxRibbonSplitButton1.Controls.Add(this.adxRibbonMenu1);
			this.adxRibbonSplitButton1.Id = "adxRibbonSplitButton_3685d46fb3cc4a10828803cc4b5be94c";
			this.adxRibbonSplitButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonSplitButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonMenu1
			// 
			this.adxRibbonMenu1.Caption = "Ribbon Menu1";
			this.adxRibbonMenu1.Controls.Add(this.adxRibbonButton2);
			this.adxRibbonMenu1.Controls.Add(this.adxRibbonButton3);
			this.adxRibbonMenu1.Controls.Add(this.adxRibbonMenuSeparator1);
			this.adxRibbonMenu1.Controls.Add(this.adxRibbonCheckBox1);
			this.adxRibbonMenu1.Controls.Add(this.adxRibbonButton4);
			this.adxRibbonMenu1.Id = "adxRibbonMenu_7d46d8c0df4d418da50600dbf080817b";
			this.adxRibbonMenu1.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonMenu1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton2
			// 
			this.adxRibbonButton2.Caption = "Ribbon Button2";
			this.adxRibbonButton2.Id = "adxRibbonButton_49a6608675f74afa9769b2e09446ba52";
			this.adxRibbonButton2.ImageMso = "HappyFace";
			this.adxRibbonButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton3
			// 
			this.adxRibbonButton3.Caption = "Ribbon Button3";
			this.adxRibbonButton3.Id = "adxRibbonButton_f48a3e9d04f94f99982cdaf947fe4946";
			this.adxRibbonButton3.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton3.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonMenuSeparator1
			// 
			this.adxRibbonMenuSeparator1.Caption = "adxRibbonMenuSeparator2";
			this.adxRibbonMenuSeparator1.Id = "adxRibbonMenuSeparator_69d117d877ee4f688fb47f98c89a4ec9";
			this.adxRibbonMenuSeparator1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonCheckBox1
			// 
			this.adxRibbonCheckBox1.Caption = "Ribbon CheckBox1";
			this.adxRibbonCheckBox1.Id = "adxRibbonCheckBox_9d6f3b0978cf41b1be296dc29db9d742";
			this.adxRibbonCheckBox1.Pressed = true;
			this.adxRibbonCheckBox1.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton4
			// 
			this.adxRibbonButton4.Caption = "Ribbon Button4";
			this.adxRibbonButton4.Id = "adxRibbonButton_b8796dda70e44061974c966c8495ed6d";
			this.adxRibbonButton4.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton4.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// FoldersRibbonContextMenu
			// 
			this.FoldersRibbonContextMenu.ContextMenuNames.AddRange(new string[] {
            "Outlook.Explorer.ContextMenuFolder",
            "Outlook.Explorer.ContextMenuStore"});
			this.FoldersRibbonContextMenu.Controls.Add(this.adxRibbonMenuSeparator4);
			this.FoldersRibbonContextMenu.Controls.Add(this.adxRibbonButton5);
			this.FoldersRibbonContextMenu.Controls.Add(this.adxRibbonSplitButton2);
			this.FoldersRibbonContextMenu.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton5
			// 
			this.adxRibbonButton5.Caption = "Ribbon Button1";
			this.adxRibbonButton5.Id = "adxRibbonButton_72837d253fe74cf1ba5ab23f0c39ad3d";
			this.adxRibbonButton5.Image = 0;
			this.adxRibbonButton5.ImageList = this.imageList1616;
			this.adxRibbonButton5.ImageTransparentColor = System.Drawing.Color.Fuchsia;
			this.adxRibbonButton5.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			this.adxRibbonButton5.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.RibbonButton_Click);
			// 
			// adxRibbonSplitButton2
			// 
			this.adxRibbonSplitButton2.Caption = "Ribbon SplitButton1";
			this.adxRibbonSplitButton2.Controls.Add(this.adxRibbonMenu2);
			this.adxRibbonSplitButton2.Id = "adxRibbonSplitButton_a85593deeea34a089518421596a13b62";
			this.adxRibbonSplitButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonSplitButton2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonMenu2
			// 
			this.adxRibbonMenu2.Caption = "Ribbon Menu1";
			this.adxRibbonMenu2.Controls.Add(this.adxRibbonButton6);
			this.adxRibbonMenu2.Controls.Add(this.adxRibbonButton7);
			this.adxRibbonMenu2.Controls.Add(this.adxRibbonMenuSeparator2);
			this.adxRibbonMenu2.Controls.Add(this.adxRibbonCheckBox2);
			this.adxRibbonMenu2.Controls.Add(this.adxRibbonButton8);
			this.adxRibbonMenu2.Id = "adxRibbonMenu_a9e5db0ce3084898aeba0b0e112e759b";
			this.adxRibbonMenu2.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonMenu2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton6
			// 
			this.adxRibbonButton6.Caption = "Ribbon Button2";
			this.adxRibbonButton6.Id = "adxRibbonButton_c4549866e70b4495b5eaa24772749bb8";
			this.adxRibbonButton6.ImageMso = "HappyFace";
			this.adxRibbonButton6.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton6.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton7
			// 
			this.adxRibbonButton7.Caption = "Ribbon Button3";
			this.adxRibbonButton7.Id = "adxRibbonButton_ad6163a51f20485c9e6ed870598e88ea";
			this.adxRibbonButton7.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton7.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonMenuSeparator2
			// 
			this.adxRibbonMenuSeparator2.Caption = "adxRibbonMenuSeparator2";
			this.adxRibbonMenuSeparator2.Id = "adxRibbonMenuSeparator_7dfa36a49d114ee88572e792e48a293f";
			this.adxRibbonMenuSeparator2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonCheckBox2
			// 
			this.adxRibbonCheckBox2.Caption = "Ribbon CheckBox1";
			this.adxRibbonCheckBox2.Id = "adxRibbonCheckBox_49dd1a043f5c4ad78a4b720fb068ec9a";
			this.adxRibbonCheckBox2.Pressed = true;
			this.adxRibbonCheckBox2.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonButton8
			// 
			this.adxRibbonButton8.Caption = "Ribbon Button4";
			this.adxRibbonButton8.Id = "adxRibbonButton_a59e30296e8641939e88edcfc72abb36";
			this.adxRibbonButton8.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButton8.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// imageList1616
			// 
			this.imageList1616.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1616.ImageStream")));
			this.imageList1616.TransparentColor = System.Drawing.Color.Transparent;
			this.imageList1616.Images.SetKeyName(0, "image.bmp");
			// 
			// adxRibbonMenuSeparator3
			// 
			this.adxRibbonMenuSeparator3.Caption = "adxRibbonMenuSeparator1";
			this.adxRibbonMenuSeparator3.Id = "adxRibbonMenuSeparator_6f023342dc844b5eb24c1a8ca7d4126a";
			this.adxRibbonMenuSeparator3.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// adxRibbonMenuSeparator4
			// 
			this.adxRibbonMenuSeparator4.Caption = "adxRibbonMenuSeparator1";
			this.adxRibbonMenuSeparator4.Id = "adxRibbonMenuSeparator_edd9b99d6b1c46d3aafd846c6f9f59ae";
			this.adxRibbonMenuSeparator4.Ribbons = AddinExpress.MSO.ADXRibbons.msrOutlookExplorer;
			// 
			// AddinModule
			// 
			this.AddinName = "Outlook Context Menu Example";
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
			{
				ItemsContextMenu.UseForRibbon = false;
				FoldersContextMenu.UseForRibbon = false;
			}
		}

		private void Button_Click(object sender)
		{
			System.Windows.Forms.MessageBox.Show("Click!");
		}

		private void RibbonButton_Click(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			System.Windows.Forms.MessageBox.Show("Click!");
		}
	}
}
