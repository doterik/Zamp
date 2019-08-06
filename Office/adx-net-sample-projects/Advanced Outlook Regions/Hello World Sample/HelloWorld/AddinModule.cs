using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;

namespace HelloWorld
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("F19A6CD0-DCAF-4DB5-B863-E546069BD257"), ProgId("HelloWorld.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			Application.EnableVisualStyles();
			InitializeComponent();
			// Please add any initialization code to the AddinInitialize event handler
		}

		private AddinExpress.OL.ADXOlFormsManager adxOlFormsManager1;
		private AddinExpress.OL.ADXOlFormsCollectionItem adxOlFormsCollectionItem1;

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
			this.adxOlFormsManager1 = new AddinExpress.OL.ADXOlFormsManager(this.components);
			this.adxOlFormsCollectionItem1 = new AddinExpress.OL.ADXOlFormsCollectionItem(this.components);
			// 
			// adxOlFormsManager1
			// 
			this.adxOlFormsManager1.Items.Add(this.adxOlFormsCollectionItem1);
			this.adxOlFormsManager1.SetOwner(this);
			// 
			// adxOlFormsCollectionItem1
			// 
			this.adxOlFormsCollectionItem1.AlwaysShowHeader = true;
			this.adxOlFormsCollectionItem1.CloseButton = true;
			this.adxOlFormsCollectionItem1.ExplorerAllowedDropRegions = ((AddinExpress.OL.ADXOlExplorerAllowedDropRegions)(((((((((((((((((AddinExpress.OL.ADXOlExplorerAllowedDropRegions.TopSubpane | AddinExpress.OL.ADXOlExplorerAllowedDropRegions.BottomSubpane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.RightSubpane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.LeftSubpane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.BottomOutlookBar)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.BottomNavigationPane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.BottomTodoBar)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.TopReadingPane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.BottomReadingPane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.LeftReadingPane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.RightReadingPane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.ReadingPane)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.FolderView)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.DockTop)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.DockBottom)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.DockRight)
						| AddinExpress.OL.ADXOlExplorerAllowedDropRegions.DockLeft)));
			this.adxOlFormsCollectionItem1.ExplorerItemTypes = ((AddinExpress.OL.ADXOlExplorerItemTypes)((((((((AddinExpress.OL.ADXOlExplorerItemTypes.olMailItem | AddinExpress.OL.ADXOlExplorerItemTypes.olAppointmentItem)
						| AddinExpress.OL.ADXOlExplorerItemTypes.olContactItem)
						| AddinExpress.OL.ADXOlExplorerItemTypes.olTaskItem)
						| AddinExpress.OL.ADXOlExplorerItemTypes.olJournalItem)
						| AddinExpress.OL.ADXOlExplorerItemTypes.olNoteItem)
						| AddinExpress.OL.ADXOlExplorerItemTypes.olPostItem)
						| AddinExpress.OL.ADXOlExplorerItemTypes.olDistributionListItem)));
			this.adxOlFormsCollectionItem1.ExplorerLayout = AddinExpress.OL.ADXOlExplorerLayout.RightSubpane;
			this.adxOlFormsCollectionItem1.FormClassName = "HelloWorld.ADXOlForm1";
			this.adxOlFormsCollectionItem1.InspectorAllowedDropRegions = ((AddinExpress.OL.ADXOlInspectorAllowedDropRegions)(((((AddinExpress.OL.ADXOlInspectorAllowedDropRegions.TopSubpane | AddinExpress.OL.ADXOlInspectorAllowedDropRegions.BottomSubpane)
						| AddinExpress.OL.ADXOlInspectorAllowedDropRegions.RightSubpane)
						| AddinExpress.OL.ADXOlInspectorAllowedDropRegions.LeftSubpane)
						| AddinExpress.OL.ADXOlInspectorAllowedDropRegions.InspectorRegion)));
			this.adxOlFormsCollectionItem1.InspectorItemTypes = AddinExpress.OL.ADXOlInspectorItemTypes.olMail;
			this.adxOlFormsCollectionItem1.InspectorLayout = AddinExpress.OL.ADXOlInspectorLayout.RightSubpane;
			this.adxOlFormsCollectionItem1.IsDragDropAllowed = true;
			this.adxOlFormsCollectionItem1.UseOfficeThemeForBackground = true;
			// 
			// AddinModule
			// 
			this.AddinName = "Hello World Sample";
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;

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
	}
}
