using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;

namespace OutlookItemEvents
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("2F1496A8-9363-454C-B251-6EEA053EE787"), ProgId("OutlookItemEvents.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			Application.EnableVisualStyles();
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXOutlookAppEvents adxOutlookEvents;

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
			this.adxOutlookEvents = new AddinExpress.MSO.ADXOutlookAppEvents(this.components);
			// 
			// adxOutlookEvents
			// 
			this.adxOutlookEvents.InspectorActivate += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_InspectorActivate);
			this.adxOutlookEvents.ItemSend += new AddinExpress.MSO.ADXOlItemSend_EventHandler(this.adxOutlookEvents_ItemSend);
			this.adxOutlookEvents.InspectorClose += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_InspectorClose);
			// 
			// AddinModule
			// 
			this.AddinName = "Outlook Item Events Example";
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
			this.AddinBeginShutdown += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinBeginShutdown);
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

		private ItemEventsClass itemEvents = null;

		private void AddinModule_AddinInitialize(object sender, EventArgs e)
		{
			itemEvents = new ItemEventsClass(this);
		}

		private void AddinModule_AddinBeginShutdown(object sender, System.EventArgs e)
		{
			if (itemEvents != null)
			{
				itemEvents.Dispose();
				itemEvents = null;
			}
		}

		private void adxOutlookEvents_InspectorActivate(object sender, object inspector, string folderName)
		{
			if (itemEvents != null)
			{
				Outlook._Inspector olInsp = inspector as Outlook._Inspector;
				itemEvents.ConnectTo(olInsp.CurrentItem, true);
			}
		}

		private void adxOutlookEvents_InspectorClose(object sender, object inspector, string folderName)
		{
			if (itemEvents != null)
				itemEvents.RemoveConnection();
		}

		private void adxOutlookEvents_ItemSend(object sender, AddinExpress.MSO.ADXOlItemSendEventArgs e)
		{
			if (itemEvents != null)
				itemEvents.RemoveConnection();
		}
	}
}
