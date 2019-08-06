using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;

namespace WordFax
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("0E7E20F7-E622-4BF3-B10B-75EC3C2FB256"), ProgId("WordFax.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			Application.EnableVisualStyles();
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXWordAppEvents adxWordEvents;
		private AddinExpress.MSO.ADXAddinAdditionalModuleItem adxAddinAdditionalModuleItem1;

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
			this.adxAddinAdditionalModuleItem1 = new AddinExpress.MSO.ADXAddinAdditionalModuleItem(this.components);
			this.adxWordEvents = new AddinExpress.MSO.ADXWordAppEvents(this.components);
			// 
			// adxAddinAdditionalModuleItem1
			// 
			this.adxAddinAdditionalModuleItem1.ModuleProgID = "WordFax.FaxModule";
			// 
			// adxWordEvents
			// 
			this.adxWordEvents.NewDocument += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxWordEvents_NewDocument);
			this.adxWordEvents.WindowActivate += new AddinExpress.MSO.ADXHostWindow_EventHandler(this.adxWordEvents_WindowActivate);
			this.adxWordEvents.WindowDeactivate += new AddinExpress.MSO.ADXHostWindow_EventHandler(this.adxWordEvents_WindowDeactivate);
			this.adxWordEvents.DocumentOpen += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxWordEvents_DocumentOpen);
			// 
			// AddinModule
			// 
			this.AddinName = "Word Fax Example";
			this.Modules.Add(this.adxAddinAdditionalModuleItem1);
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaWord;
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

		public Word._Application WordApp
		{
			get
			{
				return (HostApplication as Word._Application);
			}
		}

		private void ChangeButtonState(bool enabled)
		{
			FaxModule myModule = Modules["FaxModule"].Module as FaxModule;
			if (myModule != null)
			{
				myModule.adxClearControlsBtn.Enabled = enabled;
				myModule.adxClearControlsRibbonButton.Enabled = enabled;
			}
		}

		private void DoWindowActivate(object docObj)
		{
			FaxModule myModule = Modules["FaxModule"].Module as FaxModule;
			if (myModule != null)
				myModule.OnWindowActivate(docObj);
		}

		private void adxWordEvents_WindowActivate(object sender, object hostObj, object window)
		{
			DoWindowActivate(hostObj);
		}

		private void adxWordEvents_WindowDeactivate(object sender, object hostObj, object window)
		{
			ChangeButtonState(false);
		}

		private void AddinModule_AddinStartupComplete(object sender, System.EventArgs e)
		{
			ChangeButtonState(false);
		}

		private void adxWordEvents_NewDocument(object sender, object hostObj)
		{
			DoWindowActivate(hostObj);
		}

		private void adxWordEvents_DocumentOpen(object sender, object hostObj)
		{
			DoWindowActivate(hostObj);
		}
	}
}
