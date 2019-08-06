using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace ExcelTimesheet
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("AB313745-E493-4C89-A108-BD72E52AE8FB"), ProgId("ExcelTimesheet.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXAddinAdditionalModuleItem adxAddinAdditionalModuleItem1;
		private AddinExpress.MSO.ADXExcelAppEvents adxExcelEvents;

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
			this.adxExcelEvents = new AddinExpress.MSO.ADXExcelAppEvents(this.components);
			// 
			// adxAddinAdditionalModuleItem1
			// 
			this.adxAddinAdditionalModuleItem1.ModuleProgID = "ExcelTimesheet.TimeSheetModule";
			// 
			// adxExcelEvents
			// 
			this.adxExcelEvents.WorkbookDeactivate += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxExcelEvents_WorkbookDeactivate);
			this.adxExcelEvents.WindowActivate += new AddinExpress.MSO.ADXHostWindow_EventHandler(this.adxExcelEvents_WindowActivate);
			// 
			// AddinModule
			// 
			this.AddinName = "Excel Timesheet Example";
			this.Modules.Add(this.adxAddinAdditionalModuleItem1);
			this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
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

		public Excel._Application ExcelApp
		{
			get
			{
				return (HostApplication as Excel._Application);
			}
		}

		private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
		{
			ChangeButtonState(false);
		}

		private void adxExcelEvents_WindowActivate(object sender, object hostObj, object window)
		{
			TimesheetModule myModule = Modules["TimesheetModule"].Module as TimesheetModule;
			if (myModule != null)
				myModule.OnWindowActivate(hostObj);
		}

		private void adxExcelEvents_WorkbookDeactivate(object sender, object hostObj)
		{
			ChangeButtonState(false);
		}

		private void ChangeButtonState(bool enabled)
		{
			TimesheetModule myModule = Modules["TimesheetModule"].Module as TimesheetModule;
			if (myModule != null)
			{
				myModule.adxClearTimesheet.Enabled = enabled;
				myModule.adxRibbonButtonClearTimesheet.Enabled = enabled;
			}
		}
	}
}
