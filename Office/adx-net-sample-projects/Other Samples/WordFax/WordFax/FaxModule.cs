using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordFax
{
	/// <summary>
	///   Add-in Express Word Document Module
	/// </summary>
	[GuidAttribute("71465C99-8098-43DE-8D34-0FC4B4B216F0"), ProgId("WordFax.FaxModule")]
	public class FaxModule : AddinExpress.MSO.ADXWordDocumentModule
	{
		public FaxModule()
		{
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXMSFormsComboBox adxToBox;
		private AddinExpress.MSO.ADXMSFormsComboBox adxFromBox;
		private AddinExpress.MSO.ADXMSFormsLabelControl adxAddressLbl;
		private AddinExpress.MSO.ADXMSFormsLabelControl adxFaxLbl;
		private AddinExpress.MSO.ADXMSFormsLabelControl adxDateLbl;
		private AddinExpress.MSO.ADXMSFormsCommandButton adxDateBtn;
		private AddinExpress.MSO.ADXCommandBar adxWordFaxBar;
		internal AddinExpress.MSO.ADXCommandBarButton adxClearControlsBtn;
		private AddinExpress.MSO.ADXRibbonTab adxWordFaxRibbonTab;
		private AddinExpress.MSO.ADXRibbonGroup adxWordFaxRibbonGroup;
		internal AddinExpress.MSO.ADXRibbonButton adxClearControlsRibbonButton;

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
			this.adxToBox = new AddinExpress.MSO.ADXMSFormsComboBox(this.components);
			this.adxFromBox = new AddinExpress.MSO.ADXMSFormsComboBox(this.components);
			this.adxAddressLbl = new AddinExpress.MSO.ADXMSFormsLabelControl(this.components);
			this.adxFaxLbl = new AddinExpress.MSO.ADXMSFormsLabelControl(this.components);
			this.adxDateLbl = new AddinExpress.MSO.ADXMSFormsLabelControl(this.components);
			this.adxDateBtn = new AddinExpress.MSO.ADXMSFormsCommandButton(this.components);
			this.adxWordFaxBar = new AddinExpress.MSO.ADXCommandBar(this.components);
			this.adxClearControlsBtn = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxWordFaxRibbonTab = new AddinExpress.MSO.ADXRibbonTab(this.components);
			this.adxWordFaxRibbonGroup = new AddinExpress.MSO.ADXRibbonGroup(this.components);
			this.adxClearControlsRibbonButton = new AddinExpress.MSO.ADXRibbonButton(this.components);
			// 
			// adxToBox
			// 
			this.adxToBox.ControlName = "ToBox";
			this.adxToBox.Connect += new System.EventHandler(this.adxToBox_Connect);
			this.adxToBox.Click += new System.EventHandler(this.adxToBox_Click);
			// 
			// adxFromBox
			// 
			this.adxFromBox.ControlName = "FromBox";
			this.adxFromBox.Connect += new System.EventHandler(this.adxFromBox_Connect);
			this.adxFromBox.Click += new System.EventHandler(this.adxFromBox_Click);
			// 
			// adxAddressLbl
			// 
			this.adxAddressLbl.ControlName = "AddressLbl";
			// 
			// adxFaxLbl
			// 
			this.adxFaxLbl.ControlName = "FaxLbl";
			// 
			// adxDateLbl
			// 
			this.adxDateLbl.ControlName = "DateLbl";
			// 
			// adxDateBtn
			// 
			this.adxDateBtn.ControlName = "DateBtn";
			this.adxDateBtn.Click += new System.EventHandler(this.adxDateBtn_Click);
			// 
			// adxWordFaxBar
			// 
			this.adxWordFaxBar.CommandBarName = "Word Fax Bar";
			this.adxWordFaxBar.CommandBarTag = "6b99158d-946a-4951-a6d1-c258480ec852";
			this.adxWordFaxBar.Controls.Add(this.adxClearControlsBtn);
			this.adxWordFaxBar.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaWord;
			this.adxWordFaxBar.UpdateCounter = 13;
			// 
			// adxClearControlsBtn
			// 
			this.adxClearControlsBtn.Caption = "Clear Fields";
			this.adxClearControlsBtn.ControlTag = "9b414a31-610a-49a2-ae48-28ab4cafd0b3";
			this.adxClearControlsBtn.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxClearControlsBtn.UpdateCounter = 10;
			this.adxClearControlsBtn.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxShowControlsBtn_Click);
			// 
			// adxWordFaxRibbonTab
			// 
			this.adxWordFaxRibbonTab.Caption = "Word Fax Ribbon Tab";
			this.adxWordFaxRibbonTab.Controls.Add(this.adxWordFaxRibbonGroup);
			this.adxWordFaxRibbonTab.Id = "adxRibbonTab_438ab8b1cc304340b541318e45331623";
			this.adxWordFaxRibbonTab.Ribbons = AddinExpress.MSO.ADXRibbons.msrWordDocument;
			// 
			// adxWordFaxRibbonGroup
			// 
			this.adxWordFaxRibbonGroup.Caption = "Word Fax Ribbon Group";
			this.adxWordFaxRibbonGroup.Controls.Add(this.adxClearControlsRibbonButton);
			this.adxWordFaxRibbonGroup.Id = "adxRibbonGroup_00f6e98442dd4dd984f890c700e54eab";
			this.adxWordFaxRibbonGroup.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxWordFaxRibbonGroup.Ribbons = AddinExpress.MSO.ADXRibbons.msrWordDocument;
			// 
			// adxClearControlsRibbonButton
			// 
			this.adxClearControlsRibbonButton.Caption = "Clear Fields";
			this.adxClearControlsRibbonButton.Id = "adxRibbonButton_912c8304d0434302b949211bd3671c4d";
			this.adxClearControlsRibbonButton.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxClearControlsRibbonButton.Ribbons = AddinExpress.MSO.ADXRibbons.msrWordDocument;
			this.adxClearControlsRibbonButton.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxClearControlsRibbonButton_OnClick);
			// 
			// FaxModule
			// 
			this.Description = "ADX Word Fax Module Example";
			this.ModuleName = "FaxModule";
			this.PropertyId = "_ADX_WordFax";
			this.PropertyValue = "Created for the WordFax project.";

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

		#endregion

		private string[] Recipients = new string[]{ "Nelson, Roberto", "Young, Bruce", "Lambert, Kim" };
		private string[] Faxes = new string[]{ "808-555-0278", "809-555-4958", "357-6-870943" };

		private void adxToBox_Connect(object sender, System.EventArgs e)
		{
			try
			{
				AddinExpress.MSO.ADXMSFormsComboBox cmb = this.MSFControlByName("ToBox") as AddinExpress.MSO.ADXMSFormsComboBox;
				if (cmb.ListCount == 0)
				{
					cmb.AddItem(Recipients[0]);
					cmb.AddItem(Recipients[1]);
					cmb.AddItem(Recipients[2]);
				}
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void adxToBox_Click(object sender, System.EventArgs e)
		{
			adxFaxLbl.Value = Faxes[adxToBox.ListIndex];
		}

		private void adxFromBox_Connect(object sender, System.EventArgs e)
		{
			try
			{
				AddinExpress.MSO.ADXMSFormsComboBox cmb = this.MSFControlByName("FromBox") as AddinExpress.MSO.ADXMSFormsComboBox;
				if (cmb.ListCount == 0)
					cmb.AddItem("Fish Research Labs");
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void adxFromBox_Click(object sender, System.EventArgs e)
		{
			adxAddressLbl.Value = "Los Banos, 29 Wilkins Rd Dept. SD, 93635, U.S.A., Phone: 209-555-3292, Fax: 203-555-0416";
		}

		private void adxDateBtn_Click(object sender, System.EventArgs e)
		{
			CalendarForm f = new CalendarForm();
			if (f.ShowDialog() == DialogResult.OK)
			{
				adxDateLbl.Value = f.Calendar.SelectionStart.ToLongDateString();
			}
			f.Dispose();
		}

		private void adxShowControlsBtn_Click(object sender)
		{
			adxToBox.ListIndex = -1;
			adxFromBox.ListIndex = -1;
			adxAddressLbl.Value = string.Empty;
			adxFaxLbl.Value = string.Empty;
			adxDateLbl.Value = string.Empty;
		}

		private void adxClearControlsRibbonButton_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			adxToBox.ListIndex = -1;
			adxFromBox.ListIndex = -1;
			adxAddressLbl.Value = string.Empty;
			adxFaxLbl.Value = string.Empty;
			adxDateLbl.Value = string.Empty;
		}

		public void OnWindowActivate(object document)
		{
			if (IsConnected)
			{
				try
				{
					if (String.Compare((DocumentObj as Word._Document).FullName, (document as Word._Document).FullName, true) == 0)
					{
						adxClearControlsBtn.Enabled = true;
						adxClearControlsRibbonButton.Enabled = true;
					}
				}
				catch (Exception err)
				{
					MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}
	}
}
