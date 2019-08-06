using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace MySmartTag1
{
	/// <summary>
	///   Add-in Express Smart Tag Module
	/// </summary>
	[GuidAttribute("6054AFCA-F714-4CE9-8B5D-2935CBA42740"), ProgId("MySmartTag1.SmartTagModule")]
	public class SmartTagModule : AddinExpress.SmartTag.ADXSmartTagModule
	{
		public SmartTagModule()
		{
			InitializeComponent();
		}

		private AddinExpress.SmartTag.ADXSmartTag adxSmartTag1;
		private AddinExpress.SmartTag.ADXSmartTagAction adxSmartTagAction1;

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
			this.adxSmartTag1 = new AddinExpress.SmartTag.ADXSmartTag(this.components);
			this.adxSmartTagAction1 = new AddinExpress.SmartTag.ADXSmartTagAction(this.components);
			// 
			// adxSmartTag1
			// 
			this.adxSmartTag1.Actions.Add(this.adxSmartTagAction1);
			this.adxSmartTag1.ADXTag = "MySmartTag1";
			this.adxSmartTag1.RecognizedWords.AddRange(new string[] {
            "ADX Smart Tag"});
			this.adxSmartTag1.Caption = new AddinExpress.SmartTag.ADXLocalizedList();
			this.adxSmartTag1.Caption.Add(0, "My Smart Tag 1");
			// 
			// adxSmartTagAction1
			// 
			this.adxSmartTagAction1.ADXTag = "MySmartTagAction1";
			this.adxSmartTagAction1.Click += new AddinExpress.SmartTag.ADXSmartTagAction_EventHandler(this.adxSmartTagAction1_Click);
			this.adxSmartTagAction1.Caption = new AddinExpress.SmartTag.ADXLocalizedList();
			this.adxSmartTagAction1.Caption.Add(0, "My Action 1");
			// 
			// SmartTagModule
			// 
			this.NamespaceURI = "mysmarttag1/smarttagmodule";
			this.LibraryName = new AddinExpress.SmartTag.ADXLocalizedList();
			this.LibraryName.Add(0, "My Smart Tag Library");
			this.Description = new AddinExpress.SmartTag.ADXLocalizedList();
			this.Description.Add(0, "My Smart Tags");

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
		public static void SmartTagRegister(Type t)
		{
			AddinExpress.SmartTag.ADXSmartTagModule.ADXSmartTagRegister(t, typeof(MySmartTag1.SmartTagRecognizerImpl), typeof(MySmartTag1.SmartTagActionImpl));
		}

		[ComUnregisterFunctionAttribute]
		public static void SmartTagUnregister(Type t)
		{
			AddinExpress.SmartTag.ADXSmartTagModule.ADXSmartTagUnregister(t, typeof(MySmartTag1.SmartTagRecognizerImpl), typeof(MySmartTag1.SmartTagActionImpl));
		}

		#endregion

		private void adxSmartTagAction1_Click(object sender, AddinExpress.SmartTag.ADXSmartTagActionEventArgs e)
		{
			System.Windows.Forms.MessageBox.Show("Recognized text is '" + e.Text + "'!");
		}
	}
}
