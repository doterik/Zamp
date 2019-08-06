using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace RTDServerSample
{
	/// <summary>
	///   Add-in Express RTD Server Module
	/// </summary>
	[GuidAttribute("C64F1125-6D7F-4426-ADAB-9794BE5AE6F2"), ProgId("RTDServerSample.RTDServerModule")]
	public class RTDServerModule : AddinExpress.RTD.ADXRTDServerModule
	{
		public RTDServerModule()
		{
			InitializeComponent();
		}

		private AddinExpress.RTD.ADXRTDTopic TopicMSFTOpen;
		private AddinExpress.RTD.ADXRTDTopic TopicWMTOpen;
		private AddinExpress.RTD.ADXRTDTopic TopicATTOpen;
		private AddinExpress.RTD.ADXRTDTopic TopicINTCOpen;
		private AddinExpress.RTD.ADXRTDTopic TopicMSFTLast;
		private AddinExpress.RTD.ADXRTDTopic TopicWMTLast;
		private AddinExpress.RTD.ADXRTDTopic TopicATTLast;
		private AddinExpress.RTD.ADXRTDTopic TopicINTCLast;

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
            this.TopicMSFTOpen = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicWMTOpen = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicATTOpen = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicINTCOpen = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicMSFTLast = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicWMTLast = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicATTLast = new AddinExpress.RTD.ADXRTDTopic(this.components);
            this.TopicINTCLast = new AddinExpress.RTD.ADXRTDTopic(this.components);
            // 
            // TopicMSFTOpen
            // 
            this.TopicMSFTOpen.Enabled = false;
            this.TopicMSFTOpen.String01 = "MSFT";
            this.TopicMSFTOpen.String02 = "Open";
            this.TopicMSFTOpen.Tag = "";
            // 
            // TopicWMTOpen
            // 
            this.TopicWMTOpen.Enabled = false;
            this.TopicWMTOpen.String01 = "WMT";
            this.TopicWMTOpen.String02 = "Open";
            this.TopicWMTOpen.Tag = "";
            // 
            // TopicATTOpen
            // 
            this.TopicATTOpen.Enabled = false;
            this.TopicATTOpen.String01 = "ATT";
            this.TopicATTOpen.String02 = "Open";
            this.TopicATTOpen.Tag = "";
            // 
            // TopicINTCOpen
            // 
            this.TopicINTCOpen.Enabled = false;
            this.TopicINTCOpen.String01 = "INTC";
            this.TopicINTCOpen.String02 = "Open";
            this.TopicINTCOpen.Tag = "";
            // 
            // TopicMSFTLast
            // 
            this.TopicMSFTLast.String01 = "MSFT";
            this.TopicMSFTLast.String02 = "Last";
            this.TopicMSFTLast.Tag = "10";
            this.TopicMSFTLast.RefreshData += new AddinExpress.RTD.ADXRefreshData_EventHandler(this.CommonRefreshData);
            // 
            // TopicWMTLast
            // 
            this.TopicWMTLast.String01 = "WMT";
            this.TopicWMTLast.String02 = "Last";
            this.TopicWMTLast.Tag = "20";
            this.TopicWMTLast.RefreshData += new AddinExpress.RTD.ADXRefreshData_EventHandler(this.CommonRefreshData);
            // 
            // TopicATTLast
            // 
            this.TopicATTLast.String01 = "ATT";
            this.TopicATTLast.String02 = "Last";
            this.TopicATTLast.Tag = "30";
            this.TopicATTLast.RefreshData += new AddinExpress.RTD.ADXRefreshData_EventHandler(this.CommonRefreshData);
            // 
            // TopicINTCLast
            // 
            this.TopicINTCLast.String01 = "INTC";
            this.TopicINTCLast.String02 = "Last";
            this.TopicINTCLast.Tag = "40";
            this.TopicINTCLast.RefreshData += new AddinExpress.RTD.ADXRefreshData_EventHandler(this.CommonRefreshData);
            // 
            // RTDServerModule
            // 
            this.Interval = 4000;
            this.RTDInitialize += new System.EventHandler(this.RTDServerModule_RTDInitialize);

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
		public static void RTDServerRegister(Type t)
		{
			AddinExpress.RTD.ADXRTDServerModule.ADXRTDServerRegister(t);
		}

		[ComUnregisterFunctionAttribute]
		public static void RTDServerUnregister(Type t)
		{
			AddinExpress.RTD.ADXRTDServerModule.ADXRTDServerUnregister(t);
		}

		#endregion

		public static new RTDServerModule CurrentInstance
		{
			get
			{
				return AddinExpress.RTD.ADXRTDServerModule.CurrentInstance as RTDServerModule;
			}
		}

		private Random rnd = new Random(new Random().Next(1, 10000));

		private void RTDServerModule_RTDInitialize(object sender, EventArgs e)
		{
			// initialization
			TopicMSFTOpen.DefaultValue = (rnd.NextDouble() * 100 + 10).ToString("F");
			TopicMSFTLast.DefaultValue = TopicMSFTOpen.DefaultValue;
			TopicWMTOpen.DefaultValue = (rnd.NextDouble() * 100 + 10).ToString("F");
			TopicWMTLast.DefaultValue = TopicWMTOpen.DefaultValue;
			TopicATTOpen.DefaultValue = (rnd.NextDouble() * 100 + 10).ToString("F");
			TopicATTLast.DefaultValue = TopicATTOpen.DefaultValue;
			TopicINTCOpen.DefaultValue = (rnd.NextDouble() * 100 + 10).ToString("F");
			TopicINTCLast.DefaultValue = TopicINTCOpen.DefaultValue;
		}

		private object CommonRefreshData(object sender)
		{
			AddinExpress.RTD.ADXRTDTopic topic = sender as AddinExpress.RTD.ADXRTDTopic;
			if (topic != null)
			{
				if (topic.String02 == "Last")
				{
					double StartValue = 0;
					switch (topic.Tag)
					{
						case "10":
							StartValue = Convert.ToDouble(TopicMSFTOpen.DefaultValue);
							break;
						case "20":
							StartValue = Convert.ToDouble(TopicWMTOpen.DefaultValue);
							break;
						case "30":
							StartValue = Convert.ToDouble(TopicATTOpen.DefaultValue);
							break;
						case "40":
							StartValue = Convert.ToDouble(TopicINTCOpen.DefaultValue);
							break;
					}
					return StartValue + rnd.Next(-10, 10);
				}
			}
			return string.Empty;
		}
	}
}
