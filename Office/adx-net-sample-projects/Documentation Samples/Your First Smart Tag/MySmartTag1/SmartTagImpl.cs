using System;
using System.Runtime.InteropServices;

namespace MySmartTag1
{
	[GuidAttribute("B0690877-3BDA-4034-ACAC-9D9AC56224B4"), ClassInterface(ClassInterfaceType.None)]
	public class SmartTagRecognizerImpl : AddinExpress.SmartTag.ADXSmartTagRecognizerImpl
	{
		public SmartTagRecognizerImpl()
			: base(typeof(MySmartTag1.SmartTagModule))
		{
		}
	}

	[GuidAttribute("4659A7A3-8942-4977-87E0-6A8726A13478"), ClassInterface(ClassInterfaceType.None)]
	public class SmartTagActionImpl : AddinExpress.SmartTag.ADXSmartTagActionImpl
	{
		public SmartTagActionImpl()
			: base(typeof(MySmartTag1.SmartTagModule))
		{
		}
	}
}
