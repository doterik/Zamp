using System;
using System.Runtime.InteropServices;

namespace ExcelAutomationAddin
{
	/// <summary>
	///   Add-in Express Excel Add-in Module
	/// </summary>
	[GuidAttribute("42CC13B9-9FB7-43F1-BF29-67B82ED58D6C"),
	ProgId("ExcelAutomationAddin.ExcelAddinModule1"), ClassInterface(ClassInterfaceType.AutoDual)]
	public class ExcelAddinModule1 : AddinExpress.MSO.ADXExcelAddinModule
	{
		public ExcelAddinModule1()
		{
			InitializeComponent();
		}

		#region Component Designer generated code

		/// <summary>
		/// Required by designer support - do not modify
		/// the following method
		/// </summary>
		private void InitializeComponent()
		{

		}

		#endregion

		#region Add-in Express automatic code

		[ComRegisterFunctionAttribute]
		public static void AddinRegister(Type t)
		{
			AddinExpress.MSO.ADXExcelAddinModule.ADXExcelAddinRegister(t);
		}

		[ComUnregisterFunctionAttribute]
		public static void AddinUnregister(Type t)
		{
			AddinExpress.MSO.ADXExcelAddinModule.ADXExcelAddinUnregister(t);
		}

		#endregion

		public object MyFunc(object range)
		{
			string retVal = string.Empty;
			try
			{
				if (range is Excel.Range)
				{
					Excel.Range xlRange = range as Excel.Range;
					try
					{
						int row = xlRange.Row;
						int col = xlRange.Column;
						retVal = "The range address is: [" +
							xlRange.get_Address(row, col, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing) + "]";
						retVal = retVal.Replace("$", string.Empty);
					}
					finally { Marshal.ReleaseComObject(range); }
				}
			}
			catch
			{
				return CVErr(AddinExpress.MSO.ADXxlCVError.xlErrNull);
			}
			return retVal;
		}
	}
}
