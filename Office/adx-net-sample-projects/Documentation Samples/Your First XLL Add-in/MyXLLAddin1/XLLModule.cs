using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace MyXLLAddin1
{
	/// <summary>
	///   Add-in Express XLL Add-in Module
	/// </summary>
	[ComVisible(true)]
	public class XLLModule : AddinExpress.MSO.ADXXLLModule
	{
		public XLLModule()
		{
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXExcelFunctionCategory AdxExcelFunctionCategory1;
		private AddinExpress.MSO.ADXExcelFunctionDescriptor AdxExcelFunctionDescriptor1;
		private AddinExpress.MSO.ADXExcelParameterDescriptor AdxExcelParameterDescriptor1;

		#region Component Designer generated code

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.AdxExcelFunctionCategory1 = new AddinExpress.MSO.ADXExcelFunctionCategory(this.components);
			this.AdxExcelFunctionDescriptor1 = new AddinExpress.MSO.ADXExcelFunctionDescriptor(this.components);
			this.AdxExcelParameterDescriptor1 = new AddinExpress.MSO.ADXExcelParameterDescriptor(this.components);
			// 
			// AdxExcelFunctionCategory1
			// 
			this.AdxExcelFunctionCategory1.CategoryName = "My Category";
			this.AdxExcelFunctionCategory1.FunctionDescriptors.Add(this.AdxExcelFunctionDescriptor1);
			// 
			// AdxExcelFunctionDescriptor1
			// 
			this.AdxExcelFunctionDescriptor1.Description = "This is my first UDF";
			this.AdxExcelFunctionDescriptor1.FunctionName = "MyFunction";
			this.AdxExcelFunctionDescriptor1.ParameterDescriptors.Add(this.AdxExcelParameterDescriptor1);
			// 
			// AdxExcelParameterDescriptor1
			// 
			this.AdxExcelParameterDescriptor1.Description = "Accepts numeric values";
			this.AdxExcelParameterDescriptor1.ParameterName = "arg";
			// 
			// XLLModule
			// 
			this.AddinName = "MyXLLAddin1";

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
		public static void RegisterXLL(Type t)
		{
			AddinExpress.MSO.ADXXLLModule.RegisterXLLInternal(t);
		}

		[ComUnregisterFunctionAttribute]
		public static void UnregisterXLL(Type t)
		{
			AddinExpress.MSO.ADXXLLModule.UnregisterXLLInternal(t);
		}

		#endregion

		public static new XLLModule CurrentInstance
		{
			get
			{
				return AddinExpress.MSO.ADXXLLModule.CurrentInstance as XLLModule;
			}
		}

		public Excel._Application ExcelApp
		{
			get
			{
				return (HostApplication as Excel._Application);
			}
		}

		#region Define your UDFs in this section

		/// <summary>
		/// The container for user-defined functions (UDFs). Every UDF is a public static (Public Shared in VB.NET) method that returns a value of any base type: string, double, integer.
		/// </summary>
		internal static class XLLContainer
		{
			/// <summary>
			/// Required by Add-in Express. Please do not modify this method.
			/// </summary>
			internal static XLLModule Module
			{
				get
				{
					return AddinExpress.MSO.ADXXLLModule.
						CurrentInstance as MyXLLAddin1.XLLModule;
				}
			}

			#region Sample function

			// Demonstrates how to handle all parameter types available for UDFs.
			// Uncomment the code, click Register Add-in Express Project in the Build menu, and run Excel.

			//public static string AllSupportedExcelTypes(object arg)
			//{
			//    if (arg is double)
			//        return "Double: " + (double)arg;
			//    else if (arg is string)
			//        return "String: " + (string)arg;
			//    else if (arg is bool)
			//        return "Boolean: " + (bool)arg;
			//    else if (arg is AddinExpress.MSO.ADXExcelError)
			//        return "ExcelError: " + arg.ToString();
			//    else if (arg is object[,])
			//        return string.Format("Array[{0},{1}]", ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
			//    else if (arg is System.Reflection.Missing)
			//        return "Missing";
			//    else if (arg == null)
			//        return "Empty";
			//    else if (arg is AddinExpress.MSO.ADXExcelRef)
			//    {
			//        AddinExpress.MSO.ADXExcelRef reference = arg as AddinExpress.MSO.ADXExcelRef;
			//        return string.Format("Reference [{0},{1},{2},{3}]", reference.ColumnFirst, reference.RowFirst, reference.ColumnLast, reference.RowLast);
			//    }
			//    else if (arg is short)
			//        return "Short: " + (short)arg;
			//    else
			//        return "Unknown Type";
			//}

			#endregion

			public static object MyFunction(object arg)
			{
				if (arg is double)
				{
					System.Random rnd = new System.Random(2000);
					return rnd.NextDouble();
				}
				else
				{
					if (Module.IsInFunctionWizard)
						return "The parameter must be numeric!";
					else
						return AddinExpress.MSO.ADXExcelError.xlErrorNum;
				}
			}
		}

		#endregion
	}
}
