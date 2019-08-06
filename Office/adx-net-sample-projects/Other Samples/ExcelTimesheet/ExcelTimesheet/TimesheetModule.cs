using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;

namespace ExcelTimesheet
{
	/// <summary>
	///   Add-in Express Excel Worksheet Module
	/// </summary>
	[GuidAttribute("D6131CA1-CA5F-4FF8-A274-7B07DCA70985"), ProgId("ExcelTimesheet.TimesheetModule")]
	public class TimesheetModule : AddinExpress.MSO.ADXExcelSheetModule
	{
		public TimesheetModule()
		{
			InitializeComponent();
		}

		private AddinExpress.MSO.ADXMSFormsComboBox adxEmployeeBox;
		private AddinExpress.MSO.ADXMSFormsComboBox adxApprovedByBox;
		private AddinExpress.MSO.ADXMSFormsCommandButton adxPeriodFromBtn;
		private AddinExpress.MSO.ADXMSFormsCommandButton adxPeriodToBtn;
		private AddinExpress.MSO.ADXMSFormsCommandButton adxImportBtn;
		private AddinExpress.MSO.ADXCommandBar adxCommandBar1;
		internal AddinExpress.MSO.ADXCommandBarButton adxClearTimesheet;
		private AddinExpress.MSO.ADXRibbonTab adxRibbonTab1;
		private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
		internal AddinExpress.MSO.ADXRibbonButton adxRibbonButtonClearTimesheet;

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
			this.adxEmployeeBox = new AddinExpress.MSO.ADXMSFormsComboBox(this.components);
			this.adxApprovedByBox = new AddinExpress.MSO.ADXMSFormsComboBox(this.components);
			this.adxPeriodFromBtn = new AddinExpress.MSO.ADXMSFormsCommandButton(this.components);
			this.adxPeriodToBtn = new AddinExpress.MSO.ADXMSFormsCommandButton(this.components);
			this.adxImportBtn = new AddinExpress.MSO.ADXMSFormsCommandButton(this.components);
			this.adxCommandBar1 = new AddinExpress.MSO.ADXCommandBar(this.components);
			this.adxClearTimesheet = new AddinExpress.MSO.ADXCommandBarButton(this.components);
			this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
			this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
			this.adxRibbonButtonClearTimesheet = new AddinExpress.MSO.ADXRibbonButton(this.components);
			// 
			// adxEmployeeBox
			// 
			this.adxEmployeeBox.ControlName = "EmployeeBox";
			this.adxEmployeeBox.Connect += new System.EventHandler(this.adxEmployeeBox_Connect);
			this.adxEmployeeBox.Click += new System.EventHandler(this.adxEmployeeBox_Click);
			// 
			// adxApprovedByBox
			// 
			this.adxApprovedByBox.ControlName = "ApprovedByBox";
			this.adxApprovedByBox.Connect += new System.EventHandler(this.adxApprovedByBox_Connect);
			this.adxApprovedByBox.Click += new System.EventHandler(this.adxApprovedByBox_Click);
			// 
			// adxPeriodFromBtn
			// 
			this.adxPeriodFromBtn.ControlName = "PeriodFromBtn";
			this.adxPeriodFromBtn.Click += new System.EventHandler(this.adxPeriodFromBtn_Click);
			// 
			// adxPeriodToBtn
			// 
			this.adxPeriodToBtn.ControlName = "PeriodToBtn";
			this.adxPeriodToBtn.Click += new System.EventHandler(this.adxPeriodToBtn_Click);
			// 
			// adxImportBtn
			// 
			this.adxImportBtn.ControlName = "ImportBtn";
			this.adxImportBtn.Click += new System.EventHandler(this.adxImportBtn_Click);
			// 
			// adxCommandBar1
			// 
			this.adxCommandBar1.CommandBarName = "Excel Timesheet Bar";
			this.adxCommandBar1.CommandBarTag = "c5203094-114e-42f5-af0e-93441442b896";
			this.adxCommandBar1.Controls.Add(this.adxClearTimesheet);
			this.adxCommandBar1.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
			this.adxCommandBar1.UpdateCounter = 9;
			// 
			// adxClearTimesheet
			// 
			this.adxClearTimesheet.Caption = "Clear Timesheet";
			this.adxClearTimesheet.ControlTag = "8b0cb1ee-366b-486b-9d7e-fdef6aef0157";
			this.adxClearTimesheet.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxClearTimesheet.UpdateCounter = 13;
			this.adxClearTimesheet.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxClearTimesheet_Click);
			// 
			// adxRibbonTab1
			// 
			this.adxRibbonTab1.Caption = "Excel Timesheet Sample";
			this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
			this.adxRibbonTab1.Id = "adxRibbonTab_5065698129f84f2caddbe1a393506eb5";
			// 
			// adxRibbonGroup1
			// 
			this.adxRibbonGroup1.Caption = "Excel Timesheet";
			this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonClearTimesheet);
			this.adxRibbonGroup1.Id = "adxRibbonGroup_b97581263aaf4862acb03c62c5545d90";
			this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
			// 
			// adxRibbonButtonClearTimesheet
			// 
			this.adxRibbonButtonClearTimesheet.Caption = "Clear Timesheet";
			this.adxRibbonButtonClearTimesheet.Id = "adxRibbonButton_8087e5cf93fd45959fca33f4dbda347d";
			this.adxRibbonButtonClearTimesheet.ImageTransparentColor = System.Drawing.Color.Transparent;
			this.adxRibbonButtonClearTimesheet.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonClearTimesheet_OnClick);
			// 
			// TimesheetModule
			// 
			this.Description = "ADX Excel Worksheet Module Example";
			this.ModuleName = "TimesheetModule";
			this.PropertyId = "_ADX_ExcelTimesheet";
			this.PropertyValue = "Created for the ExcelTimesheet project.";
			this.Worksheet = "Timesheet";
			this.Activate += new System.EventHandler(this.TimesheetModule_Activate);
			this.Deactivate += new System.EventHandler(this.TimesheetModule_Deactivate);

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

		private string[] Employees = new string[] { "Faas, Matt", "Parker, Daniel", "Smith, John" };
		private string[] Positions = new string[] { "Sales Manager", "Supplier", "Office worker" };
		private string[] Departments = new string[] { "Marketing department", "Bespoke department", "Requisitioning department" };
		private string[] Numbers = new string[] { "23486", "20465", "54789" };

		private void adxEmployeeBox_Connect(object sender, System.EventArgs e)
		{
			try
			{
				AddinExpress.MSO.ADXMSFormsComboBox cmb = MSFControlByName("EmployeeBox") as AddinExpress.MSO.ADXMSFormsComboBox;
				if (cmb.ListCount == 0)
				{
					cmb.AddItem(Employees[0]);
					cmb.AddItem(Employees[1]);
					cmb.AddItem(Employees[2]);
				}
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void adxApprovedByBox_Connect(object sender, System.EventArgs e)
		{
			try
			{
				AddinExpress.MSO.ADXMSFormsComboBox cmb = MSFControlByName("ApprovedByBox") as AddinExpress.MSO.ADXMSFormsComboBox;
				if (cmb.ListCount == 0)
					cmb.AddItem("Harrison, George, director");
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void adxEmployeeBox_Click(object sender, System.EventArgs e)
		{
			try
			{
				Excel.Range r = GetRange("E10", System.Type.Missing) as Excel.Range;
				if (r != null)
					try
					{
						r.Value = Employees[adxEmployeeBox.ListIndex];
					}
					finally { Marshal.ReleaseComObject(r); }

				r = GetRange("E11", System.Type.Missing) as Excel.Range;
				if (r != null)
					try
					{
						r.Value = Positions[adxEmployeeBox.ListIndex];
					}
					finally { Marshal.ReleaseComObject(r); }

				r = GetRange("E12", System.Type.Missing) as Excel.Range;
				if (r != null)
					try
					{
						r.Value = Departments[adxEmployeeBox.ListIndex];
					}
					finally { Marshal.ReleaseComObject(r); }

				r = GetRange("I10", System.Type.Missing) as Excel.Range;
				if (r != null)
					try
					{
						r.Value = Numbers[adxEmployeeBox.ListIndex];
					}
					finally { Marshal.ReleaseComObject(r); }

				adxImportBtn.Enabled = true;
				adxEmployeeBox.Enabled = true;
				adxApprovedByBox.Enabled = true;
				adxPeriodFromBtn.Enabled = true;
				adxPeriodToBtn.Enabled = true;
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void adxPeriodFromBtn_Click(object sender, System.EventArgs e)
		{
			ShowCalendar("E16");
		}

		private void adxPeriodToBtn_Click(object sender, System.EventArgs e)
		{
			ShowCalendar("G16");
		}

		private void ShowCalendar(string cell)
		{
			CalendarForm form = new CalendarForm();
			try
			{
				if (form.ShowDialog() == DialogResult.OK)
				{
					Excel.Range r = GetRange(cell, System.Type.Missing) as Excel.Range;
					if (r != null)
						try
						{
							r.Value = form.Calendar.SelectionStart.ToShortDateString();
						}
						finally { Marshal.ReleaseComObject(r); }
				}
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				if (form != null)
					form.Dispose();
			}
		}

		private void adxApprovedByBox_Click(object sender, System.EventArgs e)
		{
			try
			{
				Excel.Range r = GetRange("M16", System.Type.Missing) as Excel.Range;
				if (r != null)
					try
					{
						r.Value = "Harrison, George, director";
					}
					finally { Marshal.ReleaseComObject(r); }
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ClearRanges()
		{
			Excel.Range r = GetRange("D19", "P35") as Excel.Range;
			if (r != null)
				try
				{
					r.ClearContents();
				}
				finally { Marshal.ReleaseComObject(r); }
		}

		private void adxImportBtn_Click(object sender, System.EventArgs e)
		{
			System.Random rnd = new System.Random();
			int rows = rnd.Next(1, 10);
			string[] letters = new string[] { "J", "K", "L", "M", "N", "O", "P" };

			try
			{
				ClearRanges();
				for (int i = 1; i <= rows; i++)
				{
					Excel.Range r = GetRange("D" + ((int)(18 + i)).ToString(), System.Type.Missing) as Excel.Range;
					if (r != null)
						try
						{
							r.Value = i.ToString() + ".";
						}
						finally { Marshal.ReleaseComObject(r); }
					
					r = GetRange("E" + ((int)(18 + i)).ToString(), System.Type.Missing) as Excel.Range;
					if (r != null)
						try
						{
							r.Value = "Task " + i.ToString() + ".";
						}
						finally { Marshal.ReleaseComObject(r); }

					int days = rnd.Next(1, 7);
					for (int d = 0; d < days; d++)
					{
						r = GetRange(letters[d] + ((int)(18 + i)).ToString(), System.Type.Missing) as Excel.Range;
						if (r != null)
							try
							{
								r.Value = rnd.Next(1, 4);
							}
							finally { Marshal.ReleaseComObject(r); }
					}
				}
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public void OnWindowActivate(object workbook)
		{
			if (IsConnected)
			{
				Excel._Workbook book = this.Parent as Excel._Workbook;
				if (book != null)
					try
					{
						string parentBookFullName = (string)book.GetType().InvokeMember(
							"FullName", System.Reflection.BindingFlags.GetProperty, null, book, null, CultureInfo.CurrentUICulture);
						string workbookFullName = (string)workbook.GetType().InvokeMember(
							"FullName", System.Reflection.BindingFlags.GetProperty, null, workbook, null, CultureInfo.CurrentUICulture);
						if (string.Compare(parentBookFullName, workbookFullName, true) == 0)
						{
							Excel._Worksheet sheet = book.ActiveSheet as Excel._Worksheet;
							if (sheet != null)
								try
								{
									if (String.Compare(sheet.Name, Worksheet, true) == 0)
									{
										adxClearTimesheet.Enabled = true;
										adxRibbonButtonClearTimesheet.Enabled = true;
									}
								}
								finally { Marshal.ReleaseComObject(sheet); }
						}
					}
					catch (Exception err)
					{
						MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					finally { Marshal.ReleaseComObject(book); }
			}
		}

		private void TimesheetModule_Activate(object sender, System.EventArgs e)
		{
			if (IsConnected)
			{
				adxClearTimesheet.Enabled = true;
				adxRibbonButtonClearTimesheet.Enabled = true;
			}
		}

		private void TimesheetModule_Deactivate(object sender, System.EventArgs e)
		{
			adxClearTimesheet.Enabled = false;
			adxRibbonButtonClearTimesheet.Enabled = false;
		}

		private void ClearRange(object cell1, object cell2)
		{
			Excel.Range r = GetRange(cell1, cell2) as Excel.Range;
			if (r != null)
				try
				{
					r.ClearContents();
				}
				finally { Marshal.ReleaseComObject(r); }
		}

		private void adxClearTimesheet_Click(object sender)
		{
			try
			{
				ClearRanges();
				ClearRange("E16", System.Type.Missing);
				ClearRange("G16", "H16");
				ClearRange("M16", "Q16");
				ClearRange("E10", "F12");
				ClearRange("I10", System.Type.Missing);
				ClearRange("M9", "Q12");
				adxApprovedByBox.ListIndex = -1;
				adxEmployeeBox.ListIndex = -1;
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message, err.Source, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void adxRibbonButtonClearTimesheet_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
		{
			adxClearTimesheet_Click(null);
		}
	}
}
