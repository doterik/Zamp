using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
//using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Test2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e) /* @region VSTO */
        {
            var bankAccounts = new List<Account>
            {
                new Account {ID = 345, Balance = 541.27},
                new Account {ID = 123, Balance = -127.44}
            };

            // This multiline lambda expression sets custom processing rules for the bankAccounts.
            DisplayInExcel(bankAccounts, (account, cell) =>
            {
                cell.Value = account.ID;
                cell.Offset[0, 1].Value = account.Balance;
                if (account.Balance < 0)
                {
                    cell.Interior.Color = 255;
                    cell.Offset[0, 1].Interior.Color = 255;
                }
            });

            var wordApp = new Word.Application();
            wordApp.Visible = true;
            wordApp.Documents.Add();
            wordApp.Selection.PasteSpecial(Link: true, DisplayAsIcon: true);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { } /* @region VSTO */

        private void DisplayInExcel(IEnumerable<Account> accounts, Action<Account, Excel.Range> DisplayFunc)
        {
            var excelApp = this.Application;
            // Add a new Excel workbook.
            excelApp.Workbooks.Add();
            excelApp.Visible = true;
            excelApp.Range["A1"].Value = "ID";
            excelApp.Range["B1"].Value = "Balance";
            excelApp.Range["A2"].Select();

            foreach (var ac in accounts)
            {
                DisplayFunc(ac, excelApp.ActiveCell);
                excelApp.ActiveCell.Offset[1, 0].Select();
            }
            // Copy the results to the Clipboard.
            excelApp.Range["A1:B3"].Copy();

            //excelApp.Columns[1].AutoFit(); // Use with: Embed Interop Types = True
            //excelApp.Columns[2].AutoFit();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion VSTO generated code
    }
}


/*
Excel.Application app = null;
Excel.Workbooks books = null;
Excel.Workbook book = null;
Excel.Sheets sheets = null;
Excel.Worksheet sheet = null;
Excel.Range range = null;
 
try
{
    app = new Excel.Application();
    books = app.Workbooks;
    book = books.Add();
    sheets = book.Sheets;
    sheet = sheets.Add();
    range = sheet.Range["A1"];
    range.Value = "Lorem Ipsum";
    book.SaveAs(@"C:\Temp\ExcelBook" + DateTime.Now.Millisecond + ".xlsx");
    book.Close();
    app.Quit();
}
finally
{
    if (range != null) Marshal.ReleaseComObject(range);
    if (sheet != null) Marshal.ReleaseComObject(sheet);
    if (sheets != null) Marshal.ReleaseComObject(sheets);
    if (book != null) Marshal.ReleaseComObject(book);
    if (books != null) Marshal.ReleaseComObject(books);
    if (app != null) Marshal.ReleaseComObject(app);
}
Pay close attention to the above code, we never used more than one dot when working with objects.
We also wrapped all the code in a try-finally, so even if the code throws and exception we will still safely release the COM objects using the ReleaseComObject method on the Marshal object.
*/

/*
To be on the safe side, you should avoid using a foreach loop and rather use a normal for loop,
and release each COM object in the collection, as illustrated below:

Excel.Application app = null;
Excel.Workbooks books = null;
Excel.Workbook book = null;
Excel.Sheets sheets = null;
 
try
{
    app = new Excel.Application();
    books = app.Workbooks;
    book = books.Open(@"C:\Temp\ExcelBook1Sheets.xlsx");
    sheets = book.Sheets;
 
    for (int i = 1; i <= sheets.Count; i++)
    {
        Excel.Worksheet sheet = sheets.Item[i];
        Console.WriteLine(sheet.Name);
        if (sheet != null) Marshal.ReleaseComObject(sheet);
    }
    book.Close();
    app.Quit();
}
finally
{
    if (sheets != null) Marshal.ReleaseComObject(sheets);
    if (book != null) Marshal.ReleaseComObject(book);
    if (books != null) Marshal.ReleaseComObject(books);
    if (app != null) Marshal.ReleaseComObject(app);
}
*/

/*
So, it is totally acceptable to see the following code in you COM add-in projects:

GC.Collect();
GC.WaitForPendingFinalizers();
GC.Collect();
GC.WaitForPendingFinalizers();
*/