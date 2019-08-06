
using JobManagement;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelConsole
{
    class Program
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        static int hWnd;

        static void Main(string[] args)
        {
            //DoNotReleaseCOM();
            //ReleaseCOM();
            ForLoopNoReleaseCOM();
            //ForLoopReleaseCOM();
            //KillExcelWithWM_CLOSE();
            //KillExcelWithProcessKill();
            //KillExcelWithWindowsJob();
        }

        static void DoNotReleaseCOM()
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Add();
            Excel.Worksheet sheet = app.Sheets.Add();

            hWnd = app.Hwnd;

            sheet.Range["A1"].Value = "Lorem Ipsum";
            book.SaveAs(@"C:\Temp\ExcelBook" + DateTime.Now.Millisecond + ".xlsx");
            book.Close();
            app.Quit();
        }

        static void ReleaseCOM()
        {
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
        }

        static void ForLoopNoReleaseCOM()
        {
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

                foreach (Excel.Worksheet sheet in sheets)
                {
                    Console.WriteLine(sheet.Name);
                    Marshal.ReleaseComObject(sheet);
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
        }

        static void ForLoopReleaseCOM()
        {
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
                Console.Read();
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
        }

        static void KillExcelWithWM_CLOSE()
        {
            DoNotReleaseCOM();
            SendMessage((IntPtr)hWnd, 0x10, IntPtr.Zero, IntPtr.Zero);
        }

        static void KillExcelWithProcessKill()
        {
            DoNotReleaseCOM();
            Process[] excelProcs = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in excelProcs)
            {
                proc.Kill();
            }
        }

        static void KillExcelWithWindowsJob()
        {
            DoNotReleaseCOM();
            Job job = new Job();
            uint pid = 0;
            GetWindowThreadProcessId(new IntPtr(hWnd), out pid);
            job.AddProcess(Process.GetProcessById((int)pid).Handle);
        }
    }
}
