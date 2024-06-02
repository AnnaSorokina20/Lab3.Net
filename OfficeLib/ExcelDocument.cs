using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeLib
{
    public class ExcelDocument : IDisposable
    {
        Excel.Application? app = null;
        Excel.Workbook? book = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;

        public ExcelDocument()
        {
            app = new Excel.Application();
            book = app.Workbooks.Add();
            sheets = book.Sheets;
            sheet = sheets[1];
        }
        public ExcelDocument(string FileName)
        {
            app = new Excel.Application();
            book = app.Workbooks.Open(FileName);
            sheets = book.Sheets;
            sheet = sheets[1];
        }

        public void SaveAs(string FileName)
        {
            book?.SaveAs(FileName);
        }

        public void Dispose()
        {
            book?.Close();
            app?.Quit();
            Release(sheet); sheet = null;
            Release(sheets); sheets = null;
            Release(book); book = null;
            Release(app); app = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public string? this[string cellName]
        {
            get => sheet?.Range[cellName].Value2.ToString();
            set
            {
                if (sheet != null)
                    sheet.Range[cellName].Value2 = value;
            }
        }
        public string? this[int row, int col]
        {
            get => sheet?.Cells[row, col].Value2.ToString();
            set
            {
                if (sheet != null)
                    sheet.Cells[row, col] = value;
            }
        }

        private void Release(object? obj)
        {
            if (obj != null)
#pragma warning disable CA1416 // Validate platform compatibility
                _ = Marshal.FinalReleaseComObject(obj);
#pragma warning restore CA1416 // Validate platform compatibility
        }
    }

    public class WordDocument : IDisposable
    {
        Word.Application? app = null;
        Word.Document? doc = null;

        public WordDocument()
        {
            app = new Word.Application();
            doc = app.Documents.Add();
        }
        public WordDocument(string FileName)
        {
            app = new Word.Application();
            doc = app.Documents.Open(FileName);
        }

        public void SaveAs(string FileName)
        {
            doc?.SaveAs2(FileName);
        }

        public void Dispose()
        {
            doc?.Close();
            app?.Quit();
            Release(doc); doc = null;
            Release(app); app = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public string? this[int paragraphIndex]
        {
            get => doc?.Paragraphs[paragraphIndex].Range.Text;
            set
            {
                if (doc != null)
                {
                    var paragraph = doc.Paragraphs.Add();
                    paragraph.Range.Text = value;
                }
            }
        }

        private void Release(object? obj)
        {
            if (obj != null)
#pragma warning disable CA1416 // Validate platform compatibility
                _ = Marshal.FinalReleaseComObject(obj);
#pragma warning restore CA1416 // Validate platform compatibility
        }
    }
}