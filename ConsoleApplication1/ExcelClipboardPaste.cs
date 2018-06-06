using System;
using NetOffice.ExcelApi;

namespace ConsoleApplication1
{
    public class ExcelClipboardPaste : IDisposable
    {
        private readonly Application _excelApp;

        public ExcelClipboardPaste()
        {
            _excelApp = new Application
            {
                Visible = false,
                DisplayAlerts = false
            };
        }

        public void Paste(string path)
        {
            var wb = _excelApp.Workbooks.Add();
            var ws = (Worksheet) wb.Worksheets.Add();
            ws.Select();
            ws.Paste();
            wb.SaveAs(path);
            wb.Close(false);
        }

        public void Dispose()
        {
            _excelApp?.Quit();
        }
    }
}