using System;
using NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace ConsoleApplication1
{
    public class WordClipboardCopy : IDisposable
    {
        private readonly Application _wordApp;

        public WordClipboardCopy()
        {
            _wordApp = new Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
        }

        public void Copy(string path)
        {
            var word = _wordApp.Documents.OpenNoRepairDialog(path);
            word.Select();
            _wordApp.Selection.Copy();
            word.Close(false);
        }

        public void Dispose()
        {
            _wordApp?.Dispose();
        }
    }
}