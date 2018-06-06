using System.Data.SqlTypes;
using NetOffice.WordApi.Enums;

namespace ConsoleApplication1
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var path = @"C:\Users\t-tan\Desktop\Book1.pdf";
            var excelpath = @"C:\Users\t-tan\Desktop\Book2.xlsx";
            using (var word = new WordClipboardCopy())
            using (var excel = new ExcelClipboardPaste())
            {
                word.Copy(path);
                excel.Paste(excelpath);
            }
        }
    }
}