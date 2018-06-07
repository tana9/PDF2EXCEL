using System;
using System.Data.SqlTypes;
using System.IO;
using System.Xml;
using NetOffice.WordApi.Enums;

namespace ConsoleApplication1
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            var path = @"C:\Users\t-tan\Desktop\Book1.pdf";
            var outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel");
            Directory.CreateDirectory(outDir);
            var excelPath = Path.Combine(outDir, Path.GetFileNameWithoutExtension(path) + ".xlsx");
            using (var word = new WordClipboardCopy())
            using (var excel = new ExcelClipboardPaste())
            {
                word.Copy(path);
                excel.Paste(excelPath);
            }
        }
    }
}