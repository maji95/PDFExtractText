/* using System;
using System.IO;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

public class PdfTextExtractorExample
{
    public static void Main()
    {
        string filePath = @"C:\path\to\Analiza";
        string text = ExtractTextFromPdf(filePath);
        Console.WriteLine(text);
    }

    public static string ExtractTextFromPdf(string path)
    {
        using (PdfReader reader = new PdfReader(path))
        {
            StringBuilder text = new StringBuilder();

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
            }

            return text.ToString();
        }
    }
}
*/