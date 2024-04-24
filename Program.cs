using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using ClosedXML.Excel;
using System.Text;
using Path = System.IO.Path;

public class PdfDataExtractor
{
    static void Main(string[] args)
    {
        string rootDirectory = @"C:\path\to\Analiza";
        var records = ProcessPdfFiles(rootDirectory);
        SaveDataToExcel(records, @"C:\path\to\output.xlsx");
    }

    public static List<(string FolderName, string FolderPath, string FileName, string FilePath, string CNP)> ProcessPdfFiles(string rootDir)
    {
        var records = new List<(string, string, string, string, string)>();

        foreach (var dirPath in Directory.GetDirectories(rootDir, "*", SearchOption.AllDirectories))
        {
            foreach (var filePath in Directory.GetFiles(dirPath, "*.pdf"))
            {
                Console.WriteLine($"Processing file: {filePath}");
                var text = ExtractTextFromPdf(filePath);
                var cnp = ExtractCNP(text);
                records.Add((Path.GetFileName(dirPath), dirPath, Path.GetFileName(filePath), filePath, cnp));
            }
        }
        return records;
    }



    private static string ExtractTextFromPdf(string filePath)
    {
        try
        {
            using (PdfReader reader = new PdfReader(filePath))
            {
                var text = new StringBuilder();
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
                return text.ToString();
            }
        }
        catch (iTextSharp.text.exceptions.InvalidPdfException ex)
        {
            Console.WriteLine($"Error processing file {filePath}: {ex.Message}");
            return "";  // Возвращаем пустую строку, если файл не может быть обработан
        }
    }


    private static string ExtractCNP(string text)
    {
        string[] lines = text.Split('\n');  // Разбиение текста на строки
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains("CNP:"))  // Проверка на наличие маркера "CNP:" с двоеточием
            {
                if (i + 1 < lines.Length)  // Убедиться, что следующая строка существует
                {
                    string nextLine = lines[i + 1].Trim();  // Обрезать пробелы в следующей строке
                    if (nextLine.Length >= 13)  // Проверить длину строки
                    {
                        return nextLine.Substring(0, 13);  // Взять первые 13 символов
                    }
                }
            }
            else if (lines[i].Contains("CNP"))  // Проверка на наличие маркера "CNP" без двоеточия
            {
                var match = Regex.Match(lines[i], @"CNP\s+(\d{13})");
                if (match.Success)
                {
                    return match.Groups[1].Value;  // Вернуть найденное значение CNP
                }
            }
        }
        return "CNP не найден";
    }



    private static void SaveDataToExcel(List<(string FolderName, string FolderPath, string FileName, string FilePath, string CNP)> data, string filePath)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Data");
            worksheet.Cell("A1").Value = "Folder Name";
            worksheet.Cell("B1").Value = "Folder Path";
            worksheet.Cell("C1").Value = "File Name";
            worksheet.Cell("D1").Value = "File Path";
            worksheet.Cell("E1").Value = "CNP";
            

            int row = 2;
            foreach (var record in data)
            {
                worksheet.Cell(row, 1).Value = record.FolderName;
                worksheet.Cell(row, 2).Value = record.FolderPath;
                worksheet.Cell(row, 3).Value = record.FileName;
                worksheet.Cell(row, 4).Value = record.FilePath;
                worksheet.Cell(row, 5).Value = record.CNP;
                
                row++;
            }

            workbook.SaveAs(filePath);
        }
    }
    
}
