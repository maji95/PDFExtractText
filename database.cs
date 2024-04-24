/* using System;
using System.Data.SQLite;
using OfficeOpenXml;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Устанавливаем контекст лицензирования для библиотеки EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // или LicenseContext.Commercial

            // Путь к файлу базы данных SQLite
            string databasePath = "C:\path\to\database.db";

            // Создаем базу данных SQLite и таблицу FolderData
            using (var connection = new SQLiteConnection($"Data Source={databasePath};Version=3;"))
            {
                connection.Open();

                // Создаем таблицу FolderData
                using (var command = new SQLiteCommand("CREATE TABLE IF NOT EXISTS FolderData (ID INTEGER PRIMARY KEY AUTOINCREMENT, FolderName TEXT, FolderPath TEXT);", connection))
                {
                    command.ExecuteNonQuery();
                }
            }

            // Загрузка данных из файла Excel и добавление их в базу данных SQLite
            LoadDataFromExcelAndInsertIntoSQLite(databasePath, @"C:\path\to\output.xlsx");
        }

        static void LoadDataFromExcelAndInsertIntoSQLite(string databasePath, string excelFilePath)
        {
            // Открываем файл Excel
            using (var package = new ExcelPackage(new System.IO.FileInfo(excelFilePath)))
            {
                // Получаем ссылку на лист Excel
                var worksheet = package.Workbook.Worksheets[0];

                // Открываем соединение с базой данных SQLite
                using (var connection = new SQLiteConnection($"Data Source={databasePath};Version=3;"))
                {
                    connection.Open();

                    // Перебираем строки в файле Excel и добавляем данные в базу данных SQLite
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        // Получаем значения ячеек из текущей строки
                        string folderName = worksheet.Cells[row, 1].Value?.ToString();
                        string folderPath = worksheet.Cells[row, 2].Value?.ToString();

                        // Вставляем данные в таблицу FolderData
                        using (var command = new SQLiteCommand("INSERT INTO FolderData (FolderName, FolderPath) VALUES (@FolderName, @FolderPath);", connection))
                        {
                            command.Parameters.AddWithValue("@FolderName", folderName);
                            command.Parameters.AddWithValue("@FolderPath", folderPath);
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }

            Console.WriteLine("Данные из файла Excel успешно добавлены в базу данных SQLite.");
        }

    }
}
*/