using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Globalization;
using System.IO;
using CalendarDataToAxData.Model;
using CalendarDataToAxData.Logic;
using System.Linq;
using Microsoft.Extensions.Configuration;
using CalendarDataToAxData.Extension;
using System.Text;
using OfficeOpenXml;
using System.Diagnostics;
using CalendarDataToAxData.Common;
using CalendarDataToAxData.Logic.CalendarReader;

namespace CalendarDataToAxData
{
    internal static class Program
    {
        static Program()
        {
            // Поддержка windows-1251
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Включаем EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        static void Main(string[] args)
        {
            try
            {
                // Валидируем App.Config
                if (!AppConfigProvider.IsSettingExists(Constants.AppConfig.KeyNames.SubjectToIgnore))
                {
                    Console.WriteLine($"В App.config должен быть ключ {Constants.AppConfig.KeyNames.SubjectToIgnore}");
                    return;
                }

                if (!AppConfigProvider.IsSettingExists(Constants.AppConfig.KeyNames.ProjectSearchPattern))
                {
                    Console.WriteLine($"В App.config должен быть ключ {Constants.AppConfig.KeyNames.ProjectSearchPattern}");
                    return;
                }


                Console.WriteLine(@"Введите путь до файла: (по умолчанию E:\outlook-export.csv)");
                var filePath = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    filePath = @"E:\outlook-export.csv";
                }

                Console.WriteLine("Куда сохрнаить файл? (по умолчанию - Рабочий стол)");
                var resultFilePath = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(resultFilePath))
                {
                    resultFilePath = Environment.GetFolderPath(
                             System.Environment.SpecialFolder.DesktopDirectory);
                }

                using var file = new FileWithEncoding(filePath, Constants.FileInfo.TargetEncoding);
                var reader = new CalendarCSVReader();
                var activitiesGroupedByDate = reader.ReadActivities(file.StreamReader);

                var fileName = EPPlusExcelWriter.WriteToFile(activitiesGroupedByDate, resultFilePath);
                Console.WriteLine($"Готово! Создан файл: {fileName}");

                new Process { StartInfo = new ProcessStartInfo(fileName) { UseShellExecute = true } }.Start();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                Console.ReadKey();
            }
        }
    }
}
