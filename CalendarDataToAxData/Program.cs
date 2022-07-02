using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Globalization;
using System.IO;
using CalendarDataToAxData.Model;
using CalendarDataToAxData.Logic;
using System.Linq;
using IronXL;
using Microsoft.Extensions.Configuration;
using CalendarDataToAxData.Extension;

namespace CalendarDataToAxData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine(@"Введите путь до файла: (по умолчанию I:\calendula.csv)");
                var filePath = Console.ReadLine();// I:\calendula.csv
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    filePath = @"I:\calendula.csv";
                }

                Console.WriteLine("Куда сохрнаить файл? (по умолчанию - Рабочий стол)");
                var resultFilePath = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(resultFilePath))
                {
                    resultFilePath = Environment.GetFolderPath(
                             System.Environment.SpecialFolder.DesktopDirectory);
                }

                var activitiesGroupedByDate = CalendarCSVReader.ReadActivities(filePath);
                var fileName = EPPlusExcelWriter.WriteToFile(activitiesGroupedByDate, resultFilePath);
                Console.WriteLine($"Готово! Создан файл: {fileName}");
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
