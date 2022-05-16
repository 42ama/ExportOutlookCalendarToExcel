using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Globalization;
using System.IO;
using CalendarDataToAxData.Model;
using CalendarDataToAxData.Logic;
using System.Linq;
using IronXL;

namespace CalendarDataToAxData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите путь до файла:");
            var filePath = Console.ReadLine();// I:\calendula.csv
            Console.WriteLine("Куда сохрнаить файл? (по умолчанию - Рабочий стол");
            var resultFilePath = Console.ReadLine(); 
            if (string.IsNullOrWhiteSpace(resultFilePath))
            {
                resultFilePath = Environment.GetFolderPath(
                         System.Environment.SpecialFolder.DesktopDirectory);
            }

            var activitiesGroupedByDate = CalendarReader.ReadActivities(filePath);
            ExcelWriter.Execute(activitiesGroupedByDate, resultFilePath);
        }
    }
}
