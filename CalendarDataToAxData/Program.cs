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
            var filePath = Console.ReadLine(); // I:\calendula.csv

            var activitiesGroupedByDate = CalendarReader.ReadActivities(filePath);
            ExcelWriter.Execute(activitiesGroupedByDate);
        }
    }
}
