using ExportOutlookCalendarToExcel.Common;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities;
using ExportOutlookCalendarToExcel.Logic.CalendarReader;
using ExportOutlookCalendarToExcel.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Logic.ResultBuilder
{
    internal class CSVActivitiesProcessor : AbstractActivitiesProcessor
    {
        public CSVActivitiesProcessor(string filePath) :base(filePath) { }

        protected override string GetFilePathToReadFrom()
        {
            Console.WriteLine(@"Введите путь до файла: (по умолчанию E:\outlook-export.csv)");
            var filePath = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = @"E:\outlook-export.csv";
            }

            return filePath;
        }

        protected override TextReader GetTextReader(string readFromPath)
        {
            var file = new FileWithEncoding(readFromPath, Constants.FileInfo.CSV.TargetEncoding);
            return file.StreamReader;
        }

        public override ActivitiesDateCollection ReadActivities()
        {
            var readFromPath = GetFilePathToReadFrom();
            using (var textReader = GetTextReader(readFromPath))
            {
                var calendarReader = new CalendarCSVReader();
                return calendarReader.ReadActivities(textReader);
            }
            
        }
        //protected override string GetResultFileDirPath()
        //{
        //    Console.WriteLine("Куда сохрнаить файл? (по умолчанию - Рабочий стол)");
        //    var resultFilePath = Console.ReadLine();
        //    if (string.IsNullOrWhiteSpace(resultFilePath))
        //    {
        //        resultFilePath = Environment.GetFolderPath(
        //                 System.Environment.SpecialFolder.DesktopDirectory);
        //    }

        //    return resultFilePath;
        //}
    }
}
