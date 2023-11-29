using CalendarDataToAxData.Common;
using CalendarDataToAxData.Logic.CalendarReader;
using CalendarDataToAxData.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarDataToAxData.Logic.ResultBuilder
{
    internal class CSVResultBuilder : AbstractResultBuilder
    {
        public override string GetFilePathToReadFrom()
        {
            Console.WriteLine(@"Введите путь до файла: (по умолчанию E:\outlook-export.csv)");
            var filePath = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                filePath = @"E:\outlook-export.csv";
            }

            return filePath;
        }

        public override TextReader GetTextReader(string readFromPath)
        {
            var file = new FileWithEncoding(readFromPath, Constants.FileInfo.TargetEncoding);
            return file.StreamReader;
        }

        public override ActivitiesDateCollection ReadActivities(TextReader reader)
        {
            var calendarReader = new CalendarCSVReader();
            return calendarReader.ReadActivities(reader);
        }
        public override string GetResultFileDirPath()
        {
            Console.WriteLine("Куда сохрнаить файл? (по умолчанию - Рабочий стол)");
            var resultFilePath = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(resultFilePath))
            {
                resultFilePath = Environment.GetFolderPath(
                         System.Environment.SpecialFolder.DesktopDirectory);
            }

            return resultFilePath;
        }
    }
}
