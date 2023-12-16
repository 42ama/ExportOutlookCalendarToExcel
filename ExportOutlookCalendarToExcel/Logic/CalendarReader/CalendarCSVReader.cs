using ExportOutlookCalendarToExcel.Model;
using ExportOutlookCalendarToExcel.Common;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel.Logic.CalendarReader
{
    /// <summary>
    /// Чтение данных календаря из CSV файла
    /// </summary>
#pragma warning disable S101 // Types should be named in PascalCase
    public class CalendarCSVReader : ICalendarReader
#pragma warning restore S101 // Types should be named in PascalCase
    {
        /// <summary>
        /// Прочитать Активности из csv-файла.
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        /// <returns>Коллекция активностей</returns>
        public ActivitiesDateCollection ReadActivities(TextReader reader)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Encoding = Constants.FileInfo.TargetEncoding
            };
            using var csv = new CsvReader(reader, config);
            csv.Context.RegisterClassMap<CalendarCSVClassMap>();

            var calendars = csv.GetRecords<CalendarCSV>();
            var activities = calendars.Select(calendar => new Activity(calendar)).ToList();
            var activitiesDateCollection = new ActivitiesDateCollection(activities);
            activitiesDateCollection.TryFillEmptyProject();
            return activitiesDateCollection;
        }
    }
}
