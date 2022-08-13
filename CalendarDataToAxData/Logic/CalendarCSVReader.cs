using CalendarDataToAxData.Model;
using CalendarDataToAxData.Common;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace CalendarDataToAxData.Logic
{
    /// <summary>
    /// Чтение данных календаря из CSV файла
    /// </summary>
    public static class CalendarCSVReader
    {
        /// <summary>
        /// Прочитать Активности из csv-файла.
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        /// <returns>Коллекция активностей</returns>
        public static ActivitiesDateCollection ReadActivities(string filePath)
        {
            Argument.NotNullOrEmpty(filePath, nameof(filePath));

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                NewLine = Environment.NewLine,
            };

            using var reader = new StreamReader(filePath);

            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            csv.Context.RegisterClassMap<CalendarCSVClassMap>();

            var calendars = csv.GetRecords<CalendarCSV>();
            var activities = calendars.Select(calendar => new Activity(calendar)).ToList();
            var activitiesDateCollection = new ActivitiesDateCollection(activities);
            activitiesDateCollection.TryFillEmptyProject();
            return activitiesDateCollection;
        }
    }
}
