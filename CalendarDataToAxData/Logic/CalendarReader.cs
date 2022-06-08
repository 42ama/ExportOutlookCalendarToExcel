using CalendarDataToAxData.Model;
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
    public static class CalendarReader
    {
        private static ISet<string> _subjectToIgnore = new HashSet<string>
        {
            "Обед",
            "Блок",
            "Встреча. Блок"
        };

        public static IEnumerable<IGrouping<string, Activity>> ReadActivities(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException($"'{nameof(filePath)}' cannot be null or empty.", nameof(filePath));
            }

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                NewLine = Environment.NewLine,
            };

            using var reader = new StreamReader(filePath);
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

            csv.Context.RegisterClassMap<CalendarCSVClassMap>();
            var calendars = csv.GetRecords<CalendarCSV>();
            var activitiesGroupedByDate = calendars
                .Select(calendar => new Activity(calendar))
                .Where(activity => !_subjectToIgnore.Contains(activity.Subject))
                .Where(activity => activity.Duration > 0)
                .OrderByDescending(activity => activity.Project)
                .GroupBy(a => a.Date)
                .ToList();

            return activitiesGroupedByDate;
        }
    }
}
