using CalendarDataToAxData.Extension;
using CalendarDataToAxData.Model;
using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CalendarDataToAxData.Logic
{
    public static class ExcelWriter
    {
        const char subjectLetter = 'A';
        const char durationLetter = 'B';
        const char durationToCalculateLetter = 'C';
        const int SUBJECT_COLUMN_LENGTH = 25500;

        public static void Execute(IEnumerable<IGrouping<string, Activity>> activitiesGroupedByDate, string resultFilePath)
        {
            if (activitiesGroupedByDate is null)
            {
                throw new ArgumentNullException(nameof(activitiesGroupedByDate));
            }

            if (string.IsNullOrEmpty(resultFilePath))
            {
                throw new ArgumentException($"'{nameof(resultFilePath)}' cannot be null or empty.", nameof(resultFilePath));
            }

            var xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            xlsxWorkbook.Metadata.Author = "CalendarDataToAxData";
            var dates = activitiesGroupedByDate.Select(i => i.Key);
            

            foreach (var activityGroupedByDate in activitiesGroupedByDate)
            {
                var date = activityGroupedByDate.Key;

                var xlsSheet = xlsxWorkbook.CreateWorkSheet(date);

                SetHeaderColumns(xlsSheet);

                var indexAfterLastRow = ProcessActivities(xlsSheet, activityGroupedByDate);

                SetColumnStyles(xlsSheet);
                AddAggregationColumns(xlsSheet, indexAfterLastRow);
            }

            var fileName = CreateFileName(dates, resultFilePath);
            xlsxWorkbook.SaveAs(fileName);
        }

        public static string CreateFileName(IEnumerable<string> dates, string resultFilePath)
        {
            var sortedDates = dates.OrderBy(i => i).ToList();

            var firstDate = sortedDates.First();
            var lastDate = sortedDates.Last();

            var fileName = $"{firstDate}_{lastDate}.xlsx";

            var filePath = Path.Combine(resultFilePath, fileName);

            return filePath;
        }

        public static void SetHeaderColumns(WorkSheet sheet)
        {
            sheet.SetValue(subjectLetter, 1, "Тема");
            sheet.SetValue(durationLetter, 1, "Длительность");
            sheet.SetValue(durationToCalculateLetter, 1, "Длительность для рассчетов");
        }

        public static void SetColumnStyles(WorkSheet sheet)
        {
            sheet.Columns[0].Width = SUBJECT_COLUMN_LENGTH;
            sheet.Columns[0].Style.WrapText = true;
        }

        public static void AddAggregationColumns(WorkSheet sheet, int indexAfterLastRow)
        {
            var rangeString = $"{durationToCalculateLetter}2:{durationToCalculateLetter}{(indexAfterLastRow -1).ToString()}";
            var durationRange = sheet[rangeString];
            var sum = durationRange.Sum();
            sheet.SetValue(durationToCalculateLetter, indexAfterLastRow, sum);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="activities"></param>
        /// <param name="indexAfterLastRow"></param>
        /// <returns>indexAfterLastRow</returns>
        public static int ProcessActivities(WorkSheet sheet, IEnumerable<Activity> activities)
        {
            var index = 2;

            foreach (var activity in activities)
            {
                sheet.SetValue(subjectLetter, index, activity.Subject);
                sheet.SetValue(durationLetter, index, activity.DurationFormated);
                sheet.SetValue(durationToCalculateLetter, index, activity.Duration);

                index++;
            }

            return index;
        }
    }
}
