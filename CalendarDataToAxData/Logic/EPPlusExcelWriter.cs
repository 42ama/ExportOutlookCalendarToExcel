using CalendarDataToAxData.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CalendarDataToAxData.Logic
{
    internal class EPPlusExcelWriter
    {

        const char projectLetter = 'A';
        const char subjectLetter = 'B';
        const int subjectColumnNumer = 2;
        const char durationLetter = 'C';
        const char durationToCalculateLetter = 'D';
        const int headerRowNumber = 1;
        const int SUBJECT_COLUMN_LENGTH = 122;

        public static string Execute(IEnumerable<IGrouping<string, Activity>> activitiesGroupedByDate, string resultFilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (activitiesGroupedByDate is null)
            {
                throw new ArgumentNullException(nameof(activitiesGroupedByDate));
            }

            if (string.IsNullOrEmpty(resultFilePath))
            {
                throw new ArgumentException($"'{nameof(resultFilePath)}' cannot be null or empty.", nameof(resultFilePath));
            }
            

            var dates = activitiesGroupedByDate.Select(i => i.Key);
            var fileName = CreateFileName(dates, resultFilePath);
            using (var package = new ExcelPackage(fileName))
            {
                foreach (var activityGroupedByDate in activitiesGroupedByDate)
                {
                    var date = activityGroupedByDate.Key;

                    if (package.Workbook.Worksheets.Any(sheet => sheet.Name == date))
                    {
                        package.Workbook.Worksheets.Delete(date);
                    }

                    var xlsSheet = package.Workbook.Worksheets.Add(date);

                    SetHeaderColumns(xlsSheet);

                    var indexAfterLastRow = ProcessActivities(xlsSheet, activityGroupedByDate);

                    SetColumnStyles(xlsSheet);
                    AddAggregationColumns(xlsSheet, indexAfterLastRow);
                }


                package.Save();
            }
            return fileName;
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

        public static void SetHeaderColumns(ExcelWorksheet sheet)
        {
            sheet.SetValue($"{projectLetter}{headerRowNumber}", "Проект");
            sheet.SetValue($"{subjectLetter}{headerRowNumber}", "Тема");
            sheet.SetValue($"{durationLetter}{headerRowNumber}", "Длительность");
            sheet.SetValue($"{durationToCalculateLetter}{headerRowNumber}", "Длительность для рассчетов");
        }

        public static void SetColumnStyles(ExcelWorksheet sheet)
        {
            sheet.Columns[subjectColumnNumer].Width = SUBJECT_COLUMN_LENGTH;
            sheet.Columns[subjectColumnNumer].Style.WrapText = true;
        }

        public static void AddAggregationColumns(ExcelWorksheet sheet, int indexAfterLastRow)
        {
            var rangeString = $"{durationToCalculateLetter}2:{durationToCalculateLetter}{(indexAfterLastRow - 1).ToString()}";
            sheet.SetValue($"{durationToCalculateLetter}{indexAfterLastRow}", $"=СУММ({rangeString})");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="activities"></param>
        /// <param name="indexAfterLastRow"></param>
        /// <returns>indexAfterLastRow</returns>
        public static int ProcessActivities(ExcelWorksheet sheet, IEnumerable<Activity> activities)
        {
            var index = 2;

            foreach (var activity in activities)
            {
                sheet.SetValue($"{projectLetter}{index}", activity.Project);
                sheet.SetValue($"{subjectLetter}{index}", activity.Subject);
                sheet.SetValue($"{durationLetter}{index}", activity.DurationFormated);
                sheet.SetValue($"{durationToCalculateLetter}{index}", activity.Duration);

                index++;
            }

            return index;
        }
    }
}
