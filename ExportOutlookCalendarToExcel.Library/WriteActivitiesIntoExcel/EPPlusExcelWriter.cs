using ExportOutlookCalendarToExcel._Common;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.WriteActivitiesIntoExcel
{
    /// <summary>
    /// Write activities data to Excel file. Uses EPPlus library.
    /// </summary>
    internal static class EPPlusExcelWriter
    {
        /// <summary>
        /// Write activities data to Excel file.
        /// </summary>
        /// <param name="activitiesDateCollection">Activities grouped by date.</param>
        /// <param name="resultDirPath">File path of directory in which Excel will be stored.</param>
        /// <returns>Path to excel file.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        internal static string WriteToFile(ActivitiesGroupedByDateCollection activitiesDateCollection, string resultDirPath)
         {
            Argument.NotNull(activitiesDateCollection, nameof(activitiesDateCollection));
            Argument.NotNullOrEmpty(resultDirPath, nameof(resultDirPath));

            var sortedDates = activitiesDateCollection.GetSortedDates();
            var fileName = CreateFileDestination(sortedDates, resultDirPath);
            using (var package = new ExcelPackage(fileName))
            {
                foreach (var activitiesWithDate in activitiesDateCollection)
                {
                    var sheetName = activitiesWithDate.Date.ToShortDateString();

                    if (package.Workbook.Worksheets.Any(sheet => sheet.Name == sheetName))
                    {
                        package.Workbook.Worksheets.Delete(sheetName);
                    }

                    var currentDateSheet = package.Workbook.Worksheets.Add(sheetName);

                    SetHeaderColumns(currentDateSheet);
                    var lastRowIndex = ProcessActivities(currentDateSheet, activitiesWithDate.Activities);
                    SetColumnStyles(currentDateSheet);
                    AddDurationSumFormula(currentDateSheet, lastRowIndex);
                }

                package.Save();

                return fileName;
            }                
        }

        /// <summary>
        /// Create path to the file.
        /// </summary>
        /// <param name="sortedDates">Dates by which activities is grouped by.</param>
        /// <param name="resultDirPath">Path to the directory in which excel file will be stored.</param>
        /// <returns>Path to the excel file.</returns>
        private static string CreateFileDestination(IOrderedEnumerable<DateTime> sortedDates, string resultDirPath)
        {
            Argument.NotNull(sortedDates, nameof(sortedDates));
            Argument.NotNullOrEmpty(resultDirPath, nameof(resultDirPath));

            var firstDate = sortedDates.First().ToShortDateString();
            var lastDate = sortedDates.Last().ToShortDateString();

            var fileName = $"{firstDate}_{lastDate}.{Constants.FileInfo.Excel.FileExtenstion}";

            var filePath = Path.Combine(resultDirPath, fileName);

            return filePath;
        }

        /// <summary>
        /// Set header column values.
        /// </summary>
        /// <param name="sheet">Excel sheets.</param>
        private static void SetHeaderColumns(ExcelWorksheet sheet)
        {
            Argument.NotNull(sheet, nameof(sheet));
            
            sheet.SetCellValue(Constants.Excel.Group.Letter,
                                Constants.Excel.Header.RowNumber,
                                Constants.Excel.Group.Name);

            sheet.SetCellValue(Constants.Excel.Subject.Letter,
                                Constants.Excel.Header.RowNumber,
                                Constants.Excel.Subject.Name);

            sheet.SetCellValue(Constants.Excel.Duration.Letter,
                                Constants.Excel.Header.RowNumber,
                                Constants.Excel.Duration.Name);
        }

        /// <summary>
        /// Set column styles
        /// </summary>
        /// <param name="sheet">Excel sheet.</param>
        private static void SetColumnStyles(ExcelWorksheet sheet)
        {
            Argument.NotNull(sheet, nameof(sheet));

            sheet.Columns[Constants.Excel.Subject.ColumnNumber].Width = Constants.Excel.Subject.ColumnLength;
            sheet.Columns[Constants.Excel.Subject.ColumnNumber].Style.WrapText = true;
        }

        /// <summary>
        /// Add duration sum column after all columns.
        /// </summary>
        /// <param name="sheet">Excel sheet.</param>
        /// <param name="lastRowIndex">Index of last row, after which sum formula will be added.</param>
        private static void AddDurationSumFormula(ExcelWorksheet sheet, int lastRowIndex)
        {
            Argument.NotNull(sheet, nameof(sheet));
            Argument.Require(lastRowIndex > 0, $"Row index should ({nameof(lastRowIndex)}) should be greater than 0.");

            var sumDurationFormula = new BasicFormula
            {
                RangeStart = Constants.Excel.FirstValueRowNumber,
                RangeEnd = lastRowIndex,
                Letter = Constants.Excel.Duration.Letter,
                FormulaOperation = Constants.FormulaOperations.Sum
            };
            var formula = sumDurationFormula.GetFormula();

            sheet.SetFormula(Constants.Excel.Duration.Letter, lastRowIndex + 1, formula);
        }

        /// <summary>
        /// Write activities into excel sheet.
        /// </summary>
        /// <param name="sheet">Excel sheet.</param>
        /// <param name="activities">Activities.</param>
        /// <returns>Index of last used row.</returns>
        private static int ProcessActivities(ExcelWorksheet sheet, IEnumerable<Activity> activities)
        {
            Argument.NotNull(sheet, nameof(sheet));
            Argument.NotNull(activities, nameof(activities));

            var indexToRecord = Constants.Excel.Header.RowNumber;

            foreach (var activity in activities)
            {
                indexToRecord++;
                
                sheet.SetCellValue(Constants.Excel.Group.Letter,
                                indexToRecord,
                                activity.Group);
                sheet.SetCellValue(Constants.Excel.Subject.Letter,
                                indexToRecord,
                                activity.Subject);
                sheet.SetCellValue(Constants.Excel.Duration.Letter,
                                indexToRecord,
                                activity.Duration);
            }

            return indexToRecord;
        }
    }
}
