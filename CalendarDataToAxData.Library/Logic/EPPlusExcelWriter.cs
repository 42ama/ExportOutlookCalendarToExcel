using CalendarDataToAxData.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CalendarDataToAxData.Common;
using CalendarDataToAxData.Extension;

namespace CalendarDataToAxData.Logic
{
    /// <summary>
    /// Запись в Excel файл, библиотекка EPPlus.
    /// </summary>
    public static class EPPlusExcelWriter
    {
        /// <summary>
        /// Выполнить запись коллекции Активностей в Файл
        /// </summary>
        /// <param name="activitiesDateCollection">Актинвости сгруппированные по дате.</param>
        /// <param name="resultFilePath">Папка для конечного местоположения.</param>
        /// <returns>Путь до файла.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static string WriteToFile(ActivitiesDateCollection activitiesDateCollection, string resultFilePath)
         {
            Argument.NotNull(activitiesDateCollection, nameof(activitiesDateCollection));
            Argument.NotNullOrEmpty(resultFilePath, nameof(resultFilePath));

            var sortedDates = activitiesDateCollection.GetSortedDates();
            var fileName = CreateFileDestination(sortedDates, resultFilePath);
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
                    AddAggregationColumns(currentDateSheet, lastRowIndex);
                }

                package.Save();

                return fileName;
            }                
        }

        /// <summary>
        /// Создает путь до файла.
        /// </summary>
        /// <param name="activitiesDateCollection">Коллекция активностей и дат</param>
        /// <param name="resultFilePath">Путь до папки с новым файлом</param>
        /// <returns>Путь до файла</returns>
        private static string CreateFileDestination(IOrderedEnumerable<DateTime> sortedDates, string resultFilePath)
        {
            Argument.NotNull(sortedDates, nameof(sortedDates));
            Argument.NotNullOrEmpty(resultFilePath, nameof(resultFilePath));

            var firstDate = sortedDates.First().ToShortDateString();
            var lastDate = sortedDates.Last().ToShortDateString();

            var fileName = $"{firstDate}_{lastDate}.xlsx";

            var filePath = Path.Combine(resultFilePath, fileName);

            return filePath;
        }

        /// <summary>
        /// Установить колонки-заголовки.
        /// </summary>
        /// <param name="sheet">Лист excel.</param>
        private static void SetHeaderColumns(ExcelWorksheet sheet)
        {
            Argument.NotNull(sheet, nameof(sheet));

            sheet.SetCellValue(Constants.Excel.Project.Letter,
                                Constants.Excel.Header.RowNumber,
                                Constants.Excel.Project.Name);

            sheet.SetCellValue(Constants.Excel.Subject.Letter,
                                Constants.Excel.Header.RowNumber,
                                Constants.Excel.Subject.Name);

            sheet.SetCellValue(Constants.Excel.Duration.Letter,
                                Constants.Excel.Header.RowNumber,
                                Constants.Excel.Duration.Name);
        }

        /// <summary>
        /// Установить стили колонок.
        /// </summary>
        /// <param name="sheet">Лист excel.</param>
        private static void SetColumnStyles(ExcelWorksheet sheet)
        {
            Argument.NotNull(sheet, nameof(sheet));

            sheet.Columns[Constants.Excel.Subject.ColumnNumber].Width = Constants.Excel.Subject.ColumnLength;
            sheet.Columns[Constants.Excel.Subject.ColumnNumber].Style.WrapText = true;
        }

        /// <summary>
        /// Добавить агрегирующую колонку с суммой по длительности.
        /// </summary>
        /// <param name="sheet">Лист excel.</param>
        private static void AddAggregationColumns(ExcelWorksheet sheet, int lastRowIndex)
        {
            Argument.NotNull(sheet, nameof(sheet));
            Argument.Require(lastRowIndex > 0, $"Для индекса {nameof(lastRowIndex)} ожидается значение больше 0.");

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
        /// Записать активности в лист excel.
        /// </summary>
        /// <param name="sheet">Лист excel.</param>
        /// <param name="activities">Коллекция активностей</param>
        /// <returns>Номер крайней использованной строчки.</returns>
        private static int ProcessActivities(ExcelWorksheet sheet, IEnumerable<Activity> activities)
        {
            Argument.NotNull(sheet, nameof(sheet));
            Argument.NotNull(activities, nameof(activities));

            var indexToRecord = Constants.Excel.Header.RowNumber;

            foreach (var activity in activities)
            {
                indexToRecord++;
                
                sheet.SetCellValue(Constants.Excel.Project.Letter,
                                indexToRecord,
                                activity.Project);
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
