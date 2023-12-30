using CsvHelper;
using ExportOutlookCalendarToExcel._Common;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities;
using ExportOutlookCalendarToExcel.Library.WriteActivitiesIntoExcel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.BuildExcel
{
    /// <summary>
    /// Regulates the process of creating, writing to and opening a new Excel file.
    /// </summary>
    internal class ExcelBuilder
    {
        /// <summary>
        /// Directory path at which class works with Excel file.
        /// </summary>
        protected string _directoryPath;

        /// <param name="directoryPath">Directory path at which class works with Excel file.</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal ExcelBuilder(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath)
                || !Directory.Exists(directoryPath))
            {
                throw new ArgumentNullException(nameof(directoryPath), $"Value {directoryPath} is invalid to pass on as directory path.");
            }

            _directoryPath = directoryPath;
        }

        /// <summary>
        /// Builds excel file: create file, fill with date, open file.
        /// </summary>
        /// <param name="activities">Activities which will be recorded to Excel file.</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal void Build(ActivitiesGroupedByDateCollection activities)
        {
            Argument.NotNull(activities, nameof(activities));

            SetupLicenceContext();

            var excelFilePath = CreateExcel(activities);

            OpenExcelFile(excelFilePath);
        }

        /// <summary>
        /// Setups license context necessary for EPPluts to work.
        /// </summary>
        private void SetupLicenceContext()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Create and fill excel file with activities data.
        /// </summary>
        /// <param name="activities">Activities which will be written to excel.</param>
        /// <returns>Full file name of new excel file.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        private string CreateExcel(ActivitiesGroupedByDateCollection activities)
        {
            var fileName = EPPlusExcelWriter.WriteToFile(activities, _directoryPath);

            return fileName;
        }

        /// <summary>
        /// Open excel file.
        /// </summary>
        /// <param name="resultFilePath">Full excel file path.</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException">If <paramref name="resultFilePath"/> does not ends with excel extension.</exception>
        private void OpenExcelFile(string resultFilePath)
        {
            Argument.NotNullOrEmpty(resultFilePath, nameof(resultFilePath));

            var isFilePathEndsWithExcelExtension = resultFilePath
                                .TrimEnd(new char[] { '\\', '/' })
                                .EndsWith(Constants.FileInfo.Excel.FileExtenstion);
            if (!isFilePathEndsWithExcelExtension)
            {
                throw new ArgumentException($"Excepting excel file path to end with {Constants.FileInfo.Excel.FileExtenstion}. But to argument was passed {resultFilePath}", nameof(resultFilePath));
            }

            new Process
            {
                StartInfo = new ProcessStartInfo(resultFilePath)
                {
                    UseShellExecute = true
                }
            }.Start();
        }
    }
}
