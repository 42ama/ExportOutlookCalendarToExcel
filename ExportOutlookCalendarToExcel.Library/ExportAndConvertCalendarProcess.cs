using ExportOutlookCalendarToExcel._Common;
using ExportOutlookCalendarToExcel.Library._Common.FilePathLocationStrategy;
using ExportOutlookCalendarToExcel.Library.BuildExcel;
using ExportOutlookCalendarToExcel.Library.CleanTempFolder;
using ExportOutlookCalendarToExcel.Library.ExportCalendarFromOutlook;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library
{
    /// <summary>
    /// Main process which export data from Outlook and converts it into excel.
    /// </summary>
    public class ExportAndConvertCalendarProcess
    {
        /// <summary>
        /// Information about results directory.
        /// </summary>
        private readonly DirectoryInfo _resultsDirectoryInfo;
        
        /// <summary>
        /// Logger.
        /// </summary>
        private readonly ILogger _logger;

        /// <param name="filePathLocationStrategy">Strategy to locate path to results directory.</param>
        /// <param name="resultsDirectoryInfo">Information about results directory.</param>
        public ExportAndConvertCalendarProcess(IFilePathLocationStrategy filePathLocationStrategy, DirectoryInfo resultsDirectoryInfo)
        {
            Argument.NotNull(filePathLocationStrategy, nameof(filePathLocationStrategy));
            Argument.NotNull(resultsDirectoryInfo, nameof(resultsDirectoryInfo));

            InitializeLibrary.Initialize(filePathLocationStrategy);

            _resultsDirectoryInfo = resultsDirectoryInfo;
            _logger = LogManager.GetCurrentClassLogger();
        }

        /// <summary>
        /// Process calendar export into excel files.
        /// </summary>
        /// <param name="exporter">Outlook exporter.</param>
        /// <param name="from">Take calendar records from date.</param>
        /// <param name="to">Take calendar records to date.</param>
        public void Process(IResultsExporter exporter, DateTime from, DateTime to)
        {
            try
            {
                var deleteTempFiles = new DeleteTempFiles(_resultsDirectoryInfo);
                deleteTempFiles.Delete();
            }
            catch (IOException ex)
            {
                _logger.Warn($"Can't delete temp files because one of them is open. Path: {_resultsDirectoryInfo}");
            }

            ExportedData exported;
            try
            {
                exported = exporter.Export(_resultsDirectoryInfo, from, to);
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, $"Export was unsuccessful.");
                return;
            }

            var icsActivitiesProcessor = new ICSActivitiesProcessor(exported.FilePath);
            var activities = icsActivitiesProcessor.ReadActivities();

            var excelBuilder = new ExcelBuilder(exported.ParentDirectoryFilePath);
            excelBuilder.Build(activities);
        }
    }
}
