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
    public class ExportAndConvertCalendarProcess
    {
        private readonly DirectoryInfo _resultsDirectoryInfo;
        private readonly ILogger _logger;

        public ExportAndConvertCalendarProcess(IFilePathLocationStrategy filePathLocationStrategy, DirectoryInfo resultsDirectoryInfo)
        {
            InitializeLibrary.Initialize(filePathLocationStrategy);

            _resultsDirectoryInfo = resultsDirectoryInfo;
            _logger = LogManager.GetCurrentClassLogger();
        }

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
