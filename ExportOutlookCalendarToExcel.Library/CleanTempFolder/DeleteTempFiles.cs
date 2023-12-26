using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.CleanTempFolder
{
    /// <summary>
    /// Delete temp files from previous runs.
    /// </summary>
    internal class DeleteTempFiles
    {
        private const string EXCEL_EXTENSION = ".xlsx";
        private const string ICS_EXTENSION = ".ics";

        private readonly DirectoryInfo _directoryWithFilesInfo;
        private readonly ILogger _logger;

        internal DeleteTempFiles(DirectoryInfo directoryWithFilesInfo)
        {
            _logger = LogManager.GetCurrentClassLogger();

            _directoryWithFilesInfo = directoryWithFilesInfo;
        }

        internal void Delete()
        {
            if (!_directoryWithFilesInfo.Exists)
            {
                _logger.Warn($"Tried to clear directory, but it doesn't exists! Directory path: {_directoryWithFilesInfo}.");
                return;
            }


            var files = _directoryWithFilesInfo.GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i].FullName;
                if (filePath.EndsWith(EXCEL_EXTENSION) || filePath.EndsWith(ICS_EXTENSION))
                {
                    File.Delete(filePath);
                }
            }
        }
    }
}
