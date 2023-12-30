using ExportOutlookCalendarToExcel._Common;
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
    /// Delete iCalendar and Excel files in directory.
    /// </summary>
    internal class DeleteTempFiles
    {
        /// <summary>
        /// Information about directory in which temp files is stored.
        /// </summary>
        private readonly DirectoryInfo _directoryWithFilesInfo;

        /// <param name="directoryWithFilesInfo">Directory in which files will be deleted</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal DeleteTempFiles(DirectoryInfo directoryWithFilesInfo)
        {
            Argument.NotNull(directoryWithFilesInfo, nameof(directoryWithFilesInfo));            

            _directoryWithFilesInfo = directoryWithFilesInfo;
        }

        /// <summary>
        /// Delete iCalendar and Excel files in directory.
        /// </summary>
        internal void Delete()
        {
            if (!_directoryWithFilesInfo.Exists)
            {
                var logger = LogManager.GetCurrentClassLogger();
                logger.Warn($"Tried to clear directory, but it doesn't exists! Directory path: {_directoryWithFilesInfo}.");
                return;
            }

            var files = _directoryWithFilesInfo.GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i].FullName;
                if (filePath.EndsWith(Constants.FileInfo.Excel.FileExtenstion)
                    || filePath.EndsWith(Constants.FileInfo.ICalendar.FileExtenstion))
                {
                    File.Delete(filePath);
                }
            }
        }
    }
}
