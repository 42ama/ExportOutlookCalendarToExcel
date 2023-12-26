// Ignore Spelling: App

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library._Common.FilePathLocationStrategy
{
    /// <summary>
    /// Strategy for locating files in AppData\Local
    /// </summary>
    public class AppDataLocalFilePathLocationStrategy : IFilePathLocationStrategy
    {
        /// <inheritdoc/>
        public string GetDirLocation()
        {
            return $@"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\ExportOutlookCalendarToExcel";
        }

        /// <inheritdoc/>
        public string GetNLogLoggerFilename()
        {
            return "${specialfolder:folder=LocalApplicationData}\\ExportOutlookCalendarToExcel\\Logs\\${shortdate}.log";
        }
    }
}
