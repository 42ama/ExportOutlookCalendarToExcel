using ExportOutlookCalendarToExcel.Library._Common.FilePathLocationStrategy;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library
{
    /// <summary>
    /// Initializes addition behavior for library to work.
    /// </summary>
    internal static class InitializeLibrary
    {
        /// <summary>
        /// Has library been already initialized?
        /// </summary>
        private static bool _initialized = false;

        /// <summary>
        /// Initializes addition behavior for library to work.
        /// </summary>
        /// <param name="filepathLocationStrategy">Strategy to find folder for logs.</param>
        internal static void Initialize(IFilePathLocationStrategy filepathLocationStrategy)
        {
            if (_initialized) { return; }

            LogManager.Setup().LoadConfiguration(builder => {
                builder.ForLogger().FilterMinLevel(LogLevel.Warn).WriteToFile(fileName: filepathLocationStrategy.GetNLogLoggerFilename());
            });

            _initialized = true;
        }
    }
}
