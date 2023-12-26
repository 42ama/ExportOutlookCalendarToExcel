using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library._Common.FilePathLocationStrategy
{
    /// <summary>
    /// Strategy for locating directory and log directory for library temp files.
    /// </summary>
    public interface IFilePathLocationStrategy
    {
        /// <summary>
        /// Get directory for temp files location.
        /// </summary>
        /// <returns>Path to directory with temp files.</returns>
        string GetDirLocation();

        /// <summary>
        /// Get nlog filename string for general logging purposes.
        /// </summary>
        /// <returns>Nlog filename.</returns>
        string GetNLogLoggerFilename();
    }
}
