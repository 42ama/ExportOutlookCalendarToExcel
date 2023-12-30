using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ExportCalendarFromOutlook
{
    /// <summary>
    /// Information about exported from outlook data.
    /// </summary>
    public class ExportedData
    {
        /// <summary>
        /// Path to exported file.
        /// </summary>
        public string FilePath { get; }

        /// <summary>
        /// Parent directory of exported file
        /// </summary>
        public string ParentDirectoryFilePath { get; }

        /// <param name="filePath">Path to exported file.</param>
        public ExportedData(string filePath)
        {
            Argument.NotNullOrEmpty(filePath, nameof(filePath));

            FilePath = filePath;
            ParentDirectoryFilePath = filePath.Substring(0, filePath.LastIndexOf('\\'));
        }
    }
}
