using ExportOutlookCalendarToExcel._Common;
using ExportOutlookCalendarToExcel.Library._Common.FilePathLocationStrategy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.BuildExcel
{
    /// <summary>
    /// Prepares directory to store export results. Information about directory stored in <c>ResultsDirectoryInfo</c> property.
    /// </summary>
    public class PrepareResultsDirCommand
    {
        /// <summary>
        /// Information about directory in which results will be stored.
        /// </summary>
        public DirectoryInfo ResultsDirectoryInfo { get; private set; }

        /// <summary>
        /// On construction creates directory for export results. Directory path is chosen by <paramref name="filePathLocationStrategy"/>. Directory information can be read from <c>ResultsDirectoryInfo</c>.
        /// </summary>
        /// <param name="filePathLocationStrategy">Strategy which determines where new directory for results will be created.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public PrepareResultsDirCommand(IFilePathLocationStrategy filePathLocationStrategy)
        {
            Argument.NotNull(filePathLocationStrategy, nameof(filePathLocationStrategy));

            var resultsFolderPath = filePathLocationStrategy.GetDirLocation();
            var resultsDirectoryInfo = new DirectoryInfo(resultsFolderPath);
            if (!resultsDirectoryInfo.Exists)
            {
                resultsDirectoryInfo.Create();
            }

            ResultsDirectoryInfo = resultsDirectoryInfo;
        }
    }
}
