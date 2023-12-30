using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Read file and return as <c>ActivitiesGroupedByDateCollection</c>.
    /// </summary>
    internal abstract class AbstractActivitiesProcessor
    {
        /// <summary>
        /// File from which activities will be read.
        /// </summary>
        protected string _filePath;

        /// <param name="filePath">File from which activities will be read.</param>
        /// <exception cref="ArgumentNullException"></exception>
        internal AbstractActivitiesProcessor(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)
                || !File.Exists(filePath))
            {
                throw new ArgumentNullException(nameof(filePath), $"Value {filePath} is invalid to pass on as file path.");
            }

            _filePath = filePath;
        }

        /// <summary>
        /// Get path of file to read activities from.
        /// </summary>
        /// <returns>Path of file to read activities from.</returns>
        protected virtual string GetFilePathToReadFrom()
        {
            return _filePath;
        }

        /// <summary>
        /// Get <c>TextReader</c> to file with activities.
        /// </summary>
        /// <param name="readFromPath">File with activities.</param>
        /// <returns><c>TextReader</c> to file with activities.</returns>
        protected virtual TextReader GetTextReader(string readFromPath)
        {
            Argument.NotNullOrEmpty(readFromPath, nameof(readFromPath));

            return new StreamReader(readFromPath);
        }

        /// <summary>
        /// Read activities from file.
        /// </summary>
        /// <returns>Activities.</returns>
        internal abstract ActivitiesGroupedByDateCollection ReadActivities();
    }
}
