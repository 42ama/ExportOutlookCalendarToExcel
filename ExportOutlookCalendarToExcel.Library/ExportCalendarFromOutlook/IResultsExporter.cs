using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ExportCalendarFromOutlook
{
    /// <summary>
    /// Export calendar records and provides information about exported file.
    /// </summary>
    public interface IResultsExporter
    {
        /// <summary>
        /// Export calendar records in period set by <paramref name="from"/> <paramref name="to"/> arguments. Exported data will be stored in <paramref name="resultsDirectory"/>
        /// </summary>
        /// <param name="resultsDirectory">Directory in which exported data will be stored.</param>
        /// <param name="from">Left boundary of period from which calendar records will be exported.</param>
        /// <param name="to">Right boundary of period from which calendar records will be exported.</param>
        /// <returns>Information about exported file.</returns>
        ExportedData Export(DirectoryInfo resultsDirectory, DateTime from, DateTime to);
    }
}
