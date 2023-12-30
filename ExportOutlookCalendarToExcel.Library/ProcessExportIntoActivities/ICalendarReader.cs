using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Read from file with activities into activities collection.
    /// </summary>
    public interface ICalendarReader
    {
        /// <summary>
        /// Read activities from file into collection.
        /// </summary>
        /// <param name="reader">Reader to file with activities.</param>
        /// <returns>Collection of activities.</returns>
        ActivitiesGroupedByDateCollection ReadActivities(TextReader reader);
    }
}
