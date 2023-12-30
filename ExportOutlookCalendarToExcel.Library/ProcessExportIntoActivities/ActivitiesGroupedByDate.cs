using System;
using System.Collections.Generic;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Activities grouped by date.
    /// </summary>
    public class ActivitiesGroupedByDate
    {
        /// <summary>
        /// Activities.
        /// </summary>
        internal IEnumerable<Activity> Activities { get; set; }

        /// <summary>
        /// Date.
        /// </summary>
        internal DateTime Date { get; set; }
    }
}
