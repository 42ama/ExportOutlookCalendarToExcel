using System;
using System.Collections.Generic;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Активности и даты.
    /// </summary>
    public class ActivitiesGroupedByDate
    {
        /// <summary>
        /// Активности.
        /// </summary>
        internal IEnumerable<Activity> Activities { get; set; }

        /// <summary>
        /// Дата.
        /// </summary>
        internal DateTime Date { get; set; }
    }
}
