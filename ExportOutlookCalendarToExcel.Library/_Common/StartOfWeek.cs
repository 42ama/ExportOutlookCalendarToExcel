using System;

namespace ExportOutlookCalendarToExcel.Library._Common
{

    /// <summary>
    /// Represents day which counts as a start of a week.
    /// </summary>
    internal class StartOfWeek
    {
        /// <summary>
        /// Day which counts as a start of a week.
        /// </summary>
        internal DayOfWeek Day { get; set; }

        internal StartOfWeek()
        {
            Day = DayOfWeek.Monday;
        }
    }
}