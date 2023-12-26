using System;

namespace ExportOutlookCalendarToExcel.Library._Common
{

    /// <summary>
    /// Represents day which counts as a start of a week.
    /// </summary>
    public class StartOfWeek
    {
        /// <summary>
        /// Day which counts as a start of a week.
        /// </summary>
        public DayOfWeek Day { get; set; }

        public StartOfWeek()
        {
            Day = DayOfWeek.Monday;
        }
    }
}