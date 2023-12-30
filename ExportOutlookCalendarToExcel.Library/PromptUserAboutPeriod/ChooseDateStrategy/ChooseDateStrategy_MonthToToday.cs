using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    /// <summary>
    /// Choose date in range from start of the current month to today.
    /// </summary>
    public class ChooseDateStrategy_MonthToToday : IChooseDateStrategy
    {
        /// <inheritdoc/>
        public DateTime From { get; private set; }

        /// <inheritdoc/>
        public DateTime To { get; private set; }

        public ChooseDateStrategy_MonthToToday()
        {
            From = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0, DateTimeKind.Local).Date;
            To = DateTime.Now.Date;
        }

    }
}