using ExportOutlookCalendarToExcel.Library._Common;
using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    /// <summary>
    /// Choose date in range from start of the current week to today.
    /// </summary>
    public class ChooseDateStrategy_WeekToToday : IChooseDateStrategy
    {
        /// <inheritdoc/>
        public DateTime From { get; private set; }

        /// <inheritdoc/>
        public DateTime To { get; private set; }

        public ChooseDateStrategy_WeekToToday()
        {
            var startOfWeek = new StartOfWeek();
            From = DateTime.Now.StartOfWeek(startOfWeek.Day).Date;
            To = DateTime.Now.Date;
        }

    }
}