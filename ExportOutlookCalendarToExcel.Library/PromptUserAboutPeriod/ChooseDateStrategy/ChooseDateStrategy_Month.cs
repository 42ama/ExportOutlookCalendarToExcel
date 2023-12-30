using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    /// <summary>
    /// Choose date in range from start of the current month to the end of the current month.
    /// </summary>
    public class ChooseDateStrategy_Month : IChooseDateStrategy
    {
        /// <inheritdoc/>
        public DateTime From { get; private set; }

        /// <inheritdoc/>
        public DateTime To { get; private set; }

        public ChooseDateStrategy_Month()
        {
            From = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0, DateTimeKind.Local).Date;
            To = From.AddMonths(1).AddDays(-1).Date;
        }

    }
}