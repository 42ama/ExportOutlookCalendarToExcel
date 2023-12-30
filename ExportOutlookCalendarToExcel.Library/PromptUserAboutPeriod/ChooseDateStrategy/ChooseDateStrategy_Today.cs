using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    /// <summary>
    /// Choose date in range from start of today to the end of today.
    /// </summary>
    public class ChooseDateStrategy_Today : IChooseDateStrategy
    {
        /// <inheritdoc/>
        public DateTime From { get; private set; }

        /// <inheritdoc/>
        public DateTime To { get; private set; }

        public ChooseDateStrategy_Today()
        {
            From = DateTime.Now.Date;
            To = DateTime.Now.Date;
        }

    }
}