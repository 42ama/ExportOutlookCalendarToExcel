using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    /// <summary>
    /// Strategy by which date period should be chosen.
    /// </summary>
    public interface IChooseDateStrategy
    {
        /// <summary>
        /// Left boundary of period.
        /// </summary>
        DateTime From { get; }

        /// <summary>
        /// Right boundary of period.
        /// </summary>
        DateTime To { get; }
    }
}