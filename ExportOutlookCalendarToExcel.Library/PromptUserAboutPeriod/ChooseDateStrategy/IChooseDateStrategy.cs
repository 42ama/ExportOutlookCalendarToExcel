using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    public interface IChooseDateStrategy
    {
        DateTime From { get; }
        DateTime To { get; }
    }
}