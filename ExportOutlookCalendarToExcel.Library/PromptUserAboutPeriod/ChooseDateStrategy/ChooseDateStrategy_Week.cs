using ExportOutlookCalendarToExcel.Library._Common;
using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    public class ChooseDateStrategy_Week : IChooseDateStrategy
    {
        public DateTime From { get; private set; }
        public DateTime To { get; private set; }

        public ChooseDateStrategy_Week()
        {
            var startOfWeek = new StartOfWeek();
            From = DateTime.Now.StartOfWeek(startOfWeek.Day).Date;
            To = DateTime.Now.EndOfWeek(startOfWeek.Day).Date;
        }

    }
}