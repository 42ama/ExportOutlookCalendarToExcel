using System;

namespace ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy
{
    public class ChooseDateStrategy_Month : IChooseDateStrategy
    {
        public DateTime From { get; private set; }
        public DateTime To { get; private set; }

        public ChooseDateStrategy_Month()
        {
            From = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0, DateTimeKind.Local).Date;
            To = From.AddMonths(1).AddDays(-1).Date;
        }

    }
}