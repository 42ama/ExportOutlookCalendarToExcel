using System;

public class ChooseDateStrategy_MonthToToday : IChooseDateStrategy
{
    public DateTime From { get; private set; }
    public DateTime To { get; private set; }

    public ChooseDateStrategy_MonthToToday()
    {
        From = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0, DateTimeKind.Local).Date;
        To = DateTime.Now.Date;
    }

}