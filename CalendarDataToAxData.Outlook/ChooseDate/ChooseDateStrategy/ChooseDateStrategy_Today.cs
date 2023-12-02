using System;

public class ChooseDateStrategy_Today : IChooseDateStrategy
{
    public DateTime From { get; private set; }
    public DateTime To { get; private set; }

    public ChooseDateStrategy_Today()
    {
        From = DateTime.Now.Date;
        To = DateTime.Now.Date;
    }

}