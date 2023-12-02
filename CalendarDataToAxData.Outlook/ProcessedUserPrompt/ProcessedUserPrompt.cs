using System.Windows.Forms;
using System;
using System.Diagnostics;

public class ProcessedUserPrompt
{
    public ProcessedUserPrompt(IChooseDateStrategy chooseDateStrategy)
    {
        var dialog = new ChooseDateDialog(chooseDateStrategy.From, chooseDateStrategy.To);
        var result = dialog.ShowDialog();

        ShouldContinue = result == DialogResult.OK;
        if (ShouldContinue)
        {
            From = dialog.From;
            To = dialog.To;
        }
    }
    
    public DateTime From { get; private set; }
    public DateTime To { get; private set; }
    public bool ShouldContinue {  get; private set; }
}