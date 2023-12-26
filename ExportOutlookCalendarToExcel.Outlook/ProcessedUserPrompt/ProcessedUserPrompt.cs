using System.Windows.Forms;
using System;
using System.Diagnostics;
using ExportOutlookCalendarToExcel.Library.PromptUserAboutPeriod.ChooseDateStrategy;

internal class ProcessedUserPrompt
{
    internal ProcessedUserPrompt(IChooseDateStrategy chooseDateStrategy)
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
    
    internal DateTime From { get; private set; }
    internal DateTime To { get; private set; }
    internal bool ShouldContinue {  get; private set; }
}