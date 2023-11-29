using System.Windows.Forms;
using System;
using System.Diagnostics;

public class AXCommandUserPrompt
{
    public AXCommandUserPrompt()
    {
        var dialog = new InputDialog();
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