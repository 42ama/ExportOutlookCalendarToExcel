using ExportOutlookCalendarToExcel.Common;
using ExportOutlookCalendarToExcel.Outlook;
using Microsoft.Office.Interop.Outlook;
using System;

/// <summary>
/// Export calendar as ICal. 
/// </summary>
public class ExportICalCommand
{
    /// <summary>
    /// Export calendar as ICal. After class construction you can access ICal file at <c>FilePath</c> property location.
    /// </summary>
    /// <param name="from">From date</param>
    /// <param name="to">To date</param>
    public ExportICalCommand(DateTime from, DateTime to)
    {
        var outlookNamespace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
        var folder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        var calendarExporter = folder.GetCalendarExporter();
        calendarExporter.CalendarDetail = OlCalendarDetail.olFullDetails;
        calendarExporter.StartDate = from;
        calendarExporter.EndDate = to;

        FilePath = Constants.FileInfo.ICS.FilePath;
        calendarExporter.SaveAsICal(FilePath);
    }

    /// <summary>
    /// Location of ICal file.
    /// </summary>
    public string FilePath { get; private set; }
}