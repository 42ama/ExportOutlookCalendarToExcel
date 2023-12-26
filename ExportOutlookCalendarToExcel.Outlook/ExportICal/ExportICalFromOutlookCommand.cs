using ExportOutlookCalendarToExcel.Library.ExportCalendarFromOutlook;
using ExportOutlookCalendarToExcel.Outlook;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;

/// <summary>
/// Export calendar as ICal. 
/// </summary>
public class ExportICalFromOutlookCommand : IResultsExporter
{
    /// <summary>
    /// Export calendar as ICal. 
    /// </summary>
    /// <param name="resultsDirectory">Directory to which export is carried out.</param>
    /// <param name="from">Calendar left boundary.</param>
    /// <param name="to">Calendar right boundary.</param>
    /// <returns>Information about expored file.</returns>
    public ExportedData Export(DirectoryInfo resultsDirectory, DateTime from, DateTime to)
    {
        var outlookNamespace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
        var folder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
        var calendarExporter = folder.GetCalendarExporter();
        calendarExporter.CalendarDetail = OlCalendarDetail.olFullDetails;
        calendarExporter.StartDate = from;
        calendarExporter.EndDate = to;

        var filePath = resultsDirectory.FullName + @"\outlook-export.ics";
        calendarExporter.SaveAsICal(filePath);

        return new ExportedData(filePath);
    }
}