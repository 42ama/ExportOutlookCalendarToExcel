using CalendarDataToAxData.Common;
using CalendarDataToAxData.Outlook;
using Microsoft.Office.Interop.Outlook;
using System;

public class AXCommandExportCSV
{
    public string FilePath { get; private set; }    
    public AXCommandExportCSV(DateTime from, DateTime to)
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
}