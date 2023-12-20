using ExportOutlookCalendarToExcel.Common;
using ExportOutlookCalendarToExcel.Outlook;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using ExportOutlookCalendarToExcel.Library.Logic;

/// <summary>
/// Export calendar as ICal. 
/// </summary>
public class ExportICalFromOutlookCommand : IResultsExporter
{
    public ExportedData Export(DirectoryInfo resultsDirectory, DateTime from, DateTime to)
    {
        try
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
        catch (System.Exception)
        {
            return new ExportedData();
        }
        
    }
}