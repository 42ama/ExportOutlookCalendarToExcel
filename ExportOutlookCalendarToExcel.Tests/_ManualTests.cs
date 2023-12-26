using System.Diagnostics;
using System.Configuration;
using System;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport;

namespace ExportOutlookCalendarToExcel.Tests
{
    [TestClass]
    internal class _ManualTests
    {
        [TestMethod]
        internal void ReaderCalendarFromPersonalAndConvert()
        {
            var readFromPath = @"C:\Users\Maxim.Alonov\AppData\Local\ExportOutlookCalendarToExcel\outlook-export.ics";

            using (var textReader = new StreamReader(readFromPath))
            {
                var calendarICSReader = new CalendarICSReader();
                var activities = calendarICSReader.ReadActivities(textReader);

                ; // breakpoint here
            }
        }

        
    }
}