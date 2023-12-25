using ExportOutlookCalendarToExcel.Logic.CalendarReader;
using System.Diagnostics;
using ExportOutlookCalendarToExcel.Logic;
using System.Configuration;
using ExportOutlookCalendarToExcel.Common;
using System;

namespace ExportOutlookCalendarToExcel.Tests
{
    [TestClass]
    public class _ManualTests
    {
        [TestMethod]
        public void ReaderCalendarFromPersonalAndConvert()
        {
            var readFromPath = @"C:\Users\Maxim.Alonov\Documents\outlook-export\outlook-export.ics";

            using (var textReader = new StreamReader(readFromPath))
            {
                var calendarICSReader = new CalendarICSReader();
                var activities = calendarICSReader.ReadActivities(textReader);

                ; // breakpoint here
            }
        }

        
    }
}