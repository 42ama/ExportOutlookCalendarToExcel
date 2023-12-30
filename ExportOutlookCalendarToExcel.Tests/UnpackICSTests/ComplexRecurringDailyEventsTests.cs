using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents;
using ExportOutlookCalendarToExcel.Tests.UnpackICSHelpers;

namespace ExportOutlookCalendarToExcel.Tests.UnpackICSTests
{
    /// <summary>
    /// Recurring events with two or more RRULE's.
    /// </summary>
    [TestClass]
    public class ComplexRecurringDailyEventsTests
    {
        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(10)]
        [TestCategory("NotImplemented")]
        public void ComplexRecurringEvent_DailyRecurrence_TwoRRULE_UnpackedCorrectly(int endRecurrenceAfterDays)
        {
            var calendar = new Calendar();

            // Adding recurring event.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 2)
                    {
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
                    },
                    new RecurrencePattern(FrequencyType.Daily, 3)
                    {
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
                    }
                }
            });
            var recurringInstancesCount = endRecurrenceAfterDays / 2 + endRecurrenceAfterDays / 3;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }
    }
}
