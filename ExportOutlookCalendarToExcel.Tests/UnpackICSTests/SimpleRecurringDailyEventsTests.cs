using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents;
using ExportOutlookCalendarToExcel.Tests.UnpackICSHelpers;

namespace ExportOutlookCalendarToExcel.Tests.UnpackICSTests
{
    /// <summary>
    /// Recurring events with single RRULE.
    /// </summary>
    [TestClass]
    public class SimpleRecurringDailyEventsTests
    {
        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(10)]
        public void SingleRecurringEvent_DailyRecurrence_UnpackedCorrectly(int endRecurrenceAfterDays)
        {
            var calendar = new Calendar();

            // Adding recurring event.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
                    }
                }
            });
            var recurringInstancesCount = endRecurrenceAfterDays;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }

        [TestMethod]
        [DataRow(3)]
        [DataRow(4)]
        [DataRow(5)]
        public void SingleRecurringEvent_DailyRecurrence_ExceptionDates_UnpackedCorrectly(int endRecurrenceAfterDays)
        {
            var calendar = new Calendar();

            // Adding recurring event.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
                    }
                },
                ExceptionDates = new List<PeriodList>
                {
                    new PeriodList
                    {
                        new Period(
                            new CalDateTime(DateTime.Now.AddDays(1).AddMinutes(-30)),
                            new CalDateTime(DateTime.Now.AddDays(1).AddMinutes(60))
                            )
                    },
                    new PeriodList
                    {
                        new Period(
                            new CalDateTime(DateTime.Now.AddDays(2).AddMinutes(-30)),
                            new CalDateTime(DateTime.Now.AddDays(2).AddMinutes(60))
                            )
                    },
                }
            });
            // Decrease by number of exception days.
            var recurringInstancesCount = endRecurrenceAfterDays - 2;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }

        [TestMethod]
        [DataRow(3)]
        [DataRow(4)]
        [DataRow(5)]
        public void SingleRecurringEvent_DailyRecurrence_PlainEvents_UnpackedCorrectly(int eventsDepth)
        {
            var calendar = new Calendar();

            // Adding recurring event.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddDays(eventsDepth)
                    }
                },
                ExceptionDates = new List<PeriodList>
                {
                    new PeriodList
                    {
                        new Period(
                            new CalDateTime(DateTime.Now.AddDays(1).AddMinutes(-30)),
                            new CalDateTime(DateTime.Now.AddDays(1).AddMinutes(60))
                            )
                    },
                    new PeriodList
                    {
                        new Period(
                            new CalDateTime(DateTime.Now.AddDays(2).AddMinutes(-30)),
                            new CalDateTime(DateTime.Now.AddDays(2).AddMinutes(60))
                            )
                    },
                }
            });
            var recurringInstancesCount = eventsDepth;

            // Adding plain events.
            PlainEventsTestHelper.AddPlainEvents(calendar, eventsDepth);
            var plainEventsCount = eventsDepth * PlainEventsTestHelper.ADD_PLAIN_EVENTS_FACTOR - 2;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount + plainEventsCount, actualEvents.Count);
        }

        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(10)]
        public void SingleRecurringEvent_DailyRecurrence_PlainEvents_ExceptionDates_UnpackedCorrectly(int eventsDepth)
        {
            var calendar = new Calendar();

            // Adding recurring event.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddDays(eventsDepth)
                    }
                }
            });
            var recurringInstancesCount = eventsDepth;

            // Adding plain events.
            PlainEventsTestHelper.AddPlainEvents(calendar, eventsDepth);
            var plainEventsCount = eventsDepth * PlainEventsTestHelper.ADD_PLAIN_EVENTS_FACTOR;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount + plainEventsCount, actualEvents.Count);
        }


        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(10)]
        public void DoubleRecurringEvents_DailyRecurrence_UnpackedCorrectly(int endRecurrenceAfterDays)
        {
            var calendar = new Calendar();

            // Adding recurring events.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddHours(1).AddDays(endRecurrenceAfterDays)
                    }
                }
            });
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now.AddMinutes(30)),
                End = new CalDateTime(DateTime.Now.AddMinutes(60)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddHours(1).AddDays(endRecurrenceAfterDays)
                    }
                }
            });
            var recurringInstancesCount = endRecurrenceAfterDays * 2;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }


        [TestMethod]
        [DataRow(2)]
        [DataRow(3)]
        [DataRow(10)]
        public void SingleRecurringEvent_DailyRecurrence_TwoMoved_UnpackedCorrectly(int endRecurrenceAfterDays)
        {
            var calendar = new Calendar();

            // Adding recurring event.
            var recurringEventUid = Guid.NewGuid().ToString();
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now),
                End = new CalDateTime(DateTime.Now.AddMinutes(30)),
                RecurrenceRules = new List<RecurrencePattern> {
                    new RecurrencePattern(FrequencyType.Daily, 1)
                    {
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
                    }
                },
                Uid = recurringEventUid
            });

            // Adding moved events.
            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now.AddMinutes(120)),
                End = new CalDateTime(DateTime.Now.AddMinutes(150)),
                Uid = recurringEventUid,
                RecurrenceId = new CalDateTime(DateTime.Now)
            });

            calendar.AddChild(new CalendarEvent
            {
                Start = new CalDateTime(DateTime.Now.AddMinutes(120)),
                End = new CalDateTime(DateTime.Now.AddMinutes(150)),
                Uid = recurringEventUid,
                RecurrenceId = new CalDateTime(DateTime.Now.AddDays(1))
            });

            var recurringInstancesCount = endRecurrenceAfterDays;


            var actualEvents = calendar.Events.UnpackEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }
    }
}
