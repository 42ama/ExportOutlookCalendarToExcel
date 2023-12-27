using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents;

namespace ExportOutlookCalendarToExcel.Tests
{
    [TestClass]
    public class UnpackICSTests
    {
        /// <summary>
        /// Factor by which you need multiply <c>eventsDepth</c> in method <c>AddPlainEvents</c> to get count of events.
        /// </summary>
        private const int ADD_PLAIN_EVENTS_FACTOR = 4;

        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(10)]
        public void PlainEvents_NoChangesAfterUnpack(int eventsDepth)
        {
            var calendar = new Calendar();
            AddPlainEvents(calendar, eventsDepth);

            var actualEvents = calendar.Events.UnpackRecurringEvents();

            for (int i = 0; i < calendar.Events.Count; i++)
            {
                Assert.AreEqual(calendar.Events[i], actualEvents[i]);
            }
        }

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


            var actualEvents = calendar.Events.UnpackRecurringEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }

        [TestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(10)]
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
                }
            });
            var recurringInstancesCount = eventsDepth;

            // Adding plain events.
            AddPlainEvents(calendar, eventsDepth);
            var plainEventsCount = eventsDepth * ADD_PLAIN_EVENTS_FACTOR;


            var actualEvents = calendar.Events.UnpackRecurringEvents();


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
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
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
                        Until = DateTime.Now.AddDays(endRecurrenceAfterDays)
                    }
                }
            });
            var recurringInstancesCount = endRecurrenceAfterDays * 2;


            var actualEvents = calendar.Events.UnpackRecurringEvents();


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


            var actualEvents = calendar.Events.UnpackRecurringEvents();


            Assert.AreEqual(recurringInstancesCount, actualEvents.Count);
        }

        /// <summary>
        /// Adds <paramref name="eventsDepth"/>*4 events to the calendar. 
        /// </summary>
        /// <param name="calendar">Calendar in which events is added.</param>
        /// <param name="eventsDepth">This times 4 events is added to the calendar. Quoter of events is added as prior to today days. Quoter of events is added as next for today events. Quoter of events as added as prior to current hour hours. Quoter of events as added as  next current hour hours. </param>
        private void AddPlainEvents(Calendar calendar, int eventsDepth)
        {
            for (int i = 0; i < eventsDepth; i++)
            {
                calendar.Events.Add(new CalendarEvent
                {
                    Start = new CalDateTime(DateTime.Now),
                    End = new CalDateTime(DateTime.Now.AddDays(i))
                });
                calendar.Events.Add(new CalendarEvent
                {
                    Start = new CalDateTime(DateTime.Now.AddHours(i)),
                    End = new CalDateTime(DateTime.Now.AddDays(i))
                });
                calendar.Events.Add(new CalendarEvent
                {
                    Start = new CalDateTime(DateTime.Now),
                    End = new CalDateTime(DateTime.Now.AddDays(-i))
                });
                calendar.Events.Add(new CalendarEvent
                {
                    Start = new CalDateTime(DateTime.Now.AddHours(-i)),
                    End = new CalDateTime(DateTime.Now.AddDays(-i))
                });
            }

        }
    }
}
