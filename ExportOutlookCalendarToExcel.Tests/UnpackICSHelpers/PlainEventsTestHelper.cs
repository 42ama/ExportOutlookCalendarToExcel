using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Tests.UnpackICSHelpers
{
    internal static class PlainEventsTestHelper
    {
        /// <summary>
        /// Factor by which you need multiply <c>eventsDepth</c> in method <c>AddPlainEvents</c> to get count of events.
        /// </summary>
        public const int ADD_PLAIN_EVENTS_FACTOR = 4;



        /// <summary>
        /// Adds <paramref name="eventsDepth"/>*4 events to the calendar. 
        /// </summary>
        /// <param name="calendar">Calendar in which events is added.</param>
        /// <param name="eventsDepth">This times 4 events is added to the calendar. Quoter of events is added as prior to today days. Quoter of events is added as next for today events. Quoter of events as added as prior to current hour hours. Quoter of events as added as  next current hour hours. </param>
        public static void AddPlainEvents(Calendar calendar, int eventsDepth)
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
