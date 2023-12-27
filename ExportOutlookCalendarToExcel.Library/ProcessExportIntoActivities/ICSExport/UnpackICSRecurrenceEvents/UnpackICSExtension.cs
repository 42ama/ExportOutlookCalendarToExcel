using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Proxies;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents
{
    public static class UnpackICSExtension
    {
        /// <summary>
        /// Unpacks all events from <c>Calendar.Events</c> into plain list. In process recurring events information is replaced with specific instances.
        /// </summary>
        /// <param name="calendarEvents">Calendar.Events</param>
        /// <returns>Plain list of events, with recurring events replaced with specific instances.</returns>
        public static List<CalendarEvent> UnpackRecurringEvents(this IUniqueComponentList<CalendarEvent> calendarEvents)
        {
            var result = new List<CalendarEvent>();

            // Find recurring events and instances of events. Key in dictionaries - event uid.
            var recurringEventParents = new List<CalendarEvent>();
            var recurringEventInstances = new Dictionary<string, IList<CalendarEvent>>();

            foreach (var calendarEvent in calendarEvents)
            {
                // Recurring event parent has set RRULE property.
                var isRecurringEventParent = calendarEvent.RecurrenceRules != null;
                if (isRecurringEventParent)
                {
                    recurringEventParents.Add(calendarEvent);
                    continue;
                }

                // Recurring event instance (rescheduled event) has set RECURRENCE-ID property.
                var isRecurringEventInstance = calendarEvent.RecurrenceId != null;
                if (isRecurringEventInstance)
                {
                    if (recurringEventInstances.TryGetValue(calendarEvent.Uid, out var foundInstances))
                    {
                        foundInstances.Add(calendarEvent);
                    }
                    else
                    {
                        recurringEventInstances[calendarEvent.Uid] = new List<CalendarEvent> { calendarEvent };
                    }
                    continue;
                }

                // If event is not recurring parent or instance consider it plain event. Such events we can just add to the result.
                result.Add(calendarEvent);
            }

            foreach (var recurringEventParent in recurringEventParents)
            {
                // Generate series of events. Key is candidate for RECURRENCE-ID.
                var generatedSeries = new Dictionary<IDateTime, CalendarEvent>();
                // !!!

                // While generating events - enrich them.
                // !!!

                // Replace events by found instances.
                var parentUid = recurringEventParent.Uid;
                if (recurringEventInstances.ContainsKey(parentUid))
                {
                    foreach (var recurringEventInstace in recurringEventInstances[parentUid])
                    {
                        // Enrich event.
                        // !!!

                        // Store in same location.
                        generatedSeries[recurringEventInstace.RecurrenceId] = recurringEventInstace;
                    }
                }
            }


            return result;
        }
    }
}
