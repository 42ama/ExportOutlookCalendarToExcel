using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents
{
    internal static class CalendarEventDeepCopyExtension
    {
        /// <summary>
        /// Creates representing copy of a event. With minimal set of fields.
        /// </summary>
        /// <param name="calendarEvent">Copy to</param>
        /// <param name="other">Copy from</param>
        public static void CustomCopyFrom(this CalendarEvent calendarEvent, CalendarEvent other)
        {
            // Deep copy isn't available for moment, using this hack. (https://github.com/rianjs/ical.net/issues/149)

            calendarEvent.Uid = other.Uid;
            calendarEvent.Created = other.Created;
            calendarEvent.LastModified = other.LastModified;
            calendarEvent.Priority = other.Priority;
            calendarEvent.Sequence = other.Sequence;
            calendarEvent.Summary = other.Summary;
            calendarEvent.Description = other.Description;
            calendarEvent.Start = other.Start;
            calendarEvent.End = other.End;
            calendarEvent.DtStamp = other.DtStamp;

            var attendees = new List<Attendee>();
            foreach (var attendee in calendarEvent.Attendees)
            {
                var attendeeNew = new Attendee();
                attendeeNew.CommonName = attendee.CommonName;

                attendees.Add(attendeeNew);
            }
            calendarEvent.Attendees = attendees;
        }
    }
}
