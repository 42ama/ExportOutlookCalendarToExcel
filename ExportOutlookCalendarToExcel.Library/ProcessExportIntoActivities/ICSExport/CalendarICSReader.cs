using Ical.Net;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ical.Net.CalendarComponents;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents;
using ExportOutlookCalendarToExcel._Common;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport
{
    /// <summary>
    /// Reader activities from iCalendar file.
    /// </summary>
    public class CalendarICSReader : ICalendarReader
    {
        /// <summary>
        /// Read activities from iCalendar file.
        /// </summary>
        /// <param name="reader">Reader in which iCalendar file is opened.</param>
        /// <returns>Read activities</returns>
        /// <exception cref="InvalidDataException"></exception>
        public ActivitiesGroupedByDateCollection ReadActivities(TextReader reader)
        {
            Argument.NotNull(reader, nameof(reader));

            var calendar = Calendar.Load(reader);
            if (calendar.Events.Count == 0)
            {
                throw new InvalidDataException("Calendar should have events!");
            }

            var activities = new List<Activity>();
            var unpackedEvents = calendar.Events.UnpackEvents();

            foreach (var unpackedEvent in unpackedEvents)
            {
                var icsActivity = new ICSActivity(unpackedEvent);
                var activity = icsActivity.AsActivity();
                activities.Add(activity);
            }

            var activityCollection = new ActivitiesGroupedByDateCollection(activities);

            return activityCollection;
        }
    }
}
