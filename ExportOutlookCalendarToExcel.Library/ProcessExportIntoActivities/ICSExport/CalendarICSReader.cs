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

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport
{
    public class CalendarICSReader : ICalendarReader
    {
        public ActivitiesDateCollection ReadActivities(TextReader reader)
        {
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

            var activityCollection = new ActivitiesDateCollection(activities);

            return activityCollection;
        }
    }
}
