using Ical.Net;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ical.Net.CalendarComponents;

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

            // Events without recurrences.
            var plainEvents = calendar.Events.Where(i => !i.RecurrenceRules.Any() && i.RecurrenceId == null);
            foreach (var calendarEvent in plainEvents)
            {
                var icsActivity = new ICSActivity(calendarEvent); // !!! Можем через DI вытаскиывать, но из-за того что много экземпляров создаём в моменте может медленно быть. 
                var activity = icsActivity.AsActivity();
                activities.Add(activity);
            }

            var recurranceParentEvents = calendar.Events.Where(i => i.RecurrenceRules.Any());
            var recurranceInstanceEventsByUid = calendar.Events.Where(i => i.RecurrenceId != null)
                                                          .GroupBy(i => i.Uid)
                                                          .ToDictionary(k => k.Key, v => v.ToList());

            foreach (var recurranceParent in recurranceParentEvents)
            {
                // In this branch system processes instances of recurring events.
                if (recurranceInstanceEventsByUid.TryGetValue(recurranceParent.Uid, out var recurranceInstanceEvents))
                {
                    
                    foreach (var recurranceInstance in recurranceInstanceEvents)
                    {
                        var icsActivity = new ICSActivityFromRecurrence(recurranceParent, recurranceInstance);
                        var activity = icsActivity.AsActivity();
                        activities.Add(activity);
                    }
                }
                else // In this branch system processes parents records of recurring events.
                {
                    var icsActivity = new ICSActivity(recurranceParent);
                    var activity = icsActivity.AsActivity();
                    activities.Add(activity);
                }
            }

            var activityCollection = new ActivitiesDateCollection(activities);

            return activityCollection;
        }
    }
}
