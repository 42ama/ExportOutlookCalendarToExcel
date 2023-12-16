using ExportOutlookCalendarToExcel.Model;
using Ical.Net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Logic.CalendarReader
{
    public class CalendarICSReader : ICalendarReader
    {
        public ActivitiesDateCollection ReadActivities(TextReader reader)
        {
            var calendar = Calendar.Load(reader);
            var activities = new List<Activity>();
            foreach (var calendarEvent in calendar.Events)
            {
                var icsActivity = new ICSActivity(calendarEvent); // !!! Можем через DI вытаскиывать, но из-за того что много экземпляров создаём в моменте может медленно быть. 
                var activity = icsActivity.AsActivity();
                activities.Add(activity);
            }            

            return new ActivitiesDateCollection(activities);
        }
    }
}
