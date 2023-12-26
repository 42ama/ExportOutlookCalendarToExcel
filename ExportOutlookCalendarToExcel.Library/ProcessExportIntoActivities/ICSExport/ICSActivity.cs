using Ical.Net;
using Ical.Net.CalendarComponents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport
{
    public class ICSActivity : AbstractActivity
    {
        private CalendarEvent _calendarEvent;

        public ICSActivity(CalendarEvent calendarEvent)
        {
            _calendarEvent = calendarEvent;
            
        }

        public override Activity AsActivity()
        {
            var activity = new Activity(_calendarEvent.Summary, _calendarEvent.Start.Value, _calendarEvent.End.Value, _calendarEvent.Attendees.Count > 0);
            return activity;
        }
    }
}
