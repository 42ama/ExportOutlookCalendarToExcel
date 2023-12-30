using ExportOutlookCalendarToExcel._Common;
using Ical.Net;
using Ical.Net.CalendarComponents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport
{
    /// <summary>
    /// Activity read from iCalendar file.
    /// </summary>
    public class ICSActivity : AbstractActivity
    {
        /// <summary>
        /// Underlaying event read from iCalendar file.
        /// </summary>
        private CalendarEvent _calendarEvent;

        /// <param name="calendarEvent">Read from iCalendar file event.</param>
        public ICSActivity(CalendarEvent calendarEvent)
        {
            Argument.NotNull(calendarEvent, nameof(calendarEvent));

            _calendarEvent = calendarEvent;            
        }

        /// <summary>
        /// View current activity as <c>Activity</c>.
        /// </summary>
        /// <returns>Activity.</returns>
        public override Activity AsActivity()
        {
            var activity = new Activity(_calendarEvent.Summary, _calendarEvent.Start.Value, _calendarEvent.End.Value, _calendarEvent.Attendees.Count > 0);
            return activity;
        }
    }
}
