using Ical.Net;
using Ical.Net.CalendarComponents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Model
{
    public class ICSActivityFromRecurrence : AbstractActivity
    {
        private CalendarEvent _parentEvent;
        private CalendarEvent _instanceEvent;

        public ICSActivityFromRecurrence(CalendarEvent parentEvent, CalendarEvent instanceEvent)
        {
            _parentEvent = parentEvent;
            _instanceEvent = instanceEvent;
        }

        public override Activity AsActivity()
        {
            var activity = new Activity(_parentEvent.Summary,
                                        _instanceEvent.Start.Value,
                                        _instanceEvent.End.Value,
                                        _parentEvent.Attendees.Count > 1);
            return activity;
        }
    }
}
