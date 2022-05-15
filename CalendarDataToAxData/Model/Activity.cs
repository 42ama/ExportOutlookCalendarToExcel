using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData.Model
{
    public class Activity
    {
        const string MEETING_PREFIX = "Встреча. ";

        public string Subject { get; }

        private TimeSpan _duration;
        public string Date { get; }
        public double Duration
        {
            get
            {
                return _duration.TotalMinutes / 60f;
            }
        }
        public string DurationFormated { get
            {
                var minutesInHours = Duration;
                return String.Format("{0:n}", minutesInHours);
            }
        }

        public Activity() { }

        public Activity(CalendarCSV calendar)
        {
            if (calendar is null)
            {
                throw new ArgumentNullException(nameof(calendar));
            }

            _duration = calendar.EndTime - calendar.StartTime;
            Date = calendar.StartDate;
            Subject = (calendar.IsMeeting
                ? MEETING_PREFIX
                : string.Empty) + calendar.Subject;
        }        
    }
}
