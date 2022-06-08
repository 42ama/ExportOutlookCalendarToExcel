using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData.Model
{
    public class Activity
    {
        const string MEETING_PREFIX = "Встреча. ";

        public string Project { get; set; }

        public string Subject { get; private set; }

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
            SetSubjectAndProject(calendar);
        }   
        
        private void SetSubjectAndProject(CalendarCSV calendar)
        {
            const int substringLenght = 15;
            var canSeparateSubjectAndProject = false;

            if (calendar.Subject.Length >= substringLenght)
            {
                canSeparateSubjectAndProject = calendar.Subject.Substring(0, substringLenght).Contains('.');
            }

            var project = String.Empty;
            var subject = String.Empty;
            if (canSeparateSubjectAndProject)
            {
                var subjectAndProjectSplit = calendar.Subject.Split('.', 2);
                project = subjectAndProjectSplit[0];
                subject = subjectAndProjectSplit[1];
            }
            else
            {
                subject = calendar.Subject;
            }

            Project = project;
            Subject = (calendar.IsMeeting
                ? MEETING_PREFIX
                : string.Empty) + subject.Trim();
        }
    }
}
