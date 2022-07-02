using CalendarDataToAxData.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData.Model
{
    /// <summary>
    /// Активность
    /// </summary>
    public class Activity
    {
        /// <summary>
        /// Проект.
        /// </summary>
        public string Project { get; set; }

        /// <summary>
        /// Тема.
        /// </summary>
        public string Subject { get; private set; }

        /// <summary>
        /// Дата.
        /// </summary>
        public DateTime Date { get; }

        /// <summary>
        /// Длительность.
        /// </summary>
        private TimeSpan _duration;

        /// <summary>
        /// Длительность.
        /// </summary>
        public double Duration
        {
            get
            {
                return _duration.TotalMinutes / 60f;
            }
        }

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        public Activity(CalendarCSV calendar)
        {
            Argument.NotNull(calendar, nameof(calendar));

            _duration = calendar.EndTime - calendar.StartTime;
            Date = DateTime.Parse(calendar.StartDate);
            SetSubjectAndProject(calendar);
        }   
        
        /// <summary>
        /// Заполнить Тему и Проект.
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        private void SetSubjectAndProject(CalendarCSV calendar)
        {
            Argument.NotNull(calendar, nameof(calendar));

            var canSeparateSubjectAndProject = false;

            if (calendar.Subject.Length >= Constants.ActivitySettings.SubstringSeparatorSearchLenght)
            {
                canSeparateSubjectAndProject = calendar.Subject
                                    .Substring(0, Constants.ActivitySettings.SubstringSeparatorSearchLenght)
                                    .Contains(Constants.ActivitySettings.Separator);
            }

            var project = String.Empty;
            var subject = String.Empty;
            if (canSeparateSubjectAndProject)
            {
                var subjectAndProjectSplit = calendar.Subject.Split(Constants.ActivitySettings.Separator, 2);
                project = subjectAndProjectSplit[0];
                subject = subjectAndProjectSplit[1];
            }
            else
            {
                subject = calendar.Subject;
            }

            Project = project;
            Subject = (calendar.IsMeeting
                ? Constants.ActivitySettings.MeetingPrefix
                : string.Empty) + subject.Trim();
        }
    }
}
