using ExportOutlookCalendarToExcel.Common;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExportOutlookCalendarToExcel.Model
{
    /// <summary>
    /// Активность
    /// </summary>
    public class Activity : AbstractActivity
    {
        /// <summary>
        /// Проект.
        /// </summary>
        public string Project { get; set; } = string.Empty;

        /// <summary>
        /// Тема.
        /// </summary>
        public string Subject { get; private set; } = string.Empty;

        /// <summary>
        /// Дата.
        /// </summary>
        public DateTime Date { get; private set; }


        /// <summary>
        /// Дата напоминания
        /// </summary>
        public DateTime DateOfNotification { get; }

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

        public Activity(string subject, DateTime from, DateTime to, bool isMeeting)
        {
            _duration =  to - from;
            Date = from.Date;
            SetSubjectAndProject(isMeeting, subject ?? Constants.ActivitySettings.SubjectFallback);
        }

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        public Activity(CalendarCSV calendar)
        {
            Argument.NotNull(calendar, nameof(calendar));

            _duration = calendar.EndTime - calendar.StartTime;
            SetDate(calendar);
            SetSubjectAndProject(calendar.IsMeeting, calendar.Subject ?? Constants.ActivitySettings.SubjectFallback);
        }

        public override Activity AsActivity()
        {
            return this;
        }

        /// <summary>
        /// Установить дату.
        /// </summary>
        /// <param name="calendar">Календарь</param>
        private void SetDate(CalendarCSV calendar)
        {
            // Заметил, что для повторяющихся событий calendar.StartDate иногда ставится датой начала серии событий, что ведёт к ошибкам при выгрузке. А вот calendar.NotificationDate хранит корректную дату. Поэтому если эти даты не равны, то лучше использовать calendar.NotificationDate

            var isStardDateIncorrect = calendar.NotificationDate != default
                                        && calendar.StartDate != calendar.NotificationDate;
            var dateStringToSet = isStardDateIncorrect
                ? calendar.NotificationDate
                : calendar.StartDate;

            Date = DateTime.Parse(dateStringToSet);
        }

        /// <summary>
        /// Заполнить Тему и Проект.
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        private void SetSubjectAndProject(bool isMeeting, string subject)
        {
            SetDefaultSubjectAndProject(isMeeting, subject);

            var searchPattern = AppConfigProvider.Get(Constants.AppConfig.KeyNames.ProjectSearchPattern);
            var regex = new Regex(searchPattern);
            var regexMatch = regex.Match(subject);

            if (regexMatch.Success && regexMatch.Groups.Count > 0)
            {
                Project = regexMatch.Groups[1].Value;
                AddSubject(isMeeting, regexMatch.Groups[2].Value);
            }
        }

        /// <summary>
        /// Заполнить Тему и Проект стандартными значениями
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        private void SetDefaultSubjectAndProject(bool isMeeting, string subject)
        {
            Project = "";
            AddSubject(isMeeting, subject);
        }

        /// <summary>
        /// Заполнить Тему и Проект стандартными значениями
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        private void AddSubject(bool isMeeting, string subject)
        {
            Argument.NotNullOrEmpty(subject, nameof(subject));

            Subject = (isMeeting
                ? Constants.ActivitySettings.MeetingPrefix
                : string.Empty) + subject.Trim();
        }
    }
}
