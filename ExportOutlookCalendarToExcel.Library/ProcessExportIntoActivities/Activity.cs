using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
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
        /// Длительность.
        /// </summary>
        public double Duration { get; private set; }

        public Activity(string subject, DateTime from, DateTime to, bool isMeeting)
        {
            Duration = CalculateDuration(to - from);
            Date = from.Date;
            SetSubjectAndProject(isMeeting, subject ?? Constants.ActivitySettings.SubjectFallback);
        }

        public override Activity AsActivity()
        {
            return this;
        }

        private double CalculateDuration(TimeSpan duration)
        {
            const double DIVISOR_IN_MINUTES = 15d;  // 15 minutes is minimal bracket.
            const double HOUR_BRACKET = DIVISOR_IN_MINUTES / 60d; // Bracket as part of hour (0.25 for 15 minutes). 
            const double PRESCISION = 0.001d;

            var durationTimeInBracketsRaw = duration.TotalMinutes / DIVISOR_IN_MINUTES;
            var durationTimeInBrackets = Math.Truncate(durationTimeInBracketsRaw);

            // If there reminder, then we rounding up time.
            var shouldTimeBeInNextBracket = (durationTimeInBracketsRaw - durationTimeInBrackets) > PRESCISION;
            if (shouldTimeBeInNextBracket)
            {
                durationTimeInBrackets++;
            }

            var durationCalculated = durationTimeInBrackets * HOUR_BRACKET;

            return durationCalculated;
        }

        /// <summary>
        /// Заполнить Тему и Проект.
        /// </summary>
        /// <param name="calendar">Календарь.</param>
        private void SetSubjectAndProject(bool isMeeting, string subject)
        {
            SetDefaultSubjectAndProject(isMeeting, subject);

            var searchPattern = "%(.*?)%(.*)";// !!! AppConfigProvider.Get(Constants.AppConfig.KeyNames.ProjectSearchPattern);
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
