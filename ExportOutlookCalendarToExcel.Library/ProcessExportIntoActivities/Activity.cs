using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Activity.
    /// </summary>
    public class Activity : AbstractActivity
    {
        /// <summary>
        /// Group.
        /// </summary>
        public string Group { get; set; } = string.Empty;

        /// <summary>
        /// Subject.
        /// </summary>
        public string Subject { get; private set; } = string.Empty;

        /// <summary>
        /// Start date.
        /// </summary>
        public DateTime Date { get; private set; }

        /// <summary>
        /// Duration of activity.
        /// </summary>
        public double Duration { get; private set; }

        /// <summary>
        /// Is current activity represents meeting.
        /// </summary>
        public bool IsMeeting { get; private set; }

        /// <param name="subject">Subject</param>
        /// <param name="from">Start of activity.</param>
        /// <param name="to">End of activity.</param>
        /// <param name="isMeeting">Activity is a meeting?</param>
        public Activity(string subject, DateTime from, DateTime to, bool isMeeting)
        {
            Duration = CalculateDuration(to - from);
            Date = from.Date;
            IsMeeting = isMeeting;
            SetSubjectAndGroup(subject ?? Constants.ActivitySettings.SubjectFallback);
        }

        /// <inheritdoc/>
        public override Activity AsActivity()
        {
            return this;
        }

        /// <summary>
        /// Calculate duration in hours and brackets of 0.25h.
        /// </summary>
        /// <param name="duration">Duration of activity.</param>
        /// <returns>Duration of activity in hours and brackets of 0.25h.</returns>
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
        /// Fill subject and group in activity.
        /// </summary>
        /// <param name="subject">Subject</param>
        private void SetSubjectAndGroup(string subject)
        {
            Argument.NotNullOrEmpty(subject, nameof(subject));

            SetDefaultSubjectAndGroup(subject);

            var searchPattern = "%(.*?)%(.*)";// !!! AppConfigProvider.Get(Constants.AppConfig.KeyNames.GroupSearchPattern);
            var regex = new Regex(searchPattern);
            var regexMatch = regex.Match(subject);

            if (regexMatch.Success && regexMatch.Groups.Count > 0)
            {
                Group = regexMatch.Groups[1].Value;
                AddSubject(regexMatch.Groups[2].Value);
            }
        }

        /// <summary>
        /// Fill subject and group in activity with default value.
        /// </summary>
        /// <param name="subject">Subject</param>
        private void SetDefaultSubjectAndGroup(string subject)
        {
            Argument.NotNullOrEmpty(subject, nameof(subject));

            Group = string.Empty;
            AddSubject(subject);
        }

        /// <summary>
        /// Fill subject in activity.
        /// </summary>
        /// <param name="subject">Subject</param>
        private void AddSubject(string subject)
        {
            Argument.NotNullOrEmpty(subject, nameof(subject));

            Subject = (IsMeeting
                ? ActivityRes.MeetingPrefix
                : string.Empty) + subject.Trim();
        }
    }
}
