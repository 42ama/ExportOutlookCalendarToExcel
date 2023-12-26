using System;

namespace ExportOutlookCalendarToExcel.Library._Common
{

    /// <summary>
    /// Find boundaries of a week.
    /// </summary>
    public static class DateTimeWeekBoundariesExtension
    {
        /// <summary>
        /// Count of days in a week.
        /// </summary>
        private const int DAYS_IN_WEEK = 7;

        /// <summary>
        /// Finds date of the start of the week for given <paramref name="dateTime"/>dateTime</paramref> with assumption that week starts at <param name="weekStart">weekStart</param>.
        /// </summary>
        /// <param name="dateTime">Date for which we are interested in the start of the week.</param>
        /// <param name="weekStart">Day at which week starts.</param>
        /// <returns>Date of the start of the week</returns>
        public static DateTime StartOfWeek(this DateTime dateTime, DayOfWeek weekStart)
        {
            var daysFromStartOfWeekToDate = dateTime.DayOfWeek - weekStart;
            var normalizedDifferenceBetweenStartOfTheWeekAndDate = (DAYS_IN_WEEK + daysFromStartOfWeekToDate) % DAYS_IN_WEEK;
            var substractFromDate = normalizedDifferenceBetweenStartOfTheWeekAndDate * -1;
            return dateTime.AddDays(substractFromDate).Date;
        }

        /// <summary>
        /// Finds date of the end of the week for given <paramref name="dateTime"/>dateTime</paramref> with assumption that week starts at <param name="weekStart">weekStart</param>.
        /// </summary>
        /// <param name="dateTime">Date for which we are interested in the end of the week.</param>
        /// <param name="weekStart">Day at which week starts.</param>
        /// <returns>Date of the end of the week</returns>
        public static DateTime EndOfWeek(this DateTime dateTime, DayOfWeek weekStart)
        {
            var daysFromStartOfWeekToDate = dateTime.DayOfWeek - weekStart;
            var normalizedDifferenceBetweenStartOfTheWeekAndDate = (DAYS_IN_WEEK + daysFromStartOfWeekToDate) % DAYS_IN_WEEK;
            var daysToEndOfWeek = DAYS_IN_WEEK - 1 - normalizedDifferenceBetweenStartOfTheWeekAndDate;

            return dateTime.AddDays(daysToEndOfWeek).Date;
        }
    }
}