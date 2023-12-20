using System;

public static class DateTimeExtensions
{
    private const int DAYS_IN_WEEK = 7;
    public static DateTime StartOfWeek(this DateTime dateTime, DayOfWeek startOfWeek)
    {
        var daysFromStartOfWeekToDate = dateTime.DayOfWeek - startOfWeek;
        var normalizedDifferenceBetweenStartOfTheWeekAndDate = (DAYS_IN_WEEK + daysFromStartOfWeekToDate) % DAYS_IN_WEEK;
        var substractFromDate = normalizedDifferenceBetweenStartOfTheWeekAndDate * -1;
        return dateTime.AddDays(substractFromDate).Date;
    }

    public static DateTime EndOfWeek(this DateTime dateTime, DayOfWeek startOfWeek)
    {
        var daysFromStartOfWeekToDate = dateTime.DayOfWeek - startOfWeek;
        var normalizedDifferenceBetweenStartOfTheWeekAndDate = (DAYS_IN_WEEK + daysFromStartOfWeekToDate) % DAYS_IN_WEEK;
        var daysToEndOfWeek = DAYS_IN_WEEK - 1 - normalizedDifferenceBetweenStartOfTheWeekAndDate;

        return dateTime.AddDays(daysToEndOfWeek).Date;
    }
}