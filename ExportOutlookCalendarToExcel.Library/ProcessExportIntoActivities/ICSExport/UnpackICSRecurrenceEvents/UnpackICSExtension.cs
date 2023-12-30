using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Proxies;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport.UnpackICSRecurrenceEvents
{
    public static class UnpackICSExtension
    {
        /// <summary>
        /// Unpacks all events from <c>Calendar.Events</c> into plain list. In process recurring events information is replaced with specific instances.
        /// </summary>
        /// <param name="calendarEvents">Calendar.Events</param>
        /// <returns>Plain list of events, with recurring events replaced with specific instances.</returns>
        public static List<CalendarEvent> UnpackEvents(this IUniqueComponentList<CalendarEvent> calendarEvents)
        {
            var result = new List<CalendarEvent>();

            // Find recurring events and instances of events. Key in dictionaries - event uid.
            var parentRecurringEvents = new List<CalendarEvent>();
            var instanceRecurringEvents = new Dictionary<string, IList<CalendarEvent>>();

            foreach (var calendarEvent in calendarEvents)
            {
                // Recurring event parent has set RRULE property.
                var isParent = calendarEvent.RecurrenceRules.Count > 0;
                if (isParent)
                {
                    parentRecurringEvents.Add(calendarEvent);
                    continue;
                }

                // Recurring event instance (rescheduled event) has set RECURRENCE-ID property.
                var isInstance = calendarEvent.RecurrenceId != null;
                if (isInstance)
                {
                    if (instanceRecurringEvents.TryGetValue(calendarEvent.Uid, out var foundInstances))
                    {
                        foundInstances.Add(calendarEvent);
                    }
                    else
                    {
                        instanceRecurringEvents[calendarEvent.Uid] = new List<CalendarEvent> { calendarEvent };
                    }
                    continue;
                }

                // If event is not recurring parent or instance consider it plain event. Such events we can just add to the result.
                result.Add(calendarEvent);
            }

            foreach (var parent in parentRecurringEvents)
            {
                // Generate series of events. Key is candidate for RECURRENCE-ID.
                var generatedSeries = GenerateSeries(parent);

                // Replace events by found instances.
                var parentUid = parent.Uid;
                if (instanceRecurringEvents.ContainsKey(parentUid))
                {
                    foreach (var instance in instanceRecurringEvents[parentUid])
                    {
                        // Enrich event.
                        CopyFromParentToIntance(parent, instance);

                        // Store in same location.
                        generatedSeries[instance.RecurrenceId] = instance;
                    }
                }

                // Events stored in values, keys here just for quick access to event replacing.
                foreach (var generatedEvent in generatedSeries.Values)
                {
                    result.Add(generatedEvent);
                }
            }

            return result;
        }

        /// <summary>
        /// Copies data from parent recurrence event to instance recurrence event.
        /// </summary>
        /// <param name="parent">Parent event from which data is copied.</param>
        /// <param name="instance">Instance event TO which data is copied.</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        private static void CopyFromParentToIntance(CalendarEvent parent, CalendarEvent instance)
        {
            if (parent == null) { throw new ArgumentNullException(nameof(parent)); }
            if (instance == null) { throw new ArgumentNullException(nameof(instance)); }

            if (parent.RecurrenceRules.Count == 0)
            {
                throw new ArgumentException("Parent events should have RRULE property set. Probably what you trying to pass here is not a parent event.", nameof(parent));
            }

            if (instance.RecurrenceRules.Count > 0)
            {
                throw new ArgumentException("Instance events should NOT have RRULE property set. Probably what you trying to pass here is not an instance event.", nameof(instance));
            }

            var instanceStart = instance.Start;
            var instanceEnd = instance.End;
            var instanceRecurranceId = instance.RecurrenceId;
            var instanceCreated = instance.Created;
            var instanceDtStamp = instance.DtStamp;

            instance.CustomCopyFrom(parent);

            instance.RecurrenceRules = default;
            instance.RecurrenceDates = default;
            instance.ExceptionDates = default;

            instance.Start = instanceStart;
            instance.End = instanceEnd;
            instance.RecurrenceId = instanceRecurranceId;
            instance.Created = instanceCreated;
            instance.DtStamp = instanceDtStamp;
        }

        /// <summary>
        /// Creates event from from parent recurrence event.
        /// </summary>
        /// <param name="parent">Parent event from which data are copied.</param>
        /// <param name="start">Start of new copy event.</param>
        /// <returns>Copied event.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        private static CalendarEvent CreateEventBasedOnParent(CalendarEvent parent, IDateTime start)
        {
            if (parent == null) { throw new ArgumentNullException(nameof(parent)); }

            if (parent.RecurrenceRules.Count == 0)
            {
                throw new ArgumentException("Parent events should have RRULE property set. Probably what you trying to pass here is not a parent event.", nameof(parent));
            }

            var copy = new CalendarEvent();
            copy.CustomCopyFrom(parent);

            copy.RecurrenceRules = default;
            copy.RecurrenceDates = default;
            copy.ExceptionDates = default;

            copy.Start = start;

            return copy;
        }

        /// <summary>
        /// Generate series of instances of recurring event.
        /// </summary>
        /// <param name="recurringEvent">Recurring event from which recurrence rules and event information will be taken.</param>
        /// <returns>Events stored in values, keys here just for quick access to event replacing.</returns>
        private static Dictionary<IDateTime, CalendarEvent> GenerateSeries(CalendarEvent recurringEvent)
        {
            if (recurringEvent == null) { throw new ArgumentNullException(nameof(recurringEvent)); }

            var eventData = new EventData
            {
                Start = recurringEvent.Start,
                Duration = recurringEvent.Duration,
                ExceptionDates = recurringEvent.ExceptionDates
            };

            var resultSeries = new Dictionary<IDateTime, CalendarEvent>();
            foreach (var rule in recurringEvent.RecurrenceRules.ToList())
            {
#pragma warning disable S1854 // Unused assignments should be removed
                var ruleEventDateTimeStartSeries = new List<IDateTime>();
#pragma warning restore S1854 // Unused assignments should be removed

                // Generate series without exception dates.
                switch (rule.Frequency)
                {
                    case FrequencyType.Daily:
                        ruleEventDateTimeStartSeries = GenerateSeries_Daily(eventData, rule);
                        break;
                    case FrequencyType.Weekly:
                        ruleEventDateTimeStartSeries = GenerateSeries_Weekly(eventData, rule);
                        break;
                    default:
                        var logger = LogManager.GetCurrentClassLogger();
                        logger.Error($"Generating series of events with {rule.Frequency} is not implemented. Current recurrence rule processing will be abandoned.");
                        break;
                }

                // Enrich events.
                foreach (var startDateTime in ruleEventDateTimeStartSeries)
                {
                    var copiedEvent = CreateEventBasedOnParent(recurringEvent, startDateTime);
                    resultSeries.Add(startDateTime, copiedEvent);
                }
            }

            return resultSeries;
        }

        /// <summary>
        /// Generate list of dates in which recurring events should occur. Works only for events with daily rule.
        /// </summary>
        /// <param name="eventData">Information about parent event.</param>
        /// <param name="recurrencePattern">Recurrence parent which describes series.</param>
        /// <returns>List of dates in which recurring events should occur.</returns>
        private static List<IDateTime> GenerateSeries_Daily(EventData eventData, RecurrencePattern recurrencePattern)
        {
            var eventDateTimeStartSeries = new List<IDateTime>();
            var daysSet = new HashSet<DayOfWeek>();
            foreach (var weekDay in recurrencePattern.ByDay)
            {
                daysSet.Add(weekDay.DayOfWeek);
            }

            var seriesDuration = recurrencePattern.Until - eventData.Start.Value;

            // As increment we are adding interval
            for (int i = 0; i < seriesDuration.Days; i +=  recurrencePattern.Interval)
            {
                var dayOnProbe = eventData.Start.AddDays(i);
                
                // If there are rules for specific days, check does day of a week match the rule.
                if (daysSet.Count > 0 && !daysSet.Contains(dayOnProbe.DayOfWeek))
                {
                    continue;
                }

                var dayAsPeriod = new Period(new CalDateTime(dayOnProbe), eventData.Duration);

                var dayIsException = false;
                foreach (var exceptionDatesPeriodList in eventData.ExceptionDates)
                {
                    foreach (var exceptionPeriod in exceptionDatesPeriodList)
                    {
                        var collidesWithPeriod = exceptionPeriod.CollidesWith(dayAsPeriod);
                        if (collidesWithPeriod)
                        {
                            dayIsException = true;
                            break;
                        }

                    }
                }

                // Check is day will be found in exception dates.
                if (dayIsException)
                {
                    continue;
                }

                eventDateTimeStartSeries.Add(dayOnProbe);
            }

            return eventDateTimeStartSeries;
        }

        /// <summary>
        /// Generate list of dates in which recurring events should occur. Works only for events with weekly rule.
        /// </summary>
        /// <param name="eventData">Information about parent event.</param>
        /// <param name="recurrencePattern">Recurrence parent which describes series.</param>
        /// <returns>List of dates in which recurring events should occur.</returns>
        private static List<IDateTime> GenerateSeries_Weekly(EventData eventData, RecurrencePattern recurrencePattern)
        {
            var eventDateTimeStartSeries = new List<IDateTime>();
            var daysSet = new HashSet<DayOfWeek>();
            foreach (var weekDay in recurrencePattern.ByDay)
            {
                daysSet.Add(weekDay.DayOfWeek);
            }

            var weeksNumbersSet = recurrencePattern.ByWeekNo.ToHashSet();// !!! This implementation does not process negatives.

            var seriesDuration = recurrencePattern.Until - eventData.Start.Value;


            var previousWeekNumber = GetIso8601WeekOfYear(eventData.Start.Value);

            for (int i = 0; i < seriesDuration.Days; i++)
            {
                var dayOnProbe = eventData.Start.AddDays(i);

                var weekNumber = GetIso8601WeekOfYear(dayOnProbe.Value);

                // If we are in different week then the previous week, we might cut it by interval parameter.
                if (weekNumber != previousWeekNumber)
                {
                    // With interval = 1 this will result to 0 for any number. This means we taking every week.
                    // With interval = 2 this will result to 0 for every other number. This means we taking every other week.
                    // And etc.
                    var shouldCreateEventsInThisWeek = (weekNumber - previousWeekNumber) % recurrencePattern.Interval == 0;
                    if (!shouldCreateEventsInThisWeek)
                    {
                        continue;
                    }
                    else
                    {
                        // If we are staying within borders of current week, we can skip calculations.
                        previousWeekNumber = weekNumber;
                    }
                }

                // If there are rules for specific week numbers, check does week match the rule.
                if (weeksNumbersSet.Count > 0 && !weeksNumbersSet.Contains(weekNumber))
                {
                    continue;
                }

                // If there are rules for specific days, check does day of a week match the rule.
                if (daysSet.Count > 0 && !daysSet.Contains(dayOnProbe.DayOfWeek))
                {
                    continue;
                }

                var dayAsPeriod = new Period(new CalDateTime(dayOnProbe), eventData.Duration);
                var dayIsException = eventData.ExceptionDates.Any(excpetionDatePeriod => excpetionDatePeriod.Contains(dayAsPeriod));

                // Check is day will be found in exception dates.
                if (dayIsException)
                {
                    continue;
                }

                eventDateTimeStartSeries.Add(dayOnProbe);
            }

            return eventDateTimeStartSeries;
        }

        // This presumes that weeks start with Monday.
        // Week 1 is the 1st week of the year with a Thursday in it.
        private static int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }
    }

    /// <summary>
    /// Information about event.
    /// </summary>
    internal class EventData
    {
        /// <summary>
        /// DateTime at which event starts.
        /// </summary>
        public IDateTime Start { get; set; }

        /// <summary>
        /// Duration of event.
        /// </summary>
        public TimeSpan Duration { get; set; }

        /// <summary>
        /// Period in which event should not occur.
        /// </summary>
        public IList<PeriodList> ExceptionDates { get; set; }
    }
        
}
