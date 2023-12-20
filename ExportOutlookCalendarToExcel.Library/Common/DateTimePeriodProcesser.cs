using ExportOutlookCalendarToExcel.Model;
using Ical.Net.CalendarComponents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Common
{
    internal class DateTimePeriodProcesser
    {
        public Dictionary<DayOfWeek, List<DateTime>> CountDays(DateTime start, DateTime end)
        {
            var allEventsTimespan = end - start;
            var daysInPeriod = allEventsTimespan.Days;

            return Enumerable.Range(0, daysInPeriod + 1)
                            .Select(x => start.AddDays(x))
                            .GroupBy(x => x.DayOfWeek)
                            .ToDictionary(i => i.Key, k => k.ToList());
        }

        // Пример работы
//        var periodProcesser = new DateTimePeriodProcesser();
//        var daysInPeriod = periodProcesser.CountDays(sortedEvents.First().Start.Value, sortedEvents.Last().Start.Value);

//        var recurringEvents = sortedEvents.Where(i => i.RecurrenceRules.Any());
//            foreach (var calendarEvent in recurringEvents)
//            {
//                foreach (var recurrenceRule in calendarEvent.RecurrenceRules)
//                {
//                    if (recurrenceRule.Frequency != FrequencyType.Weekly || recurrenceRule.ByDay.Count == 0)
//                    {
//                        continue;
//                    }

//                    foreach (var weekDay in recurrenceRule.ByDay)
//                    {
//                        if (daysInPeriod.TryGetValue(weekDay.DayOfWeek, out var datesOfReccurance))
//                        {
//                            foreach (var date in datesOfReccurance)
//                            {

//                                var icsActivity = new ICSActivityFromRecurrence(calendarEvent, date, date + calendarEvent.Duration);
//        var activity = icsActivity.AsActivity();
//        activities.Add(activity);
//                            }
//}                        
//                    }
//                }
//            }
    }
}
