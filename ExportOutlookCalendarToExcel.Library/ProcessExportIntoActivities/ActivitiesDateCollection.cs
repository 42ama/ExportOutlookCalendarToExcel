﻿using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Коллекция активностей и дат.
    /// </summary>
    public class ActivitiesDateCollection : IEnumerable<ActivitiesWithDate>
    {
        /// <summary>
        /// Коллекция активностей и дат
        /// </summary>
        private readonly IList<ActivitiesWithDate> _activities;

        internal DateTime First { get; private set; }
        internal DateTime Last { get; private set; }

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="activities">Коллекция активностей.</param>
        internal ActivitiesDateCollection(IEnumerable<Activity> activities)
        {
            Argument.NotNull(activities, nameof(activities));

            var filteredActivieis = FilterActivtities(activities);
            _activities = SelectActivitiesDateCollection(filteredActivieis);

            if (_activities.Count > 0)
            {
                First = _activities.First().Date;
                Last = _activities.Last().Date;
            }
        }

        /// <summary>
        /// Получает отсортированные по дате активности.
        /// </summary>
        /// <returns>Отсортированные по дате активности.</returns>
        internal IOrderedEnumerable<DateTime> GetSortedDates()
        {
            return _activities.Select(i => i.Date).OrderBy(i => i.Date);
        }

        /// <summary>
        /// Заполняет пустую группу найденным в текстовой строке.
        /// </summary>
        internal void TryFillEmptyGroup()
        {
            var allGroups = _activities
                    .SelectMany(activityWithDate => activityWithDate.Activities)
                    .Where(activity => !string.IsNullOrEmpty(activity.Group))
                    .GroupBy(activity => activity.Group)
                    .Select(activityGroup => activityGroup.Key);

            foreach (var activityWithDate in _activities)
            {
                foreach (var activity in activityWithDate.Activities)
                {
                    if (!string.IsNullOrEmpty(activity.Group))
                    {
                        continue;
                    }

                    var newGroup = allGroups.FirstOrDefault(group => activity.Subject.Contains(group)); // !!! Тут в Contains Должен быть StringComparison.OrdinalIgnoreCase
                    activity.Group = newGroup;
                }
            }
        }

        /// <summary>
        /// Отфильтровать Активности.
        /// </summary>
        /// <param name="activities">Активности.</param>
        /// <returns>Отфильтрованная коллекция Активностей.</returns>
        private IEnumerable<Activity> FilterActivtities(IEnumerable<Activity> activities)
        {
            var subjectsToIgnoreProvider = new ActivitySubjectsToIgnoreProvider();
            var subjectsToIgnore = subjectsToIgnoreProvider.GetSubjects();

            return activities
                .Where(activity => !subjectsToIgnore.Contains(activity.Subject))
                .Where(activity => activity.Duration > 0)
                .OrderByDescending(activity => activity.Group)
                .ThenByDescending(activity => activity.Subject);
        }

        /// <summary>
        /// Выбрать коллекцию Активностей и дат.
        /// </summary>
        /// <param name="activities">Активности.</param>
        /// <returns>Коллекция Активностей и дат.</returns>
        private List<ActivitiesWithDate> SelectActivitiesDateCollection(IEnumerable<Activity> activities)
        {
            return activities
                .GroupBy(activity => activity.Date)
                .Select(grouping => new ActivitiesWithDate
                {
                    Date = grouping.Key,
                    Activities = grouping
                })
                .OrderBy(activityGroup => activityGroup.Date)
                .ToList();
        }

        /// <summary>
        /// Получить Enumerator коллекции Активностей и дат.
        /// </summary>
        /// <returns>Enumerator коллекции Активностей и дат.</returns>
        public IEnumerator<ActivitiesWithDate> GetEnumerator()
        {
            return _activities.GetEnumerator();
        }

        /// <summary>
        /// Получить Enumerator текущей коллекции.
        /// </summary>
        /// <returns>Enumerator текущей коллекции.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
