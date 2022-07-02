﻿using CalendarDataToAxData.Extension;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CalendarDataToAxData.Common;

namespace CalendarDataToAxData.Model
{
    /// <summary>
    /// Коллекция активностей и дат.
    /// </summary>
    public class ActivitiesDateCollection : IEnumerable<ActivitiesWithDate>
    {
        /// <summary>
        /// Коллекция активностей и дат
        /// </summary>
        private IEnumerable<ActivitiesWithDate> _activities;

        /// <summary>
        /// Конструктор.
        /// </summary>
        /// <param name="activities">Коллекция активностей.</param>
        public ActivitiesDateCollection(IEnumerable<Activity> activities)
        {
            Argument.NotNull(activities, nameof(activities));

            var filteredActivieis = FilterActivtities(activities);
            _activities = SelectActivitiesDateCollection(filteredActivieis);
        }

        public IOrderedEnumerable<DateTime> GetSortedDates()
        {
            return _activities.Select(i => i.Date).OrderBy(i => i.Date);
        }

        /// <summary>
        /// Отфильтровать Активности.
        /// </summary>
        /// <param name="activities">Активности.</param>
        /// <returns>Отфильтрованная коллекция Активностей.</returns>
        private IEnumerable<Activity> FilterActivtities(IEnumerable<Activity> activities)
        {
            var subjectsToIgnore = AppConfigProvider.GetStringArray(Constants.AppConfig.KeyNames.SubjectToIgnore);
            
            return activities
                .Where(activity => !subjectsToIgnore.Contains(activity.Subject))
                .Where(activity => activity.Duration > 0)
                .OrderByDescending(activity => activity.Project);
        }

        /// <summary>
        /// Выбрать коллекцию Активностей и дат.
        /// </summary>
        /// <param name="activities">Активности.</param>
        /// <returns>Коллекция Активностей и дат.</returns>
        private IEnumerable<ActivitiesWithDate> SelectActivitiesDateCollection(IEnumerable<Activity> activities)
        {
            return activities
                .GroupBy(activity => activity.Date)
                .Select(grouping => new ActivitiesWithDate
                {
                    Date = grouping.Key,
                    Activities = grouping
                });
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
