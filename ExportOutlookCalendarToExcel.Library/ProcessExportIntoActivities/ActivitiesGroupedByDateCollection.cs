using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Filtered and ordered activities, which grouped by date.
    /// </summary>
    public class ActivitiesGroupedByDateCollection : IEnumerable<ActivitiesGroupedByDate>
    {
        /// <summary>
        /// Inner main collection of activities with dates.
        /// </summary>
        private readonly IList<ActivitiesGroupedByDate> _activities;

        internal ActivitiesGroupedByDateCollection(IEnumerable<Activity> activities)
        {
            Argument.NotNull(activities, nameof(activities));

            var filteredActivities = FilterActivtities(activities);
            var sortedActivities = SortActivities(filteredActivities);
            _activities = GroupActivitiesByDate(sortedActivities);
        }

        /// <summary>
        /// Get dates by which activities is grouped in sorted order.
        /// </summary>
        /// <returns>Dates by which activities is grouped in sorted order.</returns>
        internal IOrderedEnumerable<DateTime> GetSortedDates()
        {
            return _activities.Select(i => i.Date).OrderBy(i => i.Date);
        }

        /// <summary>
        /// Filter activities - exclude activities with duration <= 0 and subject contained in ignored subjects.
        /// </summary>
        /// <param name="activities">Activities.</param>
        /// <returns>Filtered activities.</returns>
        private IEnumerable<Activity> FilterActivtities(IEnumerable<Activity> activities)
        {
            var subjectsToIgnoreProvider = new ActivitySubjectsToIgnoreProvider();
            var subjectsToIgnore = subjectsToIgnoreProvider.GetSubjects();

            return activities
                .Where(activity => !subjectsToIgnore.Contains(activity.Subject))
                .Where(activity => activity.Duration > 0);
        }

        /// <summary>
        /// Sort activities by group and subject.
        /// </summary>
        /// <param name="activities">Activities.</param>
        /// <returns>Sorted activities.</returns>
        private IOrderedEnumerable<Activity> SortActivities(IEnumerable<Activity> activities)
        {
            return activities
                .OrderByDescending(activity => activity.Group)
                .ThenByDescending(activity => activity.Subject);
        }


        /// <summary>
        /// Group activities by date.
        /// </summary>
        /// <param name="activities">Activities.</param>
        /// <returns>Activities group by date.</returns>
        private List<ActivitiesGroupedByDate> GroupActivitiesByDate(IEnumerable<Activity> activities)
        {
            return activities
                .GroupBy(activity => activity.Date)
                .Select(grouping => new ActivitiesGroupedByDate
                {
                    Date = grouping.Key,
                    Activities = grouping
                })
                .OrderBy(activityGroup => activityGroup.Date)
                .ToList();
        }

        /// <inheritdoc/>
        public IEnumerator<ActivitiesGroupedByDate> GetEnumerator()
        {
            return _activities.GetEnumerator();
        }

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
