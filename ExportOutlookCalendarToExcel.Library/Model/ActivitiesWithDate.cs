using System;
using System.Collections.Generic;
using System.Text;

namespace ExportOutlookCalendarToExcel.Model
{
    /// <summary>
    /// Активности и даты.
    /// </summary>
    public class ActivitiesWithDate
    {
        /// <summary>
        /// Активности.
        /// </summary>
        public IEnumerable<Activity> Activities { get; set; }

        /// <summary>
        /// Дата.
        /// </summary>
        public DateTime Date { get; set; }
    }
}
