using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Which subjects should be ignored while creating activity.
    /// </summary>
    internal class ActivitySubjectsToIgnoreProvider
    {
        /// <summary>
        /// Get subjects which should be ignored while creating activity.
        /// </summary>
        /// <returns>Which subjects should be ignored while creating activity.</returns>
        public string[] GetSubjects()
        {
            return new string[]
            {
                ActivitySubjectsToIgnoreProviderRes.Lunch,
                ActivitySubjectsToIgnoreProviderRes.Block,
                ActivitySubjectsToIgnoreProviderRes.Break,
            };
        }
    }
}
