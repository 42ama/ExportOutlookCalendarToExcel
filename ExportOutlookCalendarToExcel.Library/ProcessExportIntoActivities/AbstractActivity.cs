using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    public abstract class AbstractActivity
    {

        /// <summary>
        /// View current activity as <c>Activity</c>.
        /// </summary>
        /// <returns>Activity.</returns>
        public abstract Activity AsActivity();
    }
}
