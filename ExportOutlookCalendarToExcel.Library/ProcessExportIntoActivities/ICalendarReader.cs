using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    public interface ICalendarReader
    {
        ActivitiesGroupedByDateCollection ReadActivities(TextReader reader);
    }
}
