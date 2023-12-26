using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    internal interface ICalendarReader
    {
        ActivitiesDateCollection ReadActivities(TextReader reader);
    }
}
