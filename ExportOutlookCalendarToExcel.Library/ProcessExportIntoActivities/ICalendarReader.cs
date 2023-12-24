using ExportOutlookCalendarToExcel.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Logic.CalendarReader
{
    internal interface ICalendarReader
    {
        ActivitiesDateCollection ReadActivities(TextReader reader);
    }
}
