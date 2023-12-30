using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    internal class ActivitySubjectsToIgnoreProvider
    {
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
