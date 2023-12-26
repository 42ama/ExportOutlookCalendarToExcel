using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities.ICSExport
{
    internal class ICSActivitiesProcessor : AbstractActivitiesProcessor
    {
        internal ICSActivitiesProcessor(string filePath) : base(filePath) { }

        internal override ActivitiesDateCollection ReadActivities()
        {
            var readFromPath = GetFilePathToReadFrom();

            using (var textReader = GetTextReader(readFromPath))
            {
                var сalendarICSReader = new CalendarICSReader();
                return сalendarICSReader.ReadActivities(textReader);
            }
        }
    }
}
