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
    /// <summary>
    /// Read iCalendar file and return as <c>ActivitiesGroupedByDateCollection</c>.
    /// </summary>
    internal class ICSActivitiesProcessor : AbstractActivitiesProcessor
    {
        /// <param name="filePath">Path to iCalendar tile</param>
        internal ICSActivitiesProcessor(string filePath) : base(filePath) { }

        /// <summary>
        /// Read activities from iCalendar file.
        /// </summary>
        /// <returns>Activities.</returns>
        internal override ActivitiesGroupedByDateCollection ReadActivities()
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
