using ExportOutlookCalendarToExcel.Common;
using ExportOutlookCalendarToExcel.Library.Logic.FilepathLocationStrategy;
using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities;
using ExportOutlookCalendarToExcel.Logic.CalendarReader;
using ExportOutlookCalendarToExcel.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Logic.ResultBuilder
{
    public class ICSActivitiesProcessor : AbstractActivitiesProcessor
    {
        public ICSActivitiesProcessor(string filePath) : base(filePath) { }

        public override ActivitiesDateCollection ReadActivities()
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
