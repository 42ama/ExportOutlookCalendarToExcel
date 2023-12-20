using ExportOutlookCalendarToExcel.Common;
using ExportOutlookCalendarToExcel.Library.Logic.FilepathLocationStrategy;
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
    public class ICSResultBuilder : AbstractResultBuilder
    {
        public ICSResultBuilder(string filePath) : base(filePath) { }

        protected override void Init()
        {
            base.Init();
            ConfigurationManager.AppSettings[Constants.AppConfig.KeyNames.SubjectToIgnore] = "Обед,Встреча,Блок,Встреча. Блок,Перерыв";
            ConfigurationManager.AppSettings[Constants.AppConfig.KeyNames.ProjectSearchPattern] = "%(.*?)%(.*)";
        }

        protected override ActivitiesDateCollection ReadActivities(TextReader reader)
        {
            var сalendarICSReader = new CalendarICSReader();
            return сalendarICSReader.ReadActivities(reader);
        }
    }
}
