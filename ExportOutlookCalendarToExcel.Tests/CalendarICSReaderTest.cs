using ExportOutlookCalendarToExcel.Logic.CalendarReader;
using System.Diagnostics;
using ExportOutlookCalendarToExcel.Logic;
using System.Configuration;
using ExportOutlookCalendarToExcel.Common;
using System;

namespace ExportOutlookCalendarToExcel.Tests
{
    [TestClass]
    public class CalendarICSReaderTest
    {
        [TestInitialize]
        public void Init()
        {
            ConfigurationManager.AppSettings[Constants.AppConfig.KeyNames.SubjectToIgnore] = "����,�������,����,�������. ����";
            ConfigurationManager.AppSettings[Constants.AppConfig.KeyNames.ProjectSearchPattern] = "%(.*?)%(.*)";
        }

        [TestMethod]
        public void ReaderCalendarFromPersonalAndConvert()
        {
            var a = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Personal));
            ;
        }

        
    }
}