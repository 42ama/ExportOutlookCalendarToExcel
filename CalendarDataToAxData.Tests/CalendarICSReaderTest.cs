using CalendarDataToAxData.Logic.CalendarReader;
using System.Diagnostics;
using CalendarDataToAxData.Logic;
using System.Configuration;
using CalendarDataToAxData.Common;
using System;

namespace CalendarDataToAxData.Tests
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
            var �alendarICSReader = new CalendarICSReader();
            var filePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\calender.ics";
            using var reader = new StreamReader(filePath);
            var activities = �alendarICSReader.ReadActivities(reader);
            var fileName = EPPlusExcelWriter.WriteToFile(activities, Environment.GetFolderPath(
                            System.Environment.SpecialFolder.DesktopDirectory)); 
            Console.WriteLine($"������! ������ ����: {fileName}");

            new Process { StartInfo = new ProcessStartInfo(fileName) { UseShellExecute = true } }.Start();
        }
    }
}