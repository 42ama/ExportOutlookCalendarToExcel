using ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Tests
{
    [TestClass]
    public class Activity_GroupTests
    {
        [TestMethod]
        public void Subject_SetOnlyInSubject()
        {
            var subject = "Test subject set only in subject";


            var activity = new Activity(subject
                                        , DateTime.Now
                                        ,DateTime.Now.AddHours(1)
                                        ,false);


            Assert.AreEqual(subject, activity.Subject);
        }


        [TestMethod]
        public void Subject_SetInSubjectAndGroup()
        {
            var group = "Test group";
            var subject = "Test subject set in subject";

            var activity = new Activity($"{group};{subject}"
                                        , DateTime.Now
                                        , DateTime.Now.AddHours(1)
                                        , false);

            Assert.AreEqual(group, activity.Group);
            Assert.AreEqual(subject, activity.Subject);
        }
    }
}
