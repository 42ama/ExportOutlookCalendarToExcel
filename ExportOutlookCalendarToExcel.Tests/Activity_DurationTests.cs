using ExportOutlookCalendarToExcel.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Tests
{
    [TestClass]
    public class Activity_DurationTests
    {
        [TestMethod]
        [DataRow(60, 1)]
        [DataRow(30, 0.5)]
        [DataRow(15, 0.25)]
        [DataRow(25, 0.5)]
        public void Duration_ParsedAsHours(int minutes, double hours)
        {
            var from = DateTime.Now.AddMinutes(-minutes);
            var to = DateTime.Now;

            var activity = new Activity($"{minutes} mins -> {hours} hours", from, to, false);

            Assert.AreEqual(hours, activity.Duration);
        }
    }
}
