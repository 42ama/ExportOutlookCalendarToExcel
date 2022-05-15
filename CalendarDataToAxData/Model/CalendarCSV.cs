using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CalendarDataToAxData.Model
{
    public class CalendarCSV
    {
        public string Subject { get; set; }
        public string StartDate { get; set; }
        public DateTime StartTime { get; set; }
        public string EndDate { get; set; }
        public DateTime EndTime { get; set; }
        public string IsFullDay { get; set; }
        public string IsNotificationOn { get; set; }
        public string NotificationDate { get; set; }
        public DateTime NotificationTime { get; set; }
        public string Owner { get; set; }
        public string ParticipantObligatory { get; set; }
        public string ParticipantOptional { get; set; }
        public string Resources { get; set; }
        public int InThisTime { get; set; }
        public string Importance { get; set; }
        public string Category { get; set; }
        public string Place { get; set; }
        public string Description { get; set; }
        public string Note { get; set; }
        public string Distance { get; set; }
        public string PaymentMethod { get; set; }
        public string IsPrivate { get; set; }


        public bool IsMeeting
        {
            get
            {
                var isAnyParticipants = !string.IsNullOrEmpty(ParticipantObligatory);
                var isPlaceSet = !string.IsNullOrEmpty(Place);

                return isAnyParticipants || isPlaceSet;
            }
        }

        public override string ToString()
        {
            return this.GetType().GetProperties()
                .Select(info => (info.Name, Value: info.GetValue(this, null) ?? "(null)"))
                .Aggregate(
                    new StringBuilder(),
                    (sb, pair) => sb.AppendLine($"{pair.Name}: {pair.Value}"),
                    sb => sb.ToString());
        }
    }
}
