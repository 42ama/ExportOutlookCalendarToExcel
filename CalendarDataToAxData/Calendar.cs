using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData
{
    public class Calendar
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
        public string Placy { get; set; }
        public string Description { get; set; }
        public string Note { get; set; }
        public string Distance { get; set; }
        public string PaymentMethod { get; set; }
        public string IsPrivate { get; set; }
    }

    public class CalendarClassMap : ClassMap<Calendar>
    {
        public CalendarClassMap()
        {
            Map(m => m.Subject).Name("Тема");
            Map(m => m.StartDate).Name("Дата начала");
            Map(m => m.StartTime).Name("Время начала");
            Map(m => m.EndDate).Name("Дата завершения");
            Map(m => m.EndTime).Name("Время завершения");
            Map(m => m.IsFullDay).Name("Целый день");
            Map(m => m.IsNotificationOn).Name("Напоминание вкл/выкл");
            Map(m => m.NotificationDate).Name("Дата напоминания");
            Map(m => m.NotificationTime).Name("Время напоминания");
            Map(m => m.Owner).Name("Организатор собрания");
            Map(m => m.ParticipantObligatory).Name("Обязательные участники");
            Map(m => m.ParticipantOptional).Name("Необязательные участники");
            Map(m => m.Resources).Name("Ресурсы собрания");
            Map(m => m.InThisTime).Name("В это время");
            Map(m => m.Importance).Name("Важность");
            Map(m => m.Category).Name("Категории");
            Map(m => m.Placy).Name("Место");
            Map(m => m.Description).Name("Описание");
            Map(m => m.Note).Name("Пометка");
            Map(m => m.Distance).Name("Расстояние");
            Map(m => m.PaymentMethod).Name("Способ оплаты");
            Map(m => m.IsPrivate).Name("Частное");
        }
    }
}
