using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData.Model
{
    /// <summary>
    /// Мапиинг CSV к классу <c>CalendarCSV</c>.
    /// </summary>
    public class CalendarCSVClassMap : ClassMap<CalendarCSV>
    {
        /// <summary>
        /// Конструктор.
        /// </summary>
        public CalendarCSVClassMap()
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
            Map(m => m.Place).Name("Место");
            Map(m => m.Description).Name("Описание");
            Map(m => m.Note).Name("Пометка");
            Map(m => m.Distance).Name("Расстояние");
            Map(m => m.PaymentMethod).Name("Способ оплаты");
            Map(m => m.IsPrivate).Name("Частное");
        }
    }
}
