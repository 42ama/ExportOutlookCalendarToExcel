using ExportOutlookCalendarToExcel.Common;
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace ExportOutlookCalendarToExcel.Extension
{
    /// <summary>
    /// Extensions for worksheet
    /// </summary>
    internal static class WorkSheetExtension
    {
        /// <summary>
        /// Установить значение в ячейку
        /// </summary>
        /// <param name="sheet">Данный лист.</param>
        /// <param name="letter">Буква ячейки</param>
        /// <param name="index">Номер ячейки</param>
        /// <param name="value">Новое значение</param>
        public static void SetCellValue(this ExcelWorksheet sheet, char letter, int index, object value)
        {
            Argument.Require(Char.IsLetter(letter), $"Для переменной {nameof(letter)} ожидается буква.");
            Argument.Require(index > 0, $"Для индекса {nameof(index)} ожидается значение больше 0.");

            sheet.SetValue($"{letter}{index}", value);
        }

        /// <summary>
        /// Установить формулу в ячейку
        /// </summary>
        /// <param name="sheet">Данный лист.</param>
        /// <param name="letter">Буква ячейки</param>
        /// <param name="index">Номер ячейки</param>
        /// <param name="formula">Формула</param>
        public static void SetFormula(this ExcelWorksheet sheet, char letter, int index, string formula)
        {
            Argument.Require(Char.IsLetter(letter), $"Для переменной {nameof(letter)} ожидается буква.");
            Argument.Require(index > 0, $"Для переменной {nameof(index)} ожидается значение больше 0.");
            Argument.NotNullOrEmpty(formula, nameof(formula));

            sheet.Cells[$"{letter}{index}"].Formula = formula;
        }
    }
}
