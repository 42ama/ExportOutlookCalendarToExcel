using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace ExportOutlookCalendarToExcel.Library.WriteActivitiesIntoExcel
{
    /// <summary>
    /// Extensions for worksheet
    /// </summary>
    internal static class WorkSheetExtension
    {

        /// <summary>
        /// Get cell.
        /// </summary>
        /// <param name="sheet">Excel sheet.</param>
        /// <param name="letter">Cell letter.</param>
        /// <param name="index">Cell index.</param>
        internal static ExcelRange GetCell(this ExcelWorksheet sheet, char letter, int index)
        {
            Argument.Require(Char.IsLetter(letter), $"Character in argument {nameof(letter)} should be letter.");
            Argument.Require(index > 0, $"Index in argument {nameof(index)} should be greater then 0.");

            return sheet.Cells[$"{letter}{index}"];
        }

        /// <summary>
        /// Set value into cell.
        /// </summary>
        /// <param name="sheet">Excel sheet.</param>
        /// <param name="letter">Cell letter.</param>
        /// <param name="index">Cell index.</param>
        /// <param name="value">Cell value.</param>
        internal static void SetCellValue(this ExcelWorksheet sheet, char letter, int index, object value)
        {
            Argument.Require(Char.IsLetter(letter), $"Character in argument {nameof(letter)} should be letter.");
            Argument.Require(index > 0, $"Index in argument {nameof(index)} should be greater then 0.");

            sheet.SetValue($"{letter}{index}", value);
        }

        /// <summary>
        /// Set value into cell.
        /// </summary>
        /// <param name="sheet">Excel sheet.</param>
        /// <param name="letter">Cell letter.</param>
        /// <param name="index">Cell index.</param>
        /// <param name="formula">Formula.</param>
        internal static void SetFormula(this ExcelWorksheet sheet, char letter, int index, string formula)
        {
            Argument.Require(Char.IsLetter(letter), $"Character in argument {nameof(letter)} should be letter.");
            Argument.Require(index > 0, $"Index in argument {nameof(index)} should be greater then 0.");
            Argument.NotNullOrEmpty(formula, nameof(formula));

            sheet.Cells[$"{letter}{index}"].Formula = formula;
        }
    }
}
