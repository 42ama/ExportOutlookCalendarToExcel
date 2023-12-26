using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.WriteActivitiesIntoExcel
{
    /// <summary>
    /// Базовая формула работающая с одной колонкой.
    /// </summary>
    internal class BasicFormula
    {
        internal int RangeStart { get; set; }
        internal int RangeEnd { get; set; }

        /// <summary>
        /// Буква колонки.
        /// </summary>
        internal char Letter { get; set; }

        /// <summary>
        /// Операция формулы.
        /// </summary>
        private string _formulaOperation;

        /// <summary>
        /// Операция формулы.
        /// </summary>
        internal string FormulaOperation
        {
            get { return _formulaOperation; }
            set
            {
                Argument.Require(Constants.FormulaOperations.IsAvailable(value),
                    $"Значение операции должно быть указано в {nameof(Constants.FormulaOperations)}");

                _formulaOperation = value;                
            }
        }

        /// <summary>
        /// Получить адрес диапазона формулы.
        /// </summary>
        /// <returns>Адрес диапазона формулы.</returns>
        internal string GetAddress()
        {
            return $"{Letter}{RangeStart}:{Letter}{RangeEnd}";
        }

        /// <summary>
        /// Получить значение формулы.
        /// </summary>
        /// <returns>Формула.</returns>
        internal string GetFormula()
        {
            var address = GetAddress();
            return $"={FormulaOperation}({address})";
        }
    }
}
