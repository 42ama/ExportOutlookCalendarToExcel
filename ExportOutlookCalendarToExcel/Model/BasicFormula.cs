using System;
using System.Collections.Generic;
using System.Text;
using ExportOutlookCalendarToExcel.Common;

namespace ExportOutlookCalendarToExcel.Model
{
    /// <summary>
    /// Базовая формула работающая с одной колонкой.
    /// </summary>
    internal class BasicFormula
    {
        /// <summary>
        /// Диапозон ячеек.
        /// </summary>
        public Range Range { get; set; }

        /// <summary>
        /// Буква колонки.
        /// </summary>
        public char Letter { get; set; }

        /// <summary>
        /// Операция формулы.
        /// </summary>
        private string _formulaOperation;

        /// <summary>
        /// Операция формулы.
        /// </summary>
        public string FormulaOperation
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
        public string GetAddress()
        {
            return $"{Letter}{Range.Start.Value}:{Letter}{Range.End.Value}";
        }

        /// <summary>
        /// Получить значение формулы.
        /// </summary>
        /// <returns>Формула.</returns>
        public string GetFormula()
        {
            var address = GetAddress();
            return $"={FormulaOperation}({address})";
        }
    }
}
