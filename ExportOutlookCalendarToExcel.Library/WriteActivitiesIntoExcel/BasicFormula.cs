using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExportOutlookCalendarToExcel.Library.WriteActivitiesIntoExcel
{
    /// <summary>
    /// Base formula which works with single column.
    /// </summary>
    internal class BasicFormula
    {
        /// <summary>
        /// Number of first row in formula.
        /// </summary>
        internal int RangeStart { get; set; }

        /// <summary>
        /// Number of last row in formula.
        /// </summary>
        internal int RangeEnd { get; set; }

        /// <summary>
        /// Column letter.
        /// </summary>
        internal char Letter { get; set; }

        /// <summary>
        /// Operation in formula.
        /// </summary>
        private string _formulaOperation;

        /// <summary>
        /// Operation in formula.
        /// </summary>
        internal string FormulaOperation
        {
            get { return _formulaOperation; }
            set
            {
                Argument.Require(Constants.FormulaOperations.IsAvailable(value),
                    $"Operation value should set in {nameof(Constants.FormulaOperations)}");

                _formulaOperation = value;                
            }
        }

        /// <summary>
        /// Get formula range address.
        /// </summary>
        /// <returns>Formula range address.</returns>
        internal string GetAddress()
        {
            return $"{Letter}{RangeStart}:{Letter}{RangeEnd}";
        }

        /// <summary>
        /// Get formula value.
        /// </summary>
        /// <returns>Formula value.</returns>
        internal string GetFormula()
        {
            var address = GetAddress();
            return $"={FormulaOperation}({address})";
        }
    }
}
