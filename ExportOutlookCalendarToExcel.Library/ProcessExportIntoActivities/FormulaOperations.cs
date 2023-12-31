using ExportOutlookCalendarToExcel._Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    /// <summary>
    /// Operation used in formulas.
    /// </summary>
    internal class FormulaOperations
    {
        /// <summary>
        /// Sum formula.
        /// </summary>
        internal static string Sum { get; } = "Sum";

        /// <summary>
        /// List of available operations.
        /// </summary>
        internal static string[] AvailableOperations { get; } =
            new string[]
            {
                    FormulaOperations.Sum
            };

        /// <summary>
        /// Is operation available for usage
        /// </summary>
        /// <param name="operation">Operation.</param>
        /// <returns>True if operation available for usage.</returns>
        internal static bool IsAvailable(string operation)
        {
            return AvailableOperations.Contains(operation);
        }
    }
}
