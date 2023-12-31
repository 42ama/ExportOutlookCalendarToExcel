using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.WriteActivitiesIntoExcel
{
    /// <summary>
    /// Constants specific for work in Excel.
    /// </summary>
    internal static class ExcelConstants
    {
        /// <summary>
        /// Index of first row with values.
        /// </summary>
        internal static int FirstValueRowIndex { get; } = 2;

        /// <summary>
        /// Header column settings..
        /// </summary>
        internal static class Header
        {
            /// <summary>
            /// Index of header row.
            /// </summary>
            internal static int RowIndex { get; } = 1;
        }

        /// <summary>
        /// Duration column settings.
        /// </summary>
        internal static class Duration
        {
            /// <summary>
            /// Duration column letter.
            /// </summary>
            internal static char Letter { get; } = 'C';
        }

        /// <summary>
        /// Group column settings.
        /// </summary>
        internal static class Group
        {
            /// <summary>
            /// Group column letter.
            /// </summary>
            internal static char Letter { get; } = 'A';

        }

        /// <summary>
        /// Subject column settings.
        /// </summary>
        internal static class Subject
        {
            /// <summary>
            /// Subject column letter.
            /// </summary>
            internal static char Letter { get; } = 'B';


            /// <summary>
            /// Subject column index.
            /// </summary>
            internal static int ColumnIndex { get; } = 2;

            /// <summary>
            /// Subject column length.
            /// </summary>
            internal static int ColumnLength { get; } = 122;
        }
    }
}
