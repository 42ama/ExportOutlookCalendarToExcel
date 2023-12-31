using System;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel._Common
{
    /// <summary>
    /// Common constants which can be used across library.
    /// </summary>
    internal static class CommonConstants
    {
        /// <summary>
        /// Information about files.
        /// </summary>
        internal static class FileInfo
        {
            /// <summary>
            /// Information about Excel files.
            /// </summary>
            internal static class Excel
            {
                /// <summary>
                /// Excel files extension.
                /// </summary>
                internal readonly static string FileExtenstion = ".xlsx";
            }

            /// <summary>
            /// Information about iCalendar files.
            /// </summary>
            internal static class ICalendar
            {
                /// <summary>
                /// iCalendar files extension.
                /// </summary>
                internal readonly static string FileExtenstion = ".ics";
            }
        }

        
    }
}
