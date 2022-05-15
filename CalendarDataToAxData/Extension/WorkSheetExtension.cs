using IronXL;
using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData.Extension
{
    internal static class WorkSheetExtension
    {
        public static void SetValue(this WorkSheet sheet, char letter, int index, object value)
        {
            sheet[letter + index.ToString()].Value = value;
        }
    }
}
