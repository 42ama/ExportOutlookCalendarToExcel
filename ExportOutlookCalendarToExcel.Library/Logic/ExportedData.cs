using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Logic
{
    public class ExportedData
    {
        public string FilePath { get; }

        public bool IsSuccsess { get; }

        public ExportedData(string filePath)
        {
            IsSuccsess = true;
            FilePath = filePath;
        }

        public ExportedData()
        {
            IsSuccsess = false;
        }
    }
}
