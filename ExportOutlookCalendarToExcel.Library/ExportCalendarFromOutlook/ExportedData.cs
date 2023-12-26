using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ExportCalendarFromOutlook
{
    public class ExportedData
    {
        public string FilePath { get; }

        public string ParentDirectoryFilePath { get; }

        public ExportedData(string filePath)
        {
            FilePath = filePath;
            ParentDirectoryFilePath = filePath.Substring(0, filePath.LastIndexOf('\\'));
        }
    }
}
