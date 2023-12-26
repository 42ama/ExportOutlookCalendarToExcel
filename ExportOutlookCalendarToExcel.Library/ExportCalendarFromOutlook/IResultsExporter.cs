using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ExportCalendarFromOutlook
{
    public interface IResultsExporter
    {
        ExportedData Export(DirectoryInfo resultsDirectory, DateTime from, DateTime to);
    }
}
