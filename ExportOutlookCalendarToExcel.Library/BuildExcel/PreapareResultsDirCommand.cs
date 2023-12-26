using ExportOutlookCalendarToExcel.Library._Common.FilePathLocationStrategy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.BuildExcel
{
    public class PreapareResultsDirCommand
    {
        public PreapareResultsDirCommand(IFilePathLocationStrategy filepathLocationStrategy)
        {
            var resultsFolderPath = filepathLocationStrategy.GetDirLocation();
            var resultsDirectoryInfo = new DirectoryInfo(resultsFolderPath);
            if (!resultsDirectoryInfo.Exists)
            {
                resultsDirectoryInfo.Create();
            }

            ResultsDirectoryInfo = resultsDirectoryInfo;
        }

        public DirectoryInfo ResultsDirectoryInfo { get; private set; }
    }
}
