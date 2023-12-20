using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Logic
{
    public class PreapareResultsDirCommand
    {
        public PreapareResultsDirCommand()
        {
            var personalFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            var resultsFolderPath = personalFolderPath + @"\outlook-export";
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
