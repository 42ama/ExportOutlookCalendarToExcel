using ExportOutlookCalendarToExcel.Logic.ResultBuilder;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Logic
{
    public class ExportAndConvertCalendarProcess
    {
        private readonly DirectoryInfo _resultsDirectoryInfo;

        public ExportAndConvertCalendarProcess(DirectoryInfo resultsDirectoryInfo)
        {
            _resultsDirectoryInfo = resultsDirectoryInfo;
        }

        public void Process(IResultsExporter exporter, DateTime from, DateTime to)
        {
            try
            {
                var deleteExcelFiles = new DeleteExcelFiles(_resultsDirectoryInfo);
                deleteExcelFiles.Delete();
            }
            catch (IOException ex)
            {
                // Can't delete file because one of them is open.
            }


            var exported = exporter.Export(_resultsDirectoryInfo, from, to);
            if (!exported.IsSuccsess)
            {
                return;
            }

            var icsResultBuilder = new ICSResultBuilder(exported.FilePath);
            icsResultBuilder.Build();
        }
    }
}
