using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Logic
{
    /// <summary>
    /// Delete previeous excel files in resutls directory.
    /// </summary>
    public class DeleteExcelFiles
    {
        private const string EXCEL_EXTENSION = ".xlsx";
        private const string ICS_EXTENSION = ".ics";

        private readonly DirectoryInfo _directoryWithExcelsInfo;

        public DeleteExcelFiles(DirectoryInfo directoryWithExcelsInfo)
        {
            if (!directoryWithExcelsInfo.Exists)
            {
                throw new ArgumentException($"{nameof(directoryWithExcelsInfo)} should be created", nameof(directoryWithExcelsInfo));
            }

            _directoryWithExcelsInfo = directoryWithExcelsInfo;
        }

        public void Delete()
        {
            var files = _directoryWithExcelsInfo.GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i].FullName;
                if (filePath.EndsWith(EXCEL_EXTENSION) || filePath.EndsWith(ICS_EXTENSION))
                {
                    File.Delete(filePath);
                }
            }
        }
    }
}
