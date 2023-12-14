using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarDataToAxData.Library.Logic
{
    public class DeleteExcelFiles
    {
        private const string EXCEL_EXTENSION = ".xlsx";
        private const string ICS_EXTENSION = ".ics";
        private string _folderPath;

        public DeleteExcelFiles(string folderPath)
        {
            var isDirectory = Directory.Exists(folderPath);

            if (!isDirectory)
            {
                throw new ArgumentException($"{nameof(folderPath)} should be directory!", nameof(folderPath));
            }

            _folderPath = folderPath;
        }

        public void Delete()
        {
            var files = Directory.GetFiles(_folderPath);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                if (filePath.EndsWith(EXCEL_EXTENSION) || filePath.EndsWith(ICS_EXTENSION))
                {
                    File.Delete(filePath);
                }
            }
        }
    }
}
