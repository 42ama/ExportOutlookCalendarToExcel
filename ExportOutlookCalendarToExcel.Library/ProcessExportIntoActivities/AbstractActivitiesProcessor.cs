using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.ProcessExportIntoActivities
{
    public abstract class AbstractActivitiesProcessor
    {
        protected string _filePath;

        public AbstractActivitiesProcessor(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)
                || !File.Exists(filePath))
            {
                throw new ArgumentNullException(nameof(filePath), $"Value {filePath} is invalid to pass on as file path.");
            }

            _filePath = filePath;
        }

        protected virtual string GetFilePathToReadFrom()
        {
            return _filePath;
        }

        protected virtual TextReader GetTextReader(string readFromPath)
        {
            return new StreamReader(readFromPath);
        }

        public abstract ActivitiesDateCollection ReadActivities();
    }
}
