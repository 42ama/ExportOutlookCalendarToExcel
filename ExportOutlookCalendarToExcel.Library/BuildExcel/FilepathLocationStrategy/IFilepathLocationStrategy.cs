using ExportOutlookCalendarToExcel.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Logic.FilepathLocationStrategy
{
    public interface IFilepathLocationStrategy
    {
        string GetDirLocation();
    }
}
