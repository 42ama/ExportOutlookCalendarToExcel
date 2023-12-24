using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Library.Logic.FilepathLocationStrategy
{
    public class PersonalFilepathLocationStrategy : IFilepathLocationStrategy
    {
        public string GetDirLocation()
        {
            return $@"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\outlook-export";
        }
    }
}
