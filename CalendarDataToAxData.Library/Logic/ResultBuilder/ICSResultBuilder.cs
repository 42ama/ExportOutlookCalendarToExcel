using CalendarDataToAxData.Common;
using CalendarDataToAxData.Logic.CalendarReader;
using CalendarDataToAxData.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarDataToAxData.Logic.ResultBuilder
{
    public class ICSResultBuilder : AbstractResultBuilder
    {
        public override void Init()
        {
            base.Init();
            ConfigurationManager.AppSettings[Constants.AppConfig.KeyNames.SubjectToIgnore] = "Обед,Встреча,Блок,Встреча. Блок";
            ConfigurationManager.AppSettings[Constants.AppConfig.KeyNames.ProjectSearchPattern] = "%(.*?)%(.*)";
        }

        public override void Validate()
        {
            // Убираем базовую валидацию
        }

        public override string GetFilePathToReadFrom()
        {
            return Constants.FileInfo.ICS.FilePath;
        }

        public override string GetResultFileDirPath()
        {
            return Constants.FileInfo.ICS.DirPath;
        }

        public override TextReader GetTextReader(string readFromPath)
        {
            return new StreamReader(readFromPath);
        }

        public override ActivitiesDateCollection ReadActivities(TextReader reader)
        {
            var сalendarICSReader = new CalendarICSReader();
            return сalendarICSReader.ReadActivities(reader);
        }
    }
}
