using ExportOutlookCalendarToExcel.Common;
using ExportOutlookCalendarToExcel.Logic.CalendarReader;
using ExportOutlookCalendarToExcel.Model;
using CsvHelper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Logic.ResultBuilder
{
    public abstract class AbstractResultBuilder
    {
        public void Build()
        {
            Init();
            Validate();

            var readFromPath = GetFilePathToReadFrom();

            using var textReader = GetTextReader(readFromPath);
            var activities = ReadActivities(textReader);

            var resultFileDirPath = GetResultFileDirPath();
            var excelFilePath = CreateExcel(activities, resultFileDirPath);

            Finalize(excelFilePath);
        }

        public virtual void Init()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public virtual void Validate()
        {
            // Валидируем App.Config
            if (!AppConfigProvider.IsSettingExists(Constants.AppConfig.KeyNames.SubjectToIgnore))
            {
                Console.WriteLine($"В App.config должен быть ключ {Constants.AppConfig.KeyNames.SubjectToIgnore}");
                return;
            }

            if (!AppConfigProvider.IsSettingExists(Constants.AppConfig.KeyNames.ProjectSearchPattern))
            {
                Console.WriteLine($"В App.config должен быть ключ {Constants.AppConfig.KeyNames.ProjectSearchPattern}");
                return;
            }
        }
        public abstract string GetFilePathToReadFrom();
        public abstract TextReader GetTextReader(string readFromPath);
        public abstract ActivitiesDateCollection ReadActivities(TextReader reader);
        public abstract string GetResultFileDirPath();
        public virtual string CreateExcel(ActivitiesDateCollection activities, string resultFileDirPath)
        {
            var fileName = EPPlusExcelWriter.WriteToFile(activities, resultFileDirPath);
            Console.WriteLine($"Готово! Создан файл: {fileName}");

            return fileName;
        }
        public virtual void Finalize(string resultFilePath)
        {
            new Process { StartInfo = new ProcessStartInfo(resultFilePath) { UseShellExecute = true } }.Start();
        }
    }
}
