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
using ExportOutlookCalendarToExcel.Library.Logic.FilepathLocationStrategy;

namespace ExportOutlookCalendarToExcel.Logic.ResultBuilder
{
    public abstract class AbstractResultBuilder
    {
        protected string _filePath;

        public AbstractResultBuilder(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)
                || !File.Exists(filePath))
            {
                throw new ArgumentNullException(nameof(filePath), $"Value {filePath} is invalid to pass on as file path.");
            }

            _filePath = filePath;
        }

        public void Build()
        {
            Init();
            Validate();

            var readFromPath = GetFilePathToReadFrom();

            using (var textReader = GetTextReader(readFromPath))
            {
                var activities = ReadActivities(textReader);

                var resultFileDirPath = GetResultFileDirPath();
                var excelFilePath = CreateExcel(activities, resultFileDirPath);

                Finalize(excelFilePath);
            }                
        }

        protected virtual void Init()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        protected virtual void Validate()
        {
            //// Валидируем App.Config
            //if (!AppConfigProvider.IsSettingExists(Constants.AppConfig.KeyNames.SubjectToIgnore))
            //{
            //    Console.WriteLine($"В App.config должен быть ключ {Constants.AppConfig.KeyNames.SubjectToIgnore}");
            //    return;
            //}

            //if (!AppConfigProvider.IsSettingExists(Constants.AppConfig.KeyNames.ProjectSearchPattern))
            //{
            //    Console.WriteLine($"В App.config должен быть ключ {Constants.AppConfig.KeyNames.ProjectSearchPattern}");
            //    return;
            //}
        }

        protected virtual string GetFilePathToReadFrom()
        {
            return _filePath;
        }

        protected virtual string GetResultFileDirPath()
        {
            return _filePath.Substring(0, _filePath.LastIndexOf('\\'));
        }

        protected virtual TextReader GetTextReader(string readFromPath)
        {
            return new StreamReader(readFromPath);
        }

        protected abstract ActivitiesDateCollection ReadActivities(TextReader reader);

        protected virtual string CreateExcel(ActivitiesDateCollection activities, string resultFileDirPath)
        {
            var fileName = EPPlusExcelWriter.WriteToFile(activities, resultFileDirPath);
            Console.WriteLine($"Готово! Создан файл: {fileName}");

            return fileName;
        }
        protected virtual void Finalize(string resultFilePath)
        {
            new Process { StartInfo = new ProcessStartInfo(resultFilePath) { UseShellExecute = true } }.Start();
        }
    }
}
