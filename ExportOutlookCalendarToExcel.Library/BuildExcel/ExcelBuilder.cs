﻿using ExportOutlookCalendarToExcel.Common;
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
    public class ExcelBuilder
    {
        protected string _directoryPath;

        public ExcelBuilder(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath)
                || !Directory.Exists(directoryPath))
            {
                throw new ArgumentNullException(nameof(directoryPath), $"Value {directoryPath} is invalid to pass on as directory path.");
            }

            _directoryPath = directoryPath;
        }

        public void Build(ActivitiesDateCollection activities)
        {
            Init();

            var excelFilePath = CreateExcel(activities, _directoryPath);

            OpenExcelFile(excelFilePath);
        }

        private void Init()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private string CreateExcel(ActivitiesDateCollection activities, string resultFileDirPath)
        {
            var fileName = EPPlusExcelWriter.WriteToFile(activities, resultFileDirPath);
            Console.WriteLine($"Готово! Создан файл: {fileName}");

            return fileName;
        }
        private void OpenExcelFile(string resultFilePath)
        {
            new Process { StartInfo = new ProcessStartInfo(resultFilePath) { UseShellExecute = true } }.Start();
        }
    }
}
