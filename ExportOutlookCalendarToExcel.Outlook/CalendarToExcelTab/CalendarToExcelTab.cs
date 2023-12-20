﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using ExportOutlookCalendarToExcel.Logic.ResultBuilder;
using ExportOutlookCalendarToExcel.Library.Logic;
using ExportOutlookCalendarToExcel.Common;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new AmaTab();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ExportOutlookCalendarToExcel.Outlook
{
    [ComVisible(true)]
    public class CalendarToExcelTab : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        private readonly ExportAndConvertCalendarProcess _exportAndConvertCalendarProcess;

        public CalendarToExcelTab()
        {
            var preapareResultsDirCommand  = new PreapareResultsDirCommand();
            _exportAndConvertCalendarProcess = new ExportAndConvertCalendarProcess(preapareResultsDirCommand.ResultsDirectoryInfo);
        }

        public void OnChooseDateRangeButtonClick(Office.IRibbonControl control)
        {
            var currentWeekStrategy = new ChooseDateStrategy_WeekToToday();
            var prompt = new ProcessedUserPrompt(currentWeekStrategy);
            if (!prompt.ShouldContinue)
            {
                return;
            }

            ExportAndConvertCalendar(prompt.From, prompt.To);
        }

        #region OnButtonClick

        public void OnTodayButtonClick(Office.IRibbonControl control)
        {
            var todayStrategy = new ChooseDateStrategy_Today();
            ExportAndConvertCalendar(todayStrategy.From, todayStrategy.To);
        }

        public void OnWeekToTodayButtonClick(Office.IRibbonControl control)
        {
            var weekToTodayStartegy = new ChooseDateStrategy_WeekToToday();
            ExportAndConvertCalendar(weekToTodayStartegy.From, weekToTodayStartegy.To);
        }

        public void OnWeekButtonClick(Office.IRibbonControl control)
        {
            var weekStrategy = new ChooseDateStrategy_Week();
            ExportAndConvertCalendar(weekStrategy.From, weekStrategy.To);
        }

        public void OnMonthToTodayClick(Office.IRibbonControl control)
        {
            var monthToTodayStrategy = new ChooseDateStrategy_MonthToToday();
            ExportAndConvertCalendar(monthToTodayStrategy.From, monthToTodayStrategy.To);
        }

        public void OnMonthClick(Office.IRibbonControl control)
        {
            var monthStrategy = new ChooseDateStrategy_Month();
            ExportAndConvertCalendar(monthStrategy.From, monthStrategy.To);
        }

        #endregion OnButtonClick

        public void ExportAndConvertCalendar(DateTime from, DateTime to)
        {
            var exportIcalCommand = new ExportICalFromOutlookCommand();
            _exportAndConvertCalendarProcess.Process(exportIcalCommand, from, to);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExportOutlookCalendarToExcel.Outlook.CalendarToExcelTab.CalendarToExcelTab.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this._ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
