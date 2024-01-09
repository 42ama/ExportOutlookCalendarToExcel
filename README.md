## Export Outlook calendar to Excel
This VSTO Outlook plugin enhances the functionality of Outlook by allowing users to export their calendar to Excel. It introduces a new group in the Add-ins section where user can select a date period to be exported to Excel. The available predefined periods include:
- **Today**.
- **Week** - data from the start to the end of the current week.
- **Week to today** - data from the start of the current week to today.
- **Month** - data from the start to the end of the current month.
- **Month to today** - data from the start of the current month to today.

Users also have the option to select the **Choose date range** option and manually specify the period of dates to be exported.

### Groups

Events can grouped while processing. To select an event within a group, its subject must contain two distinct parts, separated by a semicolon `;`. For example, the subject could be `Cool project; Meeting with friends`.

## Installation
1. Download [latest](https://github.com/42ama/ExportOutlookCalendarToExcel/releases/tag/latest) release.
2. Install with *setup.exe*

## Unsupported (features for future releases)
- Export for recurring events other then daily and weekly.
- Export for events with several recurrance rules. 

## Licensed Libraries Used in This Project
- EPPlus (under https://polyformproject.org/licenses/noncommercial/1.0.0/)
