using System;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel._Common
{
    public static class Constants
    {
        public static class FileInfo
        {
            public static class CSV
            {
                /// <summary>
                /// Extension.
                /// </summary>
                public static string CsvExtension { get; } = ".csv";

                /// <summary>
                /// Suffix for encoded temp file.
                /// </summary>
                public static string TempFileSuffix { get; } = "_temp_encoded";

                /// <summary>
                /// Target encoding for CSVReader.
                /// </summary>
                public static Encoding TargetEncoding { get; } = Encoding.UTF8;
            }            
        }
        

        /// <summary>
        /// Настройки активности.
        /// </summary>
        public static class ActivitySettings
        {
            /// <summary>
            /// Префикс для Темы Встречи.
            /// </summary>
            public static string MeetingPrefix { get; } = "Встреча. ";

            /// <summary>
            /// Длина поиска разделителя в строке.
            /// </summary>
            public static int SubstringSeparatorSearchLenght { get; } = 15;

            /// <summary>
            /// Разделитель для массивов в конфигурации.
            /// </summary>
            public static char Separator { get; } = '.';

            /// <summary>
            /// Fallback для темы, если она пустая.
            /// </summary>
            public static string SubjectFallback { get; } = "<пустая тема>";
        }

        /// <summary>
        /// Настройки работы с файлом конфигурации.
        /// </summary>
        public static class AppConfig
        {
            /// <summary>
            /// Разделитель для массивов в конфигурации.
            /// </summary>
            public static char ArraySeparator { get; } = ',';

            /// <summary>
            /// Значения ключей файла конфигурации.
            /// </summary>
            public static class KeyNames
            {
                /// <summary>
                /// Темы для игнорирования.
                /// </summary>
                public static string SubjectToIgnore { get; } = "SubjectToIgnore";

                /// <summary>
                /// Паттерн для поиска проекта в Теме.
                /// </summary>
                public static string ProjectSearchPattern { get; } = "ProjectSearchPattern";
            }
        }

        /// <summary>
        /// Настройки работы с Excel'ем.
        /// </summary>
        public static class Excel
        {
            /// <summary>
            /// Номер первой строки со значениями.
            /// </summary>
            public static int FirstValueRowNumber { get; } = 2;

            /// <summary>
            /// Настройки строки заголовков.
            /// </summary>
            public static class Header
            {
                /// <summary>
                /// Номер строки.
                /// </summary>
                public static int RowNumber { get; } = 1;
            }

            /// <summary>
            /// Настройки колонки Длительность.
            /// </summary>
            public static class Duration
            {
                /// <summary>
                /// Буква колонки.
                /// </summary>
                public static char Letter { get; } = 'C';

                /// <summary>
                /// Название колонки.
                /// </summary>
                public static string Name { get; } = "Длительность";
            }

            /// <summary>
            /// Настройки колонки Проект.
            /// </summary>
            public static class Project
            {
                /// <summary>
                /// Буква колонки.
                /// </summary>
                public static char Letter { get; } = 'A';

                /// <summary>
                /// Название колонки.
                /// </summary>
                public static string Name { get; } = "Проект";
            }

            /// <summary>
            /// Настройки колонки Тема.
            /// </summary>
            public static class Subject
            {
                /// <summary>
                /// Буква колонки.
                /// </summary>
                public static char Letter { get; } = 'B';

                /// <summary>
                /// Название колонки.
                /// </summary>
                public static string Name { get; } = "Тема";

                /// <summary>
                /// Номер колонки.
                /// </summary>
                public static int ColumnNumber { get; } = 2;

                /// <summary>
                /// Максимальная длина колонки.
                /// </summary>
                public static int ColumnLength { get; } = 122;
            }
        }

        /// <summary>
        /// Операции с формулами.
        /// </summary>
        public static class FormulaOperations
        {
            /// <summary>
            /// Сумма.
            /// </summary>
            public static string Sum { get; } = "Sum";

            /// <summary>
            /// Перечень доступных операций.
            /// </summary>
            public static string[] AvailableOperations { get; } =
                new string[]
                {
                    Constants.FormulaOperations.Sum
                };

            /// <summary>
            /// Проверяет доступна ли операция для формулы.
            /// </summary>
            /// <param name="operation">Операция.</param>
            /// <returns>true - если операция доступа для формулы.</returns>
            public static bool IsAvailable(string operation)
            {
                return AvailableOperations.Contains(operation);
            }
        }
    }
}
