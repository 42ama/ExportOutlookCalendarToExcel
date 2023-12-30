using System;
using System.Linq;
using System.Text;

namespace ExportOutlookCalendarToExcel._Common
{
    internal static class Constants
    {
        internal static class FileInfo
        {
            internal static class CSV
            {
                /// <summary>
                /// Extension.
                /// </summary>
                internal static string CsvExtension { get; } = ".csv";

                /// <summary>
                /// Suffix for encoded temp file.
                /// </summary>
                internal static string TempFileSuffix { get; } = "_temp_encoded";

                /// <summary>
                /// Target encoding for CSVReader.
                /// </summary>
                internal static Encoding TargetEncoding { get; } = Encoding.UTF8;
            }            
        }
        

        /// <summary>
        /// Настройки активности.
        /// </summary>
        internal static class ActivitySettings
        {

            /// <summary>
            /// Длина поиска разделителя в строке.
            /// </summary>
            internal static int SubstringSeparatorSearchLenght { get; } = 15;

            /// <summary>
            /// Разделитель для массивов в конфигурации.
            /// </summary>
            internal static char Separator { get; } = '.';

            /// <summary>
            /// Fallback для темы, если она пустая.
            /// </summary>
            internal static string SubjectFallback { get; } = "<пустая тема>";
        }

        /// <summary>
        /// Настройки работы с файлом конфигурации.
        /// </summary>
        internal static class AppConfig
        {
            /// <summary>
            /// Разделитель для массивов в конфигурации.
            /// </summary>
            internal static char ArraySeparator { get; } = ',';

            /// <summary>
            /// Значения ключей файла конфигурации.
            /// </summary>
            internal static class KeyNames
            {
                /// <summary>
                /// Темы для игнорирования.
                /// </summary>
                internal static string SubjectToIgnore { get; } = "SubjectToIgnore";

                /// <summary>
                /// Паттерн для поиска проекта в Теме.
                /// </summary>
                internal static string GroupSearchPattern { get; } = "GroupSearchPattern";
            }
        }

        /// <summary>
        /// Настройки работы с Excel'ем.
        /// </summary>
        internal static class Excel
        {
            /// <summary>
            /// Номер первой строки со значениями.
            /// </summary>
            internal static int FirstValueRowNumber { get; } = 2;

            /// <summary>
            /// Настройки строки заголовков.
            /// </summary>
            internal static class Header
            {
                /// <summary>
                /// Номер строки.
                /// </summary>
                internal static int RowNumber { get; } = 1;
            }

            /// <summary>
            /// Настройки колонки Длительность.
            /// </summary>
            internal static class Duration
            {
                /// <summary>
                /// Буква колонки.
                /// </summary>
                internal static char Letter { get; } = 'C';

                /// <summary>
                /// Название колонки.
                /// </summary>
                internal static string Name { get; } = "Длительность";
            }

            /// <summary>
            /// Настройки колонки Группа.
            /// </summary>
            internal static class Group
            {
                /// <summary>
                /// Буква колонки.
                /// </summary>
                internal static char Letter { get; } = 'A';

                /// <summary>
                /// Название колонки.
                /// </summary>
                internal static string Name { get; } = "Группа";
            }

            /// <summary>
            /// Настройки колонки Тема.
            /// </summary>
            internal static class Subject
            {
                /// <summary>
                /// Буква колонки.
                /// </summary>
                internal static char Letter { get; } = 'B';

                /// <summary>
                /// Название колонки.
                /// </summary>
                internal static string Name { get; } = "Тема";

                /// <summary>
                /// Номер колонки.
                /// </summary>
                internal static int ColumnNumber { get; } = 2;

                /// <summary>
                /// Максимальная длина колонки.
                /// </summary>
                internal static int ColumnLength { get; } = 122;
            }
        }

        /// <summary>
        /// Операции с формулами.
        /// </summary>
        internal static class FormulaOperations
        {
            /// <summary>
            /// Сумма.
            /// </summary>
            internal static string Sum { get; } = "Sum";

            /// <summary>
            /// Перечень доступных операций.
            /// </summary>
            internal static string[] AvailableOperations { get; } =
                new string[]
                {
                    Constants.FormulaOperations.Sum
                };

            /// <summary>
            /// Проверяет доступна ли операция для формулы.
            /// </summary>
            /// <param name="operation">Операция.</param>
            /// <returns>true - если операция доступа для формулы.</returns>
            internal static bool IsAvailable(string operation)
            {
                return AvailableOperations.Contains(operation);
            }
        }
    }
}
