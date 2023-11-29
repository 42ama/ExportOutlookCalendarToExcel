using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using CalendarDataToAxData.Common;
using System.Xml.Linq;

namespace CalendarDataToAxData
{
    /// <summary>
    /// Провайдер для работы с конфигурацией.
    /// </summary>
    public static class AppConfigProvider
    {
        /// <summary>
        /// Получить массив строк по ключу из файла конфигурации.
        /// </summary>
        /// <param name="key">Ключ в конфигурации.</param>
        /// <returns>Массив строк.</returns>
        public static string[] GetStringArray(string key)
        {
            Argument.NotNullOrEmpty(key, nameof(key));

            return ConfigurationManager.AppSettings[key].Split(Constants.AppConfig.ArraySeparator);
        }

        /// <summary>
        /// Получить значение по ключу из файла конфигурации.
        /// </summary>
        /// <param name="key">Ключ в конфигурации.</param>
        /// <returns>Значение из конфиграции.</returns>
        public static string Get(string key)
        {
            Argument.NotNullOrEmpty(key, nameof(key));

            return ConfigurationManager.AppSettings[key];
        }

        public static bool IsSettingExists(string key)
        {
            Argument.NotNullOrEmpty(key, nameof(key));

            return ConfigurationManager.AppSettings[key] != null;
        }
    }
}
