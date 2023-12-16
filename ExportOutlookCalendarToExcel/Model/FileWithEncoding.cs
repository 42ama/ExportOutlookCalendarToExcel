using ExportOutlookCalendarToExcel.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportOutlookCalendarToExcel.Model
{
    /// <summary>
    /// Файл прочитанный в нужной кодировке.
    /// </summary>
    internal class FileWithEncoding : IDisposable
    {
        /// <summary>
        /// Поток чтения из файла
        /// </summary>
        public StreamReader StreamReader { get; init; }

        /// <summary>
        /// Строка по которой хранится временный файл
        /// </summary>
        private readonly string _tempFilePath;

        public FileWithEncoding(string filePath, Encoding encoding)
        {
            // Outlook эксопртирует файл в 1251, однако почему-то CSVReader не может прочитать 1251, даже если её явно указать. Поэтому конвертируем в UTF-8
            var text = File.ReadAllText(filePath, Encoding.GetEncoding("windows-1251"));
            _tempFilePath = filePath.Insert(filePath.IndexOf(Constants.FileInfo.CsvExtension), Constants.FileInfo.TempFileSuffix);
            File.WriteAllText(_tempFilePath, text, encoding);
            StreamReader = new StreamReader(_tempFilePath);
        }

        public void Dispose()
        {
            StreamReader.Dispose();
            File.Delete(_tempFilePath);
        }
    }
}
