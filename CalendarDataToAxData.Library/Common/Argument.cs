using System;
using System.Collections.Generic;
using System.Text;

namespace CalendarDataToAxData.Common
{
	/// <summary>
	/// Класс для проверки аргументов.
	/// </summary>
	public class Argument
    {
		/// <summary>
		/// Ensures that argument is not null.
		/// </summary>
		/// <param name="value">Pass argument you need to check.</param>
		/// <param name="message">Error message for the exception.</param>
		/// <exception cref="ArgumentNullException"><paramref name="value" /> is <c>null</c>.</exception>
		public static void NotNull(object value, string paramName, string message = "")
		{
			if (null == value)
			{
				throw new ArgumentNullException(paramName, message);
			}
		}

		/// <summary>
		/// Ensures that string argument is not null and is not empty.
		/// </summary>
		/// <param name="value">Pass argument you need to check.</param>
		/// <param name="message">Error message for the exception.</param>
		/// <exception cref="ArgumentException"><paramref name="value" /> is <c>null</c>.</exception>
		public static void NotNullOrEmpty(string value, string paramName, string message = "")
		{
			if (string.IsNullOrEmpty(value))
			{
				throw new ArgumentException(message, paramName);
			}
		}

		/// <summary>
		/// Perform custom check on the argument.
		/// </summary>
		/// <param name="condition">Condition to check.</param>
		/// <param name="message">Error message for the exception.</param>
		/// <exception cref="ArgumentNullException"><paramref name="condition" /> is <c>false</c>.</exception>
		public static void Require(bool condition, string message)
		{
			if (!condition)
			{
				throw new ArgumentException(message);
			}
		}
	}
}
