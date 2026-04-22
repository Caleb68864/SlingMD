using System;
using System.Globalization;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Provides centralized date-to-string conversion for email, contact, and appointment domains.
    /// All date placeholder rendering should go through this service.
    /// </summary>
    public class DateFormatter
    {
        /// <summary>
        /// Formats a DateTime value using the specified format string.
        /// </summary>
        /// <param name="value">The DateTime value to format.</param>
        /// <param name="format">The format string to apply.</param>
        /// <returns>The formatted date string.</returns>
        /// <exception cref="FormatException">Thrown when the format string is invalid.</exception>
        public string Format(DateTime value, string format)
        {
            return value.ToString(format, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Formats a DateTime value using the specified format string, returning the fallback
        /// if the format string is invalid.
        /// </summary>
        /// <param name="value">The DateTime value to format.</param>
        /// <param name="format">The format string to apply.</param>
        /// <param name="fallback">The fallback string to return if formatting fails.</param>
        /// <returns>The formatted date string, or the fallback if formatting fails.</returns>
        public string FormatOrDefault(DateTime value, string format, string fallback)
        {
            if (string.IsNullOrEmpty(format))
            {
                return fallback;
            }

            try
            {
                return value.ToString(format, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                return fallback;
            }
        }
    }
}
