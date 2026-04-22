using System;
using System.IO;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that converts raw subject/name strings into filesystem-safe filename segments.
    /// Applies the post-subject-cleanup normalization (invalid char stripping, prefix removal,
    /// separator collapse) without any disk I/O or settings dependencies.
    /// </summary>
    internal class FileNameSanitizer
    {
        private static readonly Regex LeadingPrefixRegex = new Regex(@"^(?:RE_|FWD_|FW_|Re_|Fwd_)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex SeparatorRunRegex = new Regex(@"[-_]{2,}", RegexOptions.Compiled);

        /// <summary>
        /// Sanitizes <paramref name="input"/> for use as a filename: replaces invalid filename
        /// characters with underscores, strips quotes/colons/semicolons, removes a leading
        /// Re_/Fwd_ prefix, and collapses repeated separators.
        /// </summary>
        public string Sanitize(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            string cleaned = input;

            // Replace invalid filename characters with underscore.
            char[] invalidChars = Path.GetInvalidFileNameChars();
            cleaned = string.Join("_", cleaned.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

            // Strip / replace additional problematic characters.
            cleaned = cleaned.Replace("\"", string.Empty)
                             .Replace("'", string.Empty)
                             .Replace("`", string.Empty)
                             .Replace(":", "_")
                             .Replace(";", string.Empty)
                             .Trim();

            // Drop any remaining email-prefix stem that may have been converted to underscore form.
            cleaned = LeadingPrefixRegex.Replace(cleaned, string.Empty);

            // Collapse runs of separators.
            cleaned = SeparatorRunRegex.Replace(cleaned, "-");

            // Trim residual leading/trailing separators.
            return cleaned.Trim('-', '_');
        }
    }
}
