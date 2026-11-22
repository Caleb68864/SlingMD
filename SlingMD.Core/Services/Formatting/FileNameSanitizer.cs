using System;
using System.Collections.Generic;
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
        /// The characters Windows forbids in a file name: exactly what
        /// <see cref="Path.GetInvalidFileNameChars"/> returns on Windows (<c>" &lt; &gt; |</c>, NUL and
        /// the other control characters, <c>: * ? \ /</c>). Spelled out because the vault is always on
        /// Windows, while on Linux and macOS that method returns only NUL and '/', which made the
        /// result depend on the machine running the tests.
        /// </summary>
        private static readonly char[] WindowsInvalidFileNameChars = BuildWindowsInvalidFileNameChars();

        private static char[] BuildWindowsInvalidFileNameChars()
        {
            List<char> chars = new List<char> { '"', '<', '>', '|' };
            for (int c = 0; c < 32; c++)
            {
                chars.Add((char)c);
            }
            chars.AddRange(new[] { ':', '*', '?', '\\', '/' });
            return chars.ToArray();
        }

        /// <summary>
        /// Reserved Windows device names. A file named exactly one of these (any casing, with or
        /// without an extension) cannot be created — the OS routes it to the device — so a subject
        /// that reduces to e.g. "CON" or "NUL" would otherwise abort the sling with an I/O error.
        /// </summary>
        private static readonly HashSet<string> ReservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
            // Windows also treats the superscript digits 1-3 as digits here (COM¹, LPT²...).
            "COM\u00B9", "COM\u00B2", "COM\u00B3",
            "LPT\u00B9", "LPT\u00B2", "LPT\u00B3"
        };

        /// <summary>
        /// Whether Windows treats <paramref name="name"/> as a device: a reserved name on its own or
        /// followed by an extension ("NUL.txt", "Con.Ed bill"). Windows looks at the part before the
        /// first dot with trailing spaces dropped.
        /// </summary>
        internal static bool IsReservedDeviceName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            int dot = name.IndexOf('.');
            string stem = (dot >= 0 ? name.Substring(0, dot) : name).TrimEnd(' ');
            return ReservedNames.Contains(stem);
        }

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
            cleaned = string.Join("_", cleaned.Split(WindowsInvalidFileNameChars, StringSplitOptions.RemoveEmptyEntries));

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

            // Trim residual leading/trailing separators. Dots are intentionally NOT trimmed:
            // legitimate names end in a dot ("Robert Smith Jr.") and this value doubles as a
            // display/contact name, so stripping it would corrupt name matching.
            cleaned = cleaned.Trim('-', '_');

            // Guard against reserved device names, with or without an extension.
            if (IsReservedDeviceName(cleaned))
            {
                cleaned = "_" + cleaned;
            }

            return cleaned;
        }
    }
}
