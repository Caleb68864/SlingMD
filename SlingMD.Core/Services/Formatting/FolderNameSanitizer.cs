using System;
using System.Collections.Generic;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Makes an already-sanitised name safe to use as a folder name on Windows.
    /// <para>
    /// <see cref="FileNameSanitizer.Sanitize"/> keeps a trailing dot on purpose: its output doubles as a
    /// contact/display name ("Robert Smith Jr.") and a note name, where ".md" follows. A folder has no
    /// extension, and Windows silently drops trailing dots and spaces from a directory name, so a
    /// folder created as <c>Inbox\Re_ Budget-Robert Smith Jr.</c> lands on disk as
    /// <c>...Robert Smith Jr</c> while SlingMD keeps using the dotted path for links and later file
    /// writes. Every place that turns such a name into a folder must pass it through
    /// <see cref="Sanitize"/> first.
    /// </para>
    /// Pure helper with no Outlook or filesystem dependencies.
    /// </summary>
    internal static class FolderNameSanitizer
    {
        private static readonly char[] TrailingCharsWindowsDrops = { '.', ' ' };

        /// <summary>Same reserved device names <see cref="FileNameSanitizer"/> guards against.</summary>
        private static readonly HashSet<string> ReservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        /// <summary>Name used when nothing is left after trimming (e.g. a name made only of dots).</summary>
        public const string EmptyNameReplacement = "_";

        /// <summary>
        /// Removes the trailing dots and spaces Windows would drop from a directory name, and guards the
        /// result against becoming empty or a reserved device name ("CON." is "CON" on disk). A null or
        /// empty name is returned as an empty string; a name that is already safe is returned unchanged.
        /// </summary>
        public static string Sanitize(string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                return string.Empty;
            }

            string trimmed = folderName.TrimEnd(TrailingCharsWindowsDrops);
            if (trimmed.Length == 0)
            {
                // "...", " . " and the like: as a path segment "..." would resolve to the parent.
                return EmptyNameReplacement;
            }

            if (ReservedNames.Contains(trimmed))
            {
                return "_" + trimmed;
            }

            return trimmed;
        }
    }
}
