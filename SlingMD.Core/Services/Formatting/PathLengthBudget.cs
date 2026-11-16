using System;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that shortens file and folder names so the full path stays inside the Windows
    /// path limits. Win32 MAX_PATH is 260 characters including the terminating NUL, and a directory
    /// must stay under 248. Outlook.exe is not long-path aware, so the LongPathsEnabled registry
    /// switch does not lift these limits for the add-in; the budget keeps a margin below both
    /// (the atomic write in <see cref="FileService.WriteUtf8File"/> briefly adds ".slingtmp").
    /// </summary>
    internal class PathLengthBudget
    {
        public const int MaxFilePath = 240;
        public const int MaxDirectoryPath = 230;

        /// <summary>Shortest name a truncation produces, however deep the folder is.</summary>
        public const int MinNameLength = 8;

        /// <summary>Room left for an attachment's own file name inside a per-note attachment folder.</summary>
        public const int AttachmentNameReserve = 30;

        private static readonly char[] WordBreaks = { ' ', '-', '_' };
        private static readonly char[] TrailingJunk = { ' ', '-', '_', '.' };

        /// <summary>
        /// Returns the longest name that keeps <c>directory\name + suffix</c> within <see cref="MaxFilePath"/>.
        /// </summary>
        public int MaxNameLength(string directory, string suffix)
        {
            return MaxFilePath - (directory ?? string.Empty).Length - 1 - (suffix ?? string.Empty).Length;
        }

        /// <summary>
        /// Shortens <paramref name="name"/> to at most <paramref name="maxLength"/> characters (never below
        /// <see cref="MinNameLength"/>), cutting at a word break when one falls in the last third, and
        /// trimming trailing spaces, dots and separators, which Windows rejects or Obsidian renders badly.
        /// Names that already fit are returned unchanged.
        /// </summary>
        public string Fit(string name, int maxLength)
        {
            if (string.IsNullOrEmpty(name))
            {
                return name ?? string.Empty;
            }

            int limit = Math.Max(maxLength, MinNameLength);
            if (name.Length <= limit)
            {
                return name;
            }

            string cut = name.Substring(0, limit);
            int boundary = cut.LastIndexOfAny(WordBreaks);
            if (boundary >= limit * 2 / 3)
            {
                cut = cut.Substring(0, boundary);
            }

            cut = cut.TrimEnd(TrailingJunk);
            return cut.Length > 0 ? cut : name.Substring(0, limit);
        }

        /// <summary>
        /// Shortens <paramref name="name"/> so that <c>directory\name + suffix</c> fits <see cref="MaxFilePath"/>.
        /// </summary>
        public string FitInDirectory(string directory, string name, string suffix)
        {
            return Fit(name, MaxNameLength(directory, suffix));
        }
    }
}
