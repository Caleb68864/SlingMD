using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Turns user-entered tags into valid Obsidian tags: no leading <c>#</c>, spaces replaced by
    /// <c>-</c>, and only letters, digits, <c>_</c>, <c>-</c> and <c>/</c> kept. Empty and duplicate
    /// (case-insensitive) tags are dropped. Pure helper.
    /// </summary>
    internal static class TagNormalizer
    {
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+");
        private static readonly Regex InvalidCharsRegex = new Regex(@"[^\p{L}\p{N}_\-/]");

        public static List<string> Normalize(IEnumerable<string> tags)
        {
            List<string> result = new List<string>();
            if (tags == null)
            {
                return result;
            }

            foreach (string tag in tags)
            {
                string cleaned = (tag ?? string.Empty).Trim().TrimStart('#').Trim();
                cleaned = WhitespaceRegex.Replace(cleaned, "-");
                cleaned = InvalidCharsRegex.Replace(cleaned, string.Empty).Trim('-', '/');
                if (cleaned.Length > 0 && !result.Contains(cleaned, StringComparer.OrdinalIgnoreCase))
                {
                    result.Add(cleaned);
                }
            }

            return result;
        }
    }
}
