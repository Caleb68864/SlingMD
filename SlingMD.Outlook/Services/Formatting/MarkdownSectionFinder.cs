using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper for locating a Markdown section heading line (e.g. "## Notes") in
    /// document content. Matching is anchored to start-of-line, multiline, and the
    /// heading text is regex-escaped so headings containing regex metacharacters work.
    /// </summary>
    public class MarkdownSectionFinder
    {
        /// <summary>
        /// Returns the character index of the line that begins with <paramref name="heading"/>,
        /// or -1 if the heading is not found at-or-after <paramref name="startIndex"/>.
        /// The heading must be on its own line (trailing whitespace allowed) for the match
        /// to count.
        /// </summary>
        public int FindSectionStart(string content, string heading, int startIndex = 0)
        {
            if (string.IsNullOrEmpty(content) || startIndex < 0 || startIndex >= content.Length)
            {
                return -1;
            }
            if (string.IsNullOrEmpty(heading))
            {
                return -1;
            }

            Match match = Regex.Match(
                content.Substring(startIndex),
                $"^{Regex.Escape(heading)}\\s*$",
                RegexOptions.Multiline);
            return match.Success ? startIndex + match.Index : -1;
        }
    }
}
