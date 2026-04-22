using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that extracts individual fields from an Obsidian markdown file's
    /// YAML frontmatter using string regexes (no YAML parser dependency). Tolerant of
    /// the subset of frontmatter shapes SlingMD itself emits.
    /// </summary>
    public class FrontmatterReader
    {
        private static readonly Regex ThreadIdRegex = new Regex(@"threadId: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex DateRegex = new Regex(@"date: ""?(\d{4}-\d{2}-\d{2} \d{2}:\d{2}(?::\d{2})?)""?", RegexOptions.Compiled);
        private static readonly Regex TitleRegex = new Regex(@"title: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex FromRegex = new Regex(@"from: ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Compiled);
        private static readonly Regex ToRegex = new Regex(@"to:.*?\n\s*- ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex InternetMessageIdRegex = new Regex(@"internetMessageId: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex EntryIdRegex = new Regex(@"entryId: ""([^""]+)""", RegexOptions.Compiled);

        /// <summary>
        /// Returns the threadId frontmatter value, or null if not present.
        /// </summary>
        public string ExtractThreadId(string content)
        {
            return ExtractFirst(content, ThreadIdRegex);
        }

        /// <summary>
        /// Returns the internetMessageId frontmatter value, or null if not present.
        /// </summary>
        public string ExtractInternetMessageId(string content)
        {
            return ExtractFirst(content, InternetMessageIdRegex);
        }

        /// <summary>
        /// Returns the entryId frontmatter value, or null if not present.
        /// </summary>
        public string ExtractEntryId(string content)
        {
            return ExtractFirst(content, EntryIdRegex);
        }

        /// <summary>
        /// Returns the raw date string from the frontmatter (yyyy-MM-dd HH:mm or yyyy-MM-dd HH:mm:ss),
        /// or null if absent. Caller is responsible for parsing.
        /// </summary>
        public string ExtractRawDate(string content)
        {
            return ExtractFirst(content, DateRegex);
        }

        /// <summary>
        /// Returns the title frontmatter value, or null if not present.
        /// </summary>
        public string ExtractTitle(string content)
        {
            return ExtractFirst(content, TitleRegex);
        }

        /// <summary>
        /// Returns the wikilink target inside the frontmatter "from" field, or null.
        /// </summary>
        public string ExtractFromName(string content)
        {
            return ExtractFirst(content, FromRegex);
        }

        /// <summary>
        /// Returns the first wikilink target inside the frontmatter "to" list, or null.
        /// </summary>
        public string ExtractFirstToName(string content)
        {
            return ExtractFirst(content, ToRegex);
        }

        private static string ExtractFirst(string content, Regex regex)
        {
            if (string.IsNullOrEmpty(content))
            {
                return null;
            }
            Match m = regex.Match(content);
            return m.Success ? m.Groups[1].Value : null;
        }
    }
}
