using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that extracts individual fields from an Obsidian markdown file's
    /// YAML frontmatter using string regexes (no YAML parser dependency). Tolerant of
    /// the subset of frontmatter shapes SlingMD itself emits.
    /// </summary>
    internal class FrontmatterReader
    {
        private static readonly Regex ThreadIdRegex = new Regex(@"threadId: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex DateRegex = new Regex(@"date: ""?(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?)""?", RegexOptions.Compiled);
        private static readonly Regex TitleRegex = new Regex(@"title: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex FromRegex = new Regex(@"from: ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Compiled);
        private static readonly Regex ToRegex = new Regex(@"to:.*?\n\s*- ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex InternetMessageIdRegex = new Regex(@"internetMessageId: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex EntryIdRegex = new Regex(@"entryId: ""([^""]+)""", RegexOptions.Compiled);
        private static readonly Regex AliasesBlockRegex = new Regex(@"aliases:\s*\r?\n((?:[ \t]+-[ \t]+.+\r?\n?)+)", RegexOptions.Compiled);
        private static readonly Regex AliasesInlineRegex = new Regex(@"aliases:\s*\[([^\]]*)\]", RegexOptions.Compiled);
        private static readonly Regex AliasesBlockItemRegex = new Regex(@"[ \t]+-[ \t]+(.+)", RegexOptions.Compiled);
        private static readonly Regex EmailBlockRegex = new Regex(@"^emails?:[ \t]*\r?\n((?:[ \t]+-[ \t]+.+\r?\n?)+)", RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex EmailInlineRegex = new Regex(@"^emails?:[ \t]*\[([^\]]*)\]", RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex EmailScalarRegex = new Regex(@"^emails?:[ \t]*(?![\[\s])([^\r\n]+?)\r?$", RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);

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

        /// <summary>
        /// Returns all aliases from the frontmatter aliases field, supporting both
        /// block-style (dash list) and inline-array shapes. Returns an empty list
        /// when the key is absent, the file has no leading ---, or parsing fails.
        /// </summary>
        public IReadOnlyList<string> ExtractAliases(string content)
        {
            if (string.IsNullOrEmpty(content))
                return Array.Empty<string>();

            if (!content.TrimStart().StartsWith("---"))
                return Array.Empty<string>();

            Match blockMatch = AliasesBlockRegex.Match(content);
            if (blockMatch.Success)
            {
                string block = blockMatch.Groups[1].Value;
                List<string> items = new List<string>();
                foreach (Match itemMatch in AliasesBlockItemRegex.Matches(block))
                {
                    string val = itemMatch.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(val))
                        items.Add(val);
                }
                return items.AsReadOnly();
            }

            Match inlineMatch = AliasesInlineRegex.Match(content);
            if (inlineMatch.Success)
            {
                string[] parts = inlineMatch.Groups[1].Value.Split(',');
                List<string> items = new List<string>();
                foreach (string part in parts)
                {
                    string val = part.Trim();
                    if (!string.IsNullOrEmpty(val))
                        items.Add(val);
                }
                return items.AsReadOnly();
            }

            return Array.Empty<string>();
        }

        /// <summary>
        /// Returns the addresses in the <c>email</c> or <c>emails</c> field (a single value, an inline
        /// array or a dash list), without quotes or a <c>mailto:</c> prefix. Values without an @ are skipped.
        /// </summary>
        public IReadOnlyList<string> ExtractEmails(string content)
        {
            List<string> emails = new List<string>();
            if (string.IsNullOrEmpty(content) || !content.TrimStart().StartsWith("---"))
                return emails.AsReadOnly();

            foreach (Match block in EmailBlockRegex.Matches(content))
            {
                foreach (Match item in AliasesBlockItemRegex.Matches(block.Groups[1].Value))
                    AddEmail(emails, item.Groups[1].Value);
            }

            foreach (Match inline in EmailInlineRegex.Matches(content))
            {
                foreach (string part in inline.Groups[1].Value.Split(','))
                    AddEmail(emails, part);
            }

            foreach (Match scalar in EmailScalarRegex.Matches(content))
                AddEmail(emails, scalar.Groups[1].Value);

            return emails.AsReadOnly();
        }

        private static void AddEmail(List<string> emails, string raw)
        {
            string value = (raw ?? string.Empty).Trim().Trim('"', '\'').Trim();
            if (value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                value = value.Substring("mailto:".Length);
            if (value.IndexOf('@') > 0 && !emails.Contains(value, StringComparer.OrdinalIgnoreCase))
                emails.Add(value);
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
