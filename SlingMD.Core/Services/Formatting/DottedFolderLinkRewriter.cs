using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Repairs links written before <see cref="FolderNameSanitizer"/> existed, when SlingMD linked to a
    /// folder by its dotted name (<c>[[Re_ Budget-Robert Smith Jr./file.pdf]]</c>) while Windows created
    /// the folder without the trailing dot (<c>...Robert Smith Jr</c>), so Obsidian cannot find it.
    /// <para>
    /// Only link targets are touched: wikilinks and embeds (<c>[[...]]</c>, <c>![[...]]</c>) and inline
    /// Markdown links and images (<c>[..](...)</c>, <c>![..](...)</c>, including <c>&lt;...&gt;</c> targets
    /// and URL-encoded ones). In a target, every folder segment (not the file name) loses the trailing
    /// dots and spaces Windows drops, written in the target's own style (<c>%20</c> stays <c>%20</c>).
    /// A target is rewritten only when the caller says the original does not resolve and the repaired
    /// one does, so a link that works is never changed. Aliases, headings, titles, external URLs, code
    /// spans and fenced code blocks are left as they are.
    /// </para>
    /// Pure helper: whether a path exists is decided by the caller through <c>resolves</c>.
    /// </summary>
    internal static class DottedFolderLinkRewriter
    {
        // Fenced code blocks (``` or ~~~, closed by the same fence or the end of the text) and code spans.
        private static readonly Regex CodeRegex = new Regex(
            @"(?m)^[ \t]{0,3}(?<fence>`{3,}|~{3,})[^\r\n]*(?:\r?\n[\s\S]*?(?:\r?\n[ \t]{0,3}\k<fence>[ \t]*(?=\r?$)|\z)|\z)"
            + @"|(?<!`)(?<tick>`+)(?!`)[^\r\n]*?(?<!`)\k<tick>(?!`)",
            RegexOptions.Compiled);

        // [[target]], [[target|alias]], [[target#heading]], with or without the embed "!".
        private static readonly Regex WikilinkRegex = new Regex(
            @"\[\[(?<inner>[^\[\]\r\n]+)\]\]",
            RegexOptions.Compiled);

        // ](<target> "title") or ](target "title"); the target may hold one level of balanced parentheses.
        private static readonly Regex MarkdownLinkRegex = new Regex(
            @"\]\((?:<(?<angle>[^<>\r\n]*)>|(?<plain>(?:[^()\r\n]|\([^()\r\n]*\))*))\)",
            RegexOptions.Compiled);

        private static readonly Regex TrailingTitleRegex = new Regex(
            @"\s+(?:""[^""]*""|'[^']*'|\([^()]*\))\s*$",
            RegexOptions.Compiled);

        private static readonly Regex SchemeRegex = new Regex(@"^[A-Za-z][A-Za-z0-9+.\-]*:", RegexOptions.Compiled);

        /// <summary>
        /// Returns <paramref name="content"/> with every repairable link target rewritten.
        /// </summary>
        /// <param name="content">Markdown note text.</param>
        /// <param name="resolves">Given a decoded link path (as written, '/'-separated, without alias or
        /// heading), whether it points at an existing file or folder. Must compare folder names exactly,
        /// i.e. "Jr." must not match a folder named "Jr".</param>
        public static string Rewrite(string content, Func<string, bool> resolves)
        {
            int rewritten;
            return Rewrite(content, resolves, out rewritten);
        }

        /// <summary>
        /// Same as <see cref="Rewrite(string, Func{string, bool})"/>, also returning how many link targets
        /// were rewritten.
        /// </summary>
        public static string Rewrite(string content, Func<string, bool> resolves, out int rewrittenCount)
        {
            rewrittenCount = 0;
            if (string.IsNullOrEmpty(content) || resolves == null || content.IndexOf('/') < 0)
            {
                return content;
            }

            List<KeyValuePair<int, int>> code = new List<KeyValuePair<int, int>>();
            foreach (Match match in CodeRegex.Matches(content))
            {
                code.Add(new KeyValuePair<int, int>(match.Index, match.Index + match.Length));
            }

            // (start, length, replacement) of each target to replace, in text order.
            List<Tuple<int, int, string>> edits = new List<Tuple<int, int, string>>();

            foreach (Match match in WikilinkRegex.Matches(content))
            {
                if (IsInside(code, match.Index))
                {
                    continue;
                }

                Group inner = match.Groups["inner"];
                int end = inner.Value.IndexOfAny(new[] { '|', '#' });
                int length = end < 0 ? inner.Length : end;
                // "[[a/b.pdf\|alias]]" inside a table: the backslash escapes the pipe, not the path.
                if (length > 0 && end >= 0 && inner.Value[end] == '|' && inner.Value[length - 1] == '\\')
                {
                    length--;
                }

                string replacement = RepairTarget(inner.Value.Substring(0, length), false, resolves);
                if (replacement != null)
                {
                    edits.Add(Tuple.Create(inner.Index, length, replacement));
                }
            }

            foreach (Match match in MarkdownLinkRegex.Matches(content))
            {
                if (IsInside(code, match.Index))
                {
                    continue;
                }

                Group target = match.Groups["angle"].Success ? match.Groups["angle"] : match.Groups["plain"];
                string value = target.Value;
                Match title = TrailingTitleRegex.Match(value);
                if (title.Success && title.Index > 0)
                {
                    value = value.Substring(0, title.Index);
                }

                int leading = value.Length - value.TrimStart().Length;
                value = value.Trim();
                int hash = value.IndexOf('#');
                string path = hash < 0 ? value : value.Substring(0, hash);

                string replacement = RepairTarget(path, true, resolves);
                if (replacement != null)
                {
                    edits.Add(Tuple.Create(target.Index + leading, path.Length, replacement));
                }
            }

            if (edits.Count == 0)
            {
                return content;
            }

            edits.Sort((a, b) => a.Item1.CompareTo(b.Item1));
            StringBuilder builder = new StringBuilder(content.Length);
            int position = 0;
            foreach (Tuple<int, int, string> edit in edits)
            {
                if (edit.Item1 < position)
                {
                    continue; // overlapping matches (a wikilink inside a Markdown target); keep the first
                }

                builder.Append(content, position, edit.Item1 - position);
                builder.Append(edit.Item3);
                position = edit.Item1 + edit.Item2;
                rewrittenCount++;
            }

            builder.Append(content, position, content.Length - position);
            return builder.ToString();
        }

        /// <summary>
        /// Returns the repaired form of one raw link path, or null when it should be left alone: it has
        /// no dotted folder segment, is an external URL, already resolves, or the repaired path does not.
        /// </summary>
        /// <param name="rawPath">Link path as written in the note, without alias, heading or title.</param>
        /// <param name="urlEncoded">True for Markdown link targets, whose %XX escapes are decoded before
        /// <paramref name="resolves"/> is asked; wikilink targets are taken literally.</param>
        internal static string RepairTarget(string rawPath, bool urlEncoded, Func<string, bool> resolves)
        {
            if (string.IsNullOrEmpty(rawPath) || rawPath.IndexOf('/') < 0 || SchemeRegex.IsMatch(rawPath))
            {
                return null;
            }

            string[] segments = rawPath.Split('/');
            bool changed = false;
            // The last segment is the file (or note) name, which keeps its dot.
            for (int i = 0; i < segments.Length - 1; i++)
            {
                string segment = segments[i];
                if (segment.Length == 0 || segment == "." || segment == "..")
                {
                    continue;
                }

                string trimmed = TrimTrailingDotsAndSpaces(segment, urlEncoded);
                if (trimmed == segment)
                {
                    continue;
                }

                // Nothing left ("...") or a device name ("CON."): Windows could not have created that
                // folder under the dotted name, so there is nothing to repair it to.
                string decodedTrimmed = Decode(trimmed, urlEncoded);
                if (decodedTrimmed.Trim().Length == 0
                    || FolderNameSanitizer.Sanitize(decodedTrimmed) != decodedTrimmed)
                {
                    continue;
                }

                segments[i] = trimmed;
                changed = true;
            }

            if (!changed)
            {
                return null;
            }

            string repaired = string.Join("/", segments);
            if (resolves(Decode(rawPath, urlEncoded)) || !resolves(Decode(repaired, urlEncoded)))
            {
                return null;
            }

            return repaired;
        }

        /// <summary>Strips trailing '.', ' ' and, in URL-encoded text, their %2E / %20 escapes.</summary>
        private static string TrimTrailingDotsAndSpaces(string segment, bool urlEncoded)
        {
            string result = segment;
            while (result.Length > 0)
            {
                char last = result[result.Length - 1];
                if (last == '.' || last == ' ')
                {
                    result = result.Substring(0, result.Length - 1);
                }
                else if (urlEncoded && result.Length >= 3 && result[result.Length - 3] == '%'
                    && (string.Equals(result.Substring(result.Length - 2), "20", StringComparison.Ordinal)
                        || string.Equals(result.Substring(result.Length - 2), "2E", StringComparison.OrdinalIgnoreCase)))
                {
                    result = result.Substring(0, result.Length - 3);
                }
                else
                {
                    break;
                }
            }

            return result;
        }

        private static string Decode(string value, bool urlEncoded)
        {
            if (!urlEncoded)
            {
                return value;
            }

            try
            {
                return Uri.UnescapeDataString(value);
            }
            catch (UriFormatException)
            {
                return value;
            }
        }

        private static bool IsInside(List<KeyValuePair<int, int>> ranges, int index)
        {
            foreach (KeyValuePair<int, int> range in ranges)
            {
                if (index >= range.Key && index < range.Value)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
