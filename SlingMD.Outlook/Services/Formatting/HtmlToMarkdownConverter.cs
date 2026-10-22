using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Converts the visible, commonly-used portions of an Outlook HTML body to Markdown.
    /// This intentionally avoids a third-party HTML parser because Outlook bodies can contain
    /// malformed email markup. Unsupported markup is removed while its readable text is retained.
    /// </summary>
    internal static class HtmlToMarkdownConverter
    {
        private const RegexOptions HtmlOptions = RegexOptions.IgnoreCase | RegexOptions.Singleline;
        private static readonly Regex CommentRegex = new Regex(@"<!--.*?-->", HtmlOptions);
        private static readonly Regex NonContentElementRegex = new Regex(
            @"<(head|style|script|noscript|template)\b[^>]*>.*?</\1\s*>",
            HtmlOptions);
        private static readonly Regex BodyRegex = new Regex(
            @"<body\b[^>]*>(.*?)</body\s*>",
            HtmlOptions);
        private static readonly Regex ImageRegex = new Regex(
            @"<img\b[^>]*?/?>",
            HtmlOptions);
        private static readonly Regex AnchorRegex = new Regex(
            @"<a\b[^>]*?\bhref\s*=\s*(?:(?:""([^""]*)"")|(?:'([^']*)')|([^\s>]+))[^>]*>(.*?)</a\s*>",
            HtmlOptions);
        private static readonly Regex HeadingRegex = new Regex(
            @"<h([1-6])\b[^>]*>(.*?)</h\1\s*>",
            HtmlOptions);
        private static readonly Regex ListItemRegex = new Regex(
            @"<li\b[^>]*>(.*?)</li\s*>",
            HtmlOptions);
        private static readonly Regex BlockquoteRegex = new Regex(
            @"<blockquote\b[^>]*>(.*?)</blockquote\s*>",
            HtmlOptions);
        private static readonly Regex LineBreakRegex = new Regex(
            @"<br\b[^>]*/?>",
            HtmlOptions);
        private static readonly Regex HorizontalRuleRegex = new Regex(
            @"<hr\b[^>]*/?>",
            HtmlOptions);
        // td/th are listed for tables the table pass couldn't parse (unclosed markup), so their
        // cells still don't run together.
        private static readonly Regex BlockOpeningRegex = new Regex(
            @"<(?:p|div|section|article|header|footer|address|figure|figcaption|table|thead|tbody|tfoot|tr|td|th|ul|ol)\b[^>]*>",
            HtmlOptions);
        private static readonly Regex BlockClosingRegex = new Regex(
            @"</(?:p|div|section|article|header|footer|address|figure|figcaption|table|thead|tbody|tfoot|tr|td|th|ul|ol)\s*>",
            HtmlOptions);
        // Innermost table first: the body may not contain another <table>.
        private static readonly Regex InnermostTableRegex = new Regex(
            @"<table\b[^>]*>((?:(?!<table\b).)*?)</table\s*>",
            HtmlOptions);
        private static readonly Regex TableRowRegex = new Regex(
            @"<tr\b[^>]*>(.*?)</tr\s*>",
            HtmlOptions);
        private static readonly Regex TableCellRegex = new Regex(
            @"<t([dh])\b[^>]*>(.*?)</t[dh]\s*>",
            HtmlOptions);
        private static readonly Regex SafeLinksRegex = new Regex(
            @"^https?://[^/?#]*\.safelinks\.protection\.outlook\.com/[^?#]*\?(?:[^#]*?&)?url=([^&#]+)",
            RegexOptions.IgnoreCase);
        private static readonly Regex StrongRegex = new Regex(
            @"<(?:strong|b)\b[^>]*>(.*?)</(?:strong|b)\s*>",
            HtmlOptions);
        private static readonly Regex EmphasisRegex = new Regex(
            @"<(?:em|i)\b[^>]*>(.*?)</(?:em|i)\s*>",
            HtmlOptions);
        private static readonly Regex AttributeRegex = new Regex(
            @"(?<![A-Za-z0-9_-])(?<name>[A-Za-z_:][A-Za-z0-9_.:-]*)\s*=\s*(?:(?:""(?<double>[^""]*)"")|(?:'(?<single>[^']*)')|(?<unquoted>[^\s>]+))",
            HtmlOptions);
        private static readonly Regex TagRegex = new Regex(@"<[^>]*>", HtmlOptions);
        private static readonly Regex RepeatedBlankLineRegex = new Regex(@"\n{3,}", RegexOptions.None);
        private static readonly Regex TrailingWhitespaceRegex = new Regex(@"[ \t]+\n", RegexOptions.None);
        private static readonly Regex RepeatedHorizontalWhitespaceRegex = new Regex(@"[ \t]{2,}", RegexOptions.None);

        /// <summary>
        /// Converts an HTML email body to Markdown. Plain text is returned when HTML is absent or
        /// contains no readable content after cleanup.
        /// </summary>
        public static string Convert(string html, string plainTextFallback)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return plainTextFallback ?? string.Empty;
            }

            string markdown = ConvertHtml(html);
            return string.IsNullOrWhiteSpace(markdown)
                ? (plainTextFallback ?? string.Empty)
                : markdown;
        }

        private static string ConvertHtml(string html)
        {
            string content = html.Replace("\r\n", "\n").Replace("\r", "\n");
            content = CommentRegex.Replace(content, string.Empty);
            content = NonContentElementRegex.Replace(content, string.Empty);

            Match bodyMatch = BodyRegex.Match(content);
            if (bodyMatch.Success)
            {
                content = bodyMatch.Groups[1].Value;
            }

            content = HorizontalRuleRegex.Replace(content, "\n\n---\n\n");
            content = LineBreakRegex.Replace(content, "\n");
            content = ImageRegex.Replace(content, match => ConvertImage(match.Value));
            content = AnchorRegex.Replace(content, match => ConvertAnchor(match));
            content = HeadingRegex.Replace(content, match => ConvertHeading(match));
            content = ListItemRegex.Replace(content, match => "\n- " + StripTags(ConvertInlineMarkup(match.Groups[1].Value)).Trim());
            content = BlockquoteRegex.Replace(content, match => ConvertBlockquote(match.Groups[1].Value));
            content = ConvertInlineMarkup(content);
            content = ConvertTables(content);
            content = BlockOpeningRegex.Replace(content, "\n\n");
            content = BlockClosingRegex.Replace(content, "\n\n");
            content = StripTags(content);
            content = WebUtility.HtmlDecode(content);

            return CleanWhitespace(content);
        }

        private static string ConvertHeading(Match match)
        {
            int level = int.Parse(match.Groups[1].Value);
            string headingText = StripTags(ConvertInlineMarkup(match.Groups[2].Value)).Trim();
            return "\n\n" + new string('#', level) + " " + headingText + "\n\n";
        }

        private static string ConvertInlineMarkup(string value)
        {
            string result = StrongRegex.Replace(value, match => "**" + StripTags(match.Groups[1].Value).Trim() + "**");
            result = EmphasisRegex.Replace(result, match => "*" + StripTags(match.Groups[1].Value).Trim() + "*");
            return result;
        }

        private static string ConvertAnchor(Match match)
        {
            string href = FirstNonEmpty(match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value);
            string linkText = StripTags(match.Groups[4].Value).Trim();
            href = UnwrapSafeLink(WebUtility.HtmlDecode(href ?? string.Empty).Trim());

            if (string.IsNullOrEmpty(href))
            {
                return linkText;
            }

            // Outlook often shows the wrapped address as the link text as well. A link whose text is
            // its own address is written as the bare address, which Obsidian links by itself.
            // (Entities are decoded by the final pass, so text is only decoded here to compare it.)
            string decodedText = UnwrapSafeLink(WebUtility.HtmlDecode(linkText));
            if (string.IsNullOrEmpty(linkText) || string.Equals(decodedText, href, StringComparison.OrdinalIgnoreCase))
            {
                return href;
            }

            return "[" + linkText + "](" + href + ")";
        }

        /// <summary>
        /// Returns the original address behind a Microsoft Defender Safe Links URL
        /// (<c>https://*.safelinks.protection.outlook.com/?url=...</c>), or <paramref name="url"/> unchanged.
        /// </summary>
        internal static string UnwrapSafeLink(string url)
        {
            if (string.IsNullOrEmpty(url))
            {
                return url;
            }

            Match match = SafeLinksRegex.Match(url);
            if (!match.Success)
            {
                return url;
            }

            string target = Uri.UnescapeDataString(match.Groups[1].Value);
            return Uri.IsWellFormedUriString(target, UriKind.Absolute) ? target : url;
        }

        /// <summary>
        /// Converts tables, innermost first. A table whose rows all have the same number (two or more)
        /// of single-line cells becomes a Markdown pipe table with the first row as its header. Anything
        /// else is layout (signatures, newsletters, Outlook's own wrappers): each row's single-line cells
        /// are joined with spaces, and cells with several lines become their own paragraphs.
        /// </summary>
        private static string ConvertTables(string content)
        {
            string previous;
            do
            {
                previous = content;
                content = InnermostTableRegex.Replace(content, match => ConvertTable(match.Groups[1].Value));
            }
            while (content != previous);

            return content;
        }

        private static string ConvertTable(string tableHtml)
        {
            List<List<string>> rows = new List<List<string>>();
            bool hasHeaderCells = false;
            foreach (Match rowMatch in TableRowRegex.Matches(tableHtml))
            {
                List<string> cells = new List<string>();
                foreach (Match cellMatch in TableCellRegex.Matches(rowMatch.Groups[1].Value))
                {
                    hasHeaderCells |= string.Equals(cellMatch.Groups[1].Value, "h", StringComparison.OrdinalIgnoreCase);
                    cells.Add(ConvertCell(cellMatch.Groups[2].Value));
                }

                if (cells.Any(cell => cell.Length > 0))
                {
                    rows.Add(cells);
                }
            }

            if (rows.Count == 0)
            {
                // Not a parsable table: keep its text for the block pass.
                return "\n\n" + tableHtml + "\n\n";
            }

            int columns = rows[0].Count;
            bool isDataTable = columns >= 2
                && (rows.Count >= 2 || hasHeaderCells)
                && rows.All(row => row.Count == columns && row.All(cell => cell.IndexOf('\n') < 0));

            StringBuilder result = new StringBuilder("\n\n");
            if (isDataTable)
            {
                result.Append(PipeRow(rows[0]));
                result.Append("|").Append(string.Concat(Enumerable.Repeat(" --- |", columns))).Append('\n');
                foreach (List<string> row in rows.Skip(1))
                {
                    result.Append(PipeRow(row));
                }
            }
            else
            {
                foreach (List<string> row in rows)
                {
                    List<string> cells = row.Where(cell => cell.Length > 0).ToList();
                    if (cells.All(cell => cell.IndexOf('\n') < 0))
                    {
                        result.Append(string.Join(" ", cells)).Append('\n');
                    }
                    else
                    {
                        result.Append("\n").Append(string.Join("\n\n", cells)).Append("\n\n");
                    }
                }
            }

            return result.Append('\n').ToString();
        }

        private static string ConvertCell(string cellHtml)
        {
            // Entities are decoded by the final pass; non-breaking spaces are turned into spaces here
            // so spacer cells (&nbsp;) count as empty.
            string text = BlockOpeningRegex.Replace(cellHtml, "\n\n");
            text = BlockClosingRegex.Replace(text, "\n\n");
            text = Regex.Replace(StripTags(text), "&nbsp;|&#160;|&#xa0;", " ", RegexOptions.IgnoreCase);
            return CleanWhitespace(text);
        }

        private static string PipeRow(List<string> cells)
        {
            return "| " + string.Join(" | ", cells.Select(cell => cell.Replace("|", "\\|"))) + " |\n";
        }

        private static string ConvertImage(string imageTag)
        {
            string source = GetAttribute(imageTag, "src");
            if (string.IsNullOrWhiteSpace(source))
            {
                return string.Empty;
            }

            string width = GetAttribute(imageTag, "width");
            string height = GetAttribute(imageTag, "height");
            string style = GetAttribute(imageTag, "style") ?? string.Empty;
            if (IsPixelDimension(width) || IsPixelDimension(height) || HasTrackingSource(source))
            {
                return string.Empty;
            }

            if (Regex.IsMatch(style, @"(?:width|height)\s*:\s*1(?:px)?(?:\s|;|$)", RegexOptions.IgnoreCase))
            {
                return string.Empty;
            }

            string alt = GetAttribute(imageTag, "alt") ?? string.Empty;
            return "\n\n![" + WebUtility.HtmlDecode(alt).Trim() + "](" + WebUtility.HtmlDecode(source).Trim() + ")\n\n";
        }

        private static string ConvertBlockquote(string value)
        {
            string text = CleanWhitespace(WebUtility.HtmlDecode(StripTags(ConvertInlineMarkup(value))));
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            StringBuilder result = new StringBuilder("\n\n");
            string[] lines = text.Split(new[] { '\n' }, StringSplitOptions.None);
            foreach (string line in lines)
            {
                result.Append("> ").Append(line.Trim()).Append('\n');
            }

            result.Append('\n');
            return result.ToString();
        }

        private static string GetAttribute(string tag, string name)
        {
            MatchCollection matches = AttributeRegex.Matches(tag);
            foreach (Match match in matches)
            {
                if (string.Equals(match.Groups["name"].Value, name, StringComparison.OrdinalIgnoreCase))
                {
                    return FirstNonEmpty(
                        match.Groups["double"].Value,
                        match.Groups["single"].Value,
                        match.Groups["unquoted"].Value);
                }
            }

            return null;
        }

        private static bool IsPixelDimension(string dimension)
        {
            if (string.IsNullOrWhiteSpace(dimension))
            {
                return false;
            }

            string normalized = dimension.Trim().ToLowerInvariant();
            return normalized == "1" || normalized == "1px";
        }

        private static bool HasTrackingSource(string source)
        {
            return source.IndexOf("/p.gif", StringComparison.OrdinalIgnoreCase) >= 0
                || source.IndexOf("/open?", StringComparison.OrdinalIgnoreCase) >= 0
                || source.IndexOf("tracking", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string StripTags(string value)
        {
            return string.IsNullOrEmpty(value) ? string.Empty : TagRegex.Replace(value, string.Empty);
        }

        private static string CleanWhitespace(string value)
        {
            string result = value.Replace('\u00A0', ' ');
            result = TrailingWhitespaceRegex.Replace(result, "\n");
            result = RepeatedHorizontalWhitespaceRegex.Replace(result, " ");
            result = RepeatedBlankLineRegex.Replace(result, "\n\n");
            return result.Trim();
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (string value in values)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    return value;
                }
            }

            return string.Empty;
        }
    }
}
