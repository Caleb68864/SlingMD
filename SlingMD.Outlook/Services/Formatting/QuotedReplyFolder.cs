using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Separates a reply from the quoted history underneath it (the "From: … Sent: …" block Outlook
    /// adds, Gmail and Apple Mail quotes, "-----Original Message-----", "On … wrote:") and converts
    /// the email to Markdown with that history folded into a collapsed Obsidian callout, so a thread
    /// note shows only what that message added while nothing is lost.
    /// Pure helper — no Outlook dependencies.
    /// </summary>
    internal static class QuotedReplyFolder
    {
        public const string CalloutTitle = "Earlier messages";

        private const RegexOptions HtmlOptions = RegexOptions.IgnoreCase | RegexOptions.Singleline;

        private static readonly Regex[] HtmlMarkers =
        {
            // Outlook on the web / new Outlook: marker div placed right before the reply header.
            new Regex(@"<div\b[^>]*\bid\s*=\s*[""']?appendonsend\b", HtmlOptions),
            // Outlook reply/forward header, usually preceded by a separator line.
            new Regex(@"(?:<hr\b[^>]*>\s*)?<div\b[^>]*\bid\s*=\s*[""']?divRplyFwdMsg\b", HtmlOptions),
            // Outlook desktop: the thin top border above "From: … Sent: …".
            new Regex(@"<div\b[^>]*border-top\s*:\s*solid\s+#(?:E1E1E1|B5C4DF)\b[^>]*>", HtmlOptions),
            // Gmail.
            new Regex(@"<div\b[^>]*\bclass\s*=\s*[""']?[^""'>]*\bgmail_quote\b", HtmlOptions),
            // Apple Mail and others.
            new Regex(@"<blockquote\b[^>]*\btype\s*=\s*[""']?cite\b", HtmlOptions)
        };

        private static readonly Regex[] TextMarkers =
        {
            new Regex(@"^[ \t]*-{2,}[ \t]*Original Message[ \t]*-{2,}[ \t]*\r?$", RegexOptions.IgnoreCase | RegexOptions.Multiline),
            new Regex(@"^[ \t]*From:[ \t].+\r?\n[ \t]*(?:Sent|Date):[ \t]", RegexOptions.IgnoreCase | RegexOptions.Multiline),
            new Regex(@"^[ \t]*On[ \t].+wrote:[ \t]*\r?$", RegexOptions.IgnoreCase | RegexOptions.Multiline)
        };

        private static readonly Regex BodyOpenRegex = new Regex(@"<body\b[^>]*>", HtmlOptions);
        private static readonly Regex LeadingRuleRegex = new Regex(@"\A(?:\s*---\s*)+");

        /// <summary>
        /// Converts the email to Markdown like <see cref="HtmlToMarkdownConverter.Convert"/>, but with the
        /// quoted history folded into a collapsed <c>[!quote]-</c> callout. When no quoted history is
        /// found, or nothing is left above it, the plain conversion is returned.
        /// </summary>
        /// <param name="clean">Optional; applied to the reply and to the quoted history before the history
        /// is wrapped in the callout (whose lines start with <c>&gt; </c>), so line-anchored patterns still
        /// match inside it.</param>
        public static string ConvertFolded(string html, string plainText, Func<string, string> clean = null)
        {
            Func<string, string> apply = clean ?? (text => text);
            string reply;
            string quoted;
            if (!string.IsNullOrWhiteSpace(html))
            {
                int index = FindEarliest(html, HtmlMarkers, StartOfBody(html));
                if (index < 0)
                {
                    return apply(HtmlToMarkdownConverter.Convert(html, plainText));
                }

                reply = HtmlToMarkdownConverter.Convert(html.Substring(0, index), null);
                // The separator line Outlook puts above the header adds nothing inside the callout.
                quoted = LeadingRuleRegex.Replace(HtmlToMarkdownConverter.Convert(html.Substring(index), null), string.Empty);
            }
            else
            {
                string text = plainText ?? string.Empty;
                int index = FindEarliest(text, TextMarkers, 0);
                if (index < 0)
                {
                    return apply(text);
                }

                reply = text.Substring(0, index).Trim();
                quoted = text.Substring(index).Trim();
            }

            if (string.IsNullOrWhiteSpace(reply))
            {
                return apply(HtmlToMarkdownConverter.Convert(html, plainText));
            }

            reply = apply(reply);
            quoted = apply(quoted);
            return string.IsNullOrWhiteSpace(quoted) ? reply : reply + "\n\n" + ToCollapsedCallout(quoted);
        }

        /// <summary>Wraps <paramref name="markdown"/> in a collapsed Obsidian quote callout.</summary>
        internal static string ToCollapsedCallout(string markdown)
        {
            StringBuilder callout = new StringBuilder("> [!quote]- " + CalloutTitle);
            foreach (string line in markdown.Replace("\r\n", "\n").Split('\n'))
            {
                callout.Append('\n').Append(line.Length == 0 ? ">" : "> " + line);
            }
            return callout.ToString();
        }

        private static int StartOfBody(string html)
        {
            Match body = BodyOpenRegex.Match(html);
            return body.Success ? body.Index + body.Length : 0;
        }

        private static int FindEarliest(string input, Regex[] markers, int startAt)
        {
            return markers
                .Select(marker => marker.Match(input, startAt))
                .Where(match => match.Success)
                .Select(match => match.Index)
                .DefaultIfEmpty(-1)
                .Min();
        }
    }
}
