using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Removes user-configured noise (external-sender banners, confidentiality footers, etc.) from
    /// the Markdown email body. Patterns come from <see cref="ObsidianSettings.BodyCleanupPatterns"/>
    /// and run case-insensitive and multiline, so ^ and $ anchor to individual lines. A pattern
    /// that must span lines can opt in with an inline (?s).
    /// </summary>
    internal class BodyCleanerService
    {
        // Bodies are far longer than subjects and sender-controlled, so the ReDoS budget per pattern
        // is larger than SubjectCleanerService's 200ms but still bounded.
        private static readonly TimeSpan PatternTimeout = TimeSpan.FromSeconds(1);
        private static readonly Regex TrailingWhitespaceRegex = new Regex(@"[ \t]+(?=\n|$)");
        private static readonly Regex RepeatedBlankLineRegex = new Regex(@"\n{3,}");

        private readonly ObsidianSettings _settings;

        /// <summary>
        /// Initializes a new instance of the <see cref="BodyCleanerService"/> class.
        /// </summary>
        /// <param name="settings">The settings containing body cleanup patterns.</param>
        public BodyCleanerService(ObsidianSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// Applies every configured body cleanup pattern in order.
        /// </summary>
        /// <param name="body">The Markdown body to clean.</param>
        /// <param name="removedContent">True when at least one pattern removed text.</param>
        /// <returns>
        /// The cleaned body, or <paramref name="body"/> unchanged when no pattern removed anything.
        /// </returns>
        public string Clean(string body, out bool removedContent)
        {
            removedContent = false;
            if (string.IsNullOrEmpty(body))
            {
                return body ?? string.Empty;
            }

            IReadOnlyList<string> patterns = _settings.BodyCleanupPatterns ?? new List<string>();
            if (patterns.Count == 0)
            {
                return body;
            }

            // Plain-text bodies arrive with CRLF; normalize so $ anchors at line ends as users expect.
            string cleaned = body.Replace("\r\n", "\n").Replace("\r", "\n");

            foreach (string pattern in patterns)
            {
                if (string.IsNullOrWhiteSpace(pattern))
                {
                    continue;
                }

                try
                {
                    string replaced = Regex.Replace(cleaned, pattern, string.Empty, RegexOptions.IgnoreCase | RegexOptions.Multiline, PatternTimeout);
                    if (!string.Equals(replaced, cleaned, StringComparison.Ordinal))
                    {
                        removedContent = true;
                        cleaned = replaced;
                    }
                }
                catch (ArgumentException)
                {
                    // Skip invalid regex patterns silently
                }
                catch (RegexMatchTimeoutException)
                {
                    // Skip a pattern that exceeds the backtracking budget rather than hang.
                }
            }

            if (!removedContent)
            {
                return body;
            }

            // Tidy only the gaps removals leave behind; unlike subjects, body line structure is kept.
            cleaned = TrailingWhitespaceRegex.Replace(cleaned, string.Empty);
            cleaned = RepeatedBlankLineRegex.Replace(cleaned, "\n\n");
            return cleaned.Trim();
        }
    }
}
