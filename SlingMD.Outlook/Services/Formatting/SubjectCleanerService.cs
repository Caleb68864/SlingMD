using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Provides centralized subject line cleanup functionality for email processing.
    /// Removes Re:/Fwd: prefixes, [EXTERNAL] tags, and other noise patterns while
    /// preserving words like "pre-release" that contain "re-" as part of the word.
    /// </summary>
    internal class SubjectCleanerService
    {
        private readonly ObsidianSettings _settings;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubjectCleanerService"/> class.
        /// </summary>
        /// <param name="settings">The settings containing subject cleanup patterns.</param>
        public SubjectCleanerService(ObsidianSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// Performs a full cleanup pass on the subject line, applying all configured
        /// cleanup patterns from settings.
        /// </summary>
        /// <param name="subject">The subject line to clean.</param>
        /// <returns>The cleaned subject line.</returns>
        public string Clean(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            string cleaned = subject;
            IReadOnlyList<string> patterns = _settings.SubjectCleanupPatterns ?? new List<string>();

            foreach (string pattern in patterns)
            {
                if (string.IsNullOrWhiteSpace(pattern))
                {
                    continue;
                }

                try
                {
                    cleaned = Regex.Replace(cleaned, pattern, string.Empty, RegexOptions.IgnoreCase);
                }
                catch (ArgumentException)
                {
                    // Skip invalid regex patterns silently
                }
            }

            // Normalize whitespace
            cleaned = Regex.Replace(cleaned, @"\s+", " ");
            return cleaned.Trim();
        }

        /// <summary>
        /// Normalizes the subject for thread grouping by stripping Re:/Fwd: prefixes
        /// and [EXTERNAL] tags, but leaving other cleanup patterns unapplied.
        /// This is used for generating consistent thread IDs.
        /// </summary>
        /// <param name="subject">The subject line to normalize.</param>
        /// <returns>The normalized subject suitable for thread grouping.</returns>
        public string NormalizeForGrouping(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            string normalized = subject;

            // Remove leading Re:/Fwd: prefixes with word boundary to avoid matching inside words
            // This matches the legacy behavior for thread ID generation
            normalized = Regex.Replace(normalized, @"^(?:(?:Re|Fwd|FW|RE|FWD)[- :]|\[EXTERNAL\]|\s)+", string.Empty, RegexOptions.IgnoreCase);

            // Also remove any "Re:" that might appear after [EXTERNAL]
            normalized = Regex.Replace(normalized, @"^Re:\s+", string.Empty, RegexOptions.IgnoreCase);

            return normalized.Trim();
        }
    }
}
