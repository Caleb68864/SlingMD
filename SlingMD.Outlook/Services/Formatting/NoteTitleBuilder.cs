using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Builds note titles from format strings with token substitution and length truncation.
    /// Shared by email and appointment title formatting.
    /// </summary>
    public class NoteTitleBuilder
    {
        /// <summary>
        /// The ellipsis character used when truncating titles.
        /// </summary>
        public const string Ellipsis = "...";

        private static readonly Regex TrailingSeparatorRegex = new Regex(@"[-\s]+$", RegexOptions.Compiled);

        /// <summary>
        /// Builds a title from a format string, substituting tokens and enforcing a maximum length.
        /// </summary>
        /// <param name="format">The format string with <c>{TokenName}</c> placeholders.</param>
        /// <param name="tokens">Dictionary of token names to values. Missing tokens render as empty strings.</param>
        /// <param name="maxLen">Maximum length of the returned string. If truncation is needed, an ellipsis is appended and the final length will be &lt;= maxLen.</param>
        /// <returns>The formatted title, truncated with ellipsis if necessary.</returns>
        public string Build(string format, Dictionary<string, string> tokens, int maxLen)
        {
            if (string.IsNullOrEmpty(format))
            {
                return string.Empty;
            }

            if (maxLen <= 0)
            {
                return string.Empty;
            }

            string result = format;

            // Substitute all tokens
            if (tokens != null)
            {
                foreach (KeyValuePair<string, string> kvp in tokens)
                {
                    string placeholder = "{" + kvp.Key + "}";
                    string value = kvp.Value ?? string.Empty;
                    result = result.Replace(placeholder, value);
                }
            }

            // Remove any remaining unmatched tokens (render as empty)
            result = RemoveUnmatchedTokens(result);

            // Trim whitespace
            result = result.Trim();

            // Truncate if needed
            if (result.Length > maxLen)
            {
                int truncateAt = maxLen - Ellipsis.Length;
                if (truncateAt <= 0)
                {
                    // maxLen is very small, just return truncated ellipsis
                    return Ellipsis.Substring(0, Math.Min(Ellipsis.Length, maxLen));
                }
                result = result.Substring(0, truncateAt) + Ellipsis;
            }

            return result;
        }

        /// <summary>
        /// Same as <see cref="Build"/> but additionally strips any trailing run of dashes and
        /// whitespace. Useful when a token like <c>{Date}</c> renders empty and would leave a
        /// dangling separator (e.g. "Subject - ").
        /// </summary>
        public string BuildTrimmed(string format, Dictionary<string, string> tokens, int maxLen)
        {
            string result = Build(format, tokens, maxLen);
            if (string.IsNullOrEmpty(result))
            {
                return result;
            }
            return TrailingSeparatorRegex.Replace(result, string.Empty).Trim();
        }

        /// <summary>
        /// Removes unmatched <c>{Token}</c> placeholders from the input string.
        /// </summary>
        private string RemoveUnmatchedTokens(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            // Simple approach: find all {word} patterns and remove them
            int startIndex = 0;
            while (startIndex < input.Length)
            {
                int openBrace = input.IndexOf('{', startIndex);
                if (openBrace < 0)
                {
                    break;
                }

                int closeBrace = input.IndexOf('}', openBrace + 1);
                if (closeBrace < 0)
                {
                    break;
                }

                // Check if content between braces looks like a token (alphanumeric only)
                string tokenContent = input.Substring(openBrace + 1, closeBrace - openBrace - 1);
                if (IsValidTokenName(tokenContent))
                {
                    // Remove this token placeholder
                    input = input.Substring(0, openBrace) + input.Substring(closeBrace + 1);
                    // Don't advance startIndex since we removed content
                }
                else
                {
                    // Not a token, skip past this brace
                    startIndex = openBrace + 1;
                }
            }

            return input;
        }

        /// <summary>
        /// Checks if a string is a valid token name (alphanumeric characters only).
        /// </summary>
        private bool IsValidTokenName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            foreach (char c in name)
            {
                if (!char.IsLetterOrDigit(c))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
