using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Formats <see cref="ContactName"/> instances using user-defined format strings.
    /// Supports tokens: {FullName}, {FirstName}, {LastName}, {MiddleName}, {Suffix},
    /// {DisplayName}, {ShortName}, {Email}, {FirstInitial}, {LastInitial}.
    /// </summary>
    internal class ContactLinkFormatter
    {
        /// <summary>
        /// Known tokens that can be used in format strings.
        /// </summary>
        public static readonly HashSet<string> KnownTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "FullName",
            "FirstName",
            "LastName",
            "MiddleName",
            "Suffix",
            "DisplayName",
            "ShortName",
            "Email",
            "FirstInitial",
            "LastInitial"
        };

        /// <summary>
        /// Pattern to match tokens in format strings like {TokenName}.
        /// </summary>
        private static readonly Regex TokenPattern = new Regex(@"\{([^}]+)\}", RegexOptions.Compiled);

        /// <summary>
        /// Formats a contact name using the specified format string.
        /// </summary>
        /// <param name="name">The contact name to format.</param>
        /// <param name="format">The format string with tokens like {FullName}, {FirstName}, etc.</param>
        /// <returns>The formatted string. If all referenced tokens are empty, falls back to DisplayName or FullName.</returns>
        public string Format(ContactName name, string format)
        {
            if (name == null)
            {
                return string.Empty;
            }

            if (string.IsNullOrEmpty(format))
            {
                return string.Empty;
            }

            // Build token values dictionary
            Dictionary<string, string> tokenValues = BuildTokenValues(name);

            // Track if any token produced non-empty output
            bool anyTokenRendered = false;

            // Replace all tokens
            string result = TokenPattern.Replace(format, match =>
            {
                string tokenName = match.Groups[1].Value;

                // Look up token case-insensitively
                foreach (KeyValuePair<string, string> kvp in tokenValues)
                {
                    if (string.Equals(kvp.Key, tokenName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(kvp.Value))
                        {
                            anyTokenRendered = true;
                        }
                        return kvp.Value;
                    }
                }

                // Unknown token renders as empty (not as the literal {Token})
                return string.Empty;
            });

            // If no tokens rendered any content, fall back to DisplayName or FullName
            if (!anyTokenRendered)
            {
                string fallback = !string.IsNullOrEmpty(name.DisplayName) ? name.DisplayName : name.FullName;
                if (!string.IsNullOrEmpty(fallback))
                {
                    return fallback;
                }
            }

            return result;
        }

        /// <summary>
        /// Formats a list of contact names and joins them with a separator.
        /// </summary>
        /// <param name="names">The contact names to format.</param>
        /// <param name="format">The format string with tokens.</param>
        /// <param name="separator">The separator to use between formatted names.</param>
        /// <returns>A string of formatted names joined by the separator, skipping null/empty entries.</returns>
        public string FormatList(IEnumerable<ContactName> names, string format, string separator)
        {
            if (names == null)
            {
                return string.Empty;
            }

            List<string> formattedNames = new List<string>();

            foreach (ContactName name in names)
            {
                if (name == null)
                {
                    continue;
                }

                string formatted = Format(name, format);
                if (!string.IsNullOrEmpty(formatted))
                {
                    formattedNames.Add(formatted);
                }
            }

            return string.Join(separator ?? string.Empty, formattedNames);
        }

        /// <summary>
        /// Checks if a format string contains any unknown tokens.
        /// </summary>
        /// <param name="format">The format string to check.</param>
        /// <returns>A list of unknown token names found in the format string.</returns>
        public List<string> GetUnknownTokens(string format)
        {
            List<string> unknownTokens = new List<string>();

            if (string.IsNullOrEmpty(format))
            {
                return unknownTokens;
            }

            MatchCollection matches = TokenPattern.Matches(format);
            foreach (Match match in matches)
            {
                string tokenName = match.Groups[1].Value;
                bool isKnown = false;

                foreach (string known in KnownTokens)
                {
                    if (string.Equals(known, tokenName, StringComparison.OrdinalIgnoreCase))
                    {
                        isKnown = true;
                        break;
                    }
                }

                if (!isKnown && !unknownTokens.Contains(tokenName))
                {
                    unknownTokens.Add(tokenName);
                }
            }

            return unknownTokens;
        }

        /// <summary>
        /// Builds a dictionary of token values from a ContactName.
        /// </summary>
        private Dictionary<string, string> BuildTokenValues(ContactName name)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            values["FullName"] = name.FullName ?? string.Empty;
            values["FirstName"] = name.FirstName ?? string.Empty;
            values["LastName"] = name.LastName ?? string.Empty;
            values["MiddleName"] = name.MiddleName ?? string.Empty;
            values["Suffix"] = name.Suffix ?? string.Empty;
            values["DisplayName"] = name.DisplayName ?? string.Empty;
            values["ShortName"] = name.ShortName ?? string.Empty;
            values["Email"] = name.Email ?? string.Empty;

            // Compute initials
            string firstName = name.FirstName ?? string.Empty;
            string lastName = name.LastName ?? string.Empty;

            values["FirstInitial"] = firstName.Length > 0 ? firstName.Substring(0, 1).ToUpperInvariant() : string.Empty;
            values["LastInitial"] = lastName.Length > 0 ? lastName.Substring(0, 1).ToUpperInvariant() : string.Empty;

            return values;
        }
    }
}
