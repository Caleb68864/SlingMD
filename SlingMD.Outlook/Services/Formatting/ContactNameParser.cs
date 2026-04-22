using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Parses display names into structured <see cref="ContactName"/> parts.
    /// Pure helper with no Outlook Interop dependencies.
    /// </summary>
    public class ContactNameParser
    {
        /// <summary>
        /// Common name suffixes that should be parsed separately.
        /// </summary>
        private static readonly HashSet<string> KnownSuffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Jr.", "Jr", "Sr.", "Sr", "II", "III", "IV", "V",
            "Esq.", "Esq", "PhD", "Ph.D.", "MD", "M.D.", "DDS", "D.D.S."
        };

        /// <summary>
        /// Parses a display name into structured <see cref="ContactName"/> parts.
        /// </summary>
        /// <param name="displayName">The display name to parse.</param>
        /// <param name="email">Optional email address to use as fallback or for the Email property.</param>
        /// <returns>A <see cref="ContactName"/> with all fields populated (empty strings where data is missing).</returns>
        public ContactName Parse(string displayName, string email = null)
        {
            ContactName result = new ContactName
            {
                Email = email ?? string.Empty
            };

            // If displayName is empty/whitespace, try to extract from email
            if (string.IsNullOrWhiteSpace(displayName))
            {
                if (!string.IsNullOrWhiteSpace(email))
                {
                    string localPart = ExtractLocalPart(email);
                    result.FirstName = localPart;
                    result.FullName = localPart;
                    result.DisplayName = localPart;
                    result.ShortName = localPart;
                }
                return result;
            }

            string trimmedName = displayName.Trim();
            result.DisplayName = trimmedName;

            // Check for "LastName, FirstName" format
            if (trimmedName.Contains(","))
            {
                ParseCommaFormat(trimmedName, result);
            }
            else
            {
                ParseSpaceFormat(trimmedName, result);
            }

            // Ensure FullName is set (FirstName + LastName)
            if (string.IsNullOrEmpty(result.FullName))
            {
                result.FullName = BuildFullName(result);
            }

            // ShortName defaults to FirstName or FullName if FirstName is empty
            if (string.IsNullOrEmpty(result.ShortName))
            {
                result.ShortName = !string.IsNullOrEmpty(result.FirstName) ? result.FirstName : result.FullName;
            }

            return result;
        }

        /// <summary>
        /// Extracts the local part (before @) from an email address.
        /// </summary>
        private string ExtractLocalPart(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return string.Empty;
            }

            int atIndex = email.IndexOf('@');
            if (atIndex > 0)
            {
                return email.Substring(0, atIndex);
            }

            return email;
        }

        /// <summary>
        /// Parses names in "LastName, FirstName MiddleName Suffix" format.
        /// </summary>
        private void ParseCommaFormat(string name, ContactName result)
        {
            int commaIndex = name.IndexOf(',');
            string lastName = name.Substring(0, commaIndex).Trim();
            string remaining = name.Substring(commaIndex + 1).Trim();

            result.LastName = lastName;

            if (!string.IsNullOrEmpty(remaining))
            {
                string[] parts = remaining.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length >= 1)
                {
                    result.FirstName = parts[0];
                }

                if (parts.Length >= 2)
                {
                    // Check if the last part is a suffix
                    string lastPart = parts[parts.Length - 1];
                    if (KnownSuffixes.Contains(lastPart))
                    {
                        result.Suffix = lastPart;

                        // Middle name is everything between first and suffix
                        if (parts.Length > 2)
                        {
                            result.MiddleName = string.Join(" ", parts, 1, parts.Length - 2);
                        }
                    }
                    else
                    {
                        // No suffix; everything else is middle name
                        result.MiddleName = string.Join(" ", parts, 1, parts.Length - 1);
                    }
                }
            }

            // Build FullName as "FirstName LastName" (not the original comma format)
            result.FullName = BuildFullName(result);
        }

        /// <summary>
        /// Parses names in "FirstName MiddleName LastName Suffix" format.
        /// </summary>
        private void ParseSpaceFormat(string name, ContactName result)
        {
            string[] parts = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
            {
                return;
            }

            if (parts.Length == 1)
            {
                // Single name like "Madonna"
                result.FirstName = parts[0];
                result.FullName = parts[0];
                return;
            }

            // Check if last part is a suffix
            string lastPart = parts[parts.Length - 1];
            bool hasSuffix = KnownSuffixes.Contains(lastPart);
            int effectiveLength = hasSuffix ? parts.Length - 1 : parts.Length;

            if (hasSuffix)
            {
                result.Suffix = lastPart;
            }

            if (effectiveLength == 1)
            {
                // Only first name + suffix
                result.FirstName = parts[0];
            }
            else if (effectiveLength == 2)
            {
                // First + Last (+ optional suffix)
                result.FirstName = parts[0];
                result.LastName = parts[1];
            }
            else
            {
                // First + Middle(s) + Last (+ optional suffix)
                result.FirstName = parts[0];
                result.LastName = parts[effectiveLength - 1];

                // Middle name(s) are everything between first and last
                List<string> middleParts = new List<string>();
                for (int i = 1; i < effectiveLength - 1; i++)
                {
                    middleParts.Add(parts[i]);
                }
                result.MiddleName = string.Join(" ", middleParts);
            }

            result.FullName = BuildFullName(result);
        }

        /// <summary>
        /// Builds the FullName from FirstName and LastName.
        /// </summary>
        private string BuildFullName(ContactName result)
        {
            if (!string.IsNullOrEmpty(result.FirstName) && !string.IsNullOrEmpty(result.LastName))
            {
                return result.FirstName + " " + result.LastName;
            }
            else if (!string.IsNullOrEmpty(result.FirstName))
            {
                return result.FirstName;
            }
            else if (!string.IsNullOrEmpty(result.LastName))
            {
                return result.LastName;
            }

            return string.Empty;
        }
    }
}
