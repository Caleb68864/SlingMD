using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    internal class ContactNameNormalizer
    {
        private static readonly ContactNameParser _parser = new ContactNameParser();

        private static readonly HashSet<string> _honorifics = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Dr", "Dr.", "Mr", "Mr.", "Mrs", "Mrs.", "Ms", "Ms.", "Miss",
            "Prof", "Prof.", "Rev", "Rev.", "Sir", "Lord", "Lady",
            "Capt", "Capt.", "Sgt", "Sgt.", "Lt", "Lt.", "Col", "Col.", "Gen", "Gen."
        };

        private static readonly HashSet<string> _suffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Jr", "Jr.", "Sr", "Sr.", "II", "III", "IV", "V",
            "Esq", "Esq.", "PhD", "Ph.D.", "MD", "M.D.", "DDS", "D.D.S."
        };

        public string Normalize(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
                return string.Empty;

            string name = displayName.Trim();

            // Strip parenthetical content (e.g., company names)
            name = Regex.Replace(name, @"\s*\([^)]*\)\s*", " ").Trim();

            // Strip leading honorific
            name = StripLeadingHonorific(name);

            // Normalize trailing ", Suffix" → " Suffix" before parsing
            name = NormalizeTrailingSuffix(name);

            // Parse into structured parts
            Models.ContactName parsed = _parser.Parse(name);

            var parts = new List<string>();
            if (!string.IsNullOrEmpty(parsed.FirstName))
                parts.Add(parsed.FirstName);

            if (!string.IsNullOrEmpty(parsed.MiddleName))
            {
                foreach (string segment in parsed.MiddleName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (!IsInitial(segment))
                        parts.Add(segment);
                }
            }

            if (!string.IsNullOrEmpty(parsed.LastName))
                parts.Add(parsed.LastName);

            if (!string.IsNullOrEmpty(parsed.Suffix))
                parts.Add(parsed.Suffix);

            string joined = string.Join(" ", parts).ToLowerInvariant();
            joined = joined.Replace(".", "");
            return Regex.Replace(joined, @"\s+", " ").Trim();
        }

        public (string firstKey, string lastKey) NormalizeFirstLast(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
                return (string.Empty, string.Empty);

            string name = displayName.Trim();
            name = Regex.Replace(name, @"\s*\([^)]*\)\s*", " ").Trim();
            name = StripLeadingHonorific(name);
            name = NormalizeTrailingSuffix(name);

            Models.ContactName parsed = _parser.Parse(name);
            return (
                (parsed.FirstName ?? string.Empty).ToLowerInvariant(),
                (parsed.LastName ?? string.Empty).ToLowerInvariant()
            );
        }

        private string StripLeadingHonorific(string name)
        {
            string[] tokens = name.Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2 && _honorifics.Contains(tokens[0]))
                return tokens[1];
            return name;
        }

        private string NormalizeTrailingSuffix(string name)
        {
            int commaIdx = name.LastIndexOf(',');
            if (commaIdx < 0)
                return name;

            string afterComma = name.Substring(commaIdx + 1).Trim();
            if (_suffixes.Contains(afterComma))
            {
                string beforeComma = name.Substring(0, commaIdx).Trim();
                return beforeComma + " " + afterComma;
            }

            return name;
        }

        private static bool IsInitial(string token)
        {
            return (token.Length == 1 && char.IsLetter(token[0])) ||
                   (token.Length == 2 && char.IsLetter(token[0]) && token[1] == '.');
        }
    }
}
