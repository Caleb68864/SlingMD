using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class RuleEngine
    {
        /// <summary>
        /// Returns the first enabled rule that matches the given email metadata, or null if none.
        /// First match wins. Case-insensitive.
        /// </summary>
        public virtual AutoSlingRule Match(
            string senderEmail,
            string senderDomain,
            string categories,
            IReadOnlyList<AutoSlingRule> rules)
        {
            if (rules == null || rules.Count == 0)
            {
                return null;
            }

            foreach (AutoSlingRule rule in rules)
            {
                if (!rule.Enabled)
                {
                    continue;
                }

                if (string.IsNullOrEmpty(rule.Pattern))
                {
                    continue;
                }

                // Normalize the rule type so a hand-edited "sender"/"CATEGORY" still matches (the
                // XML contract promises case-insensitivity; the switch was previously exact-case and
                // silently disabled any off-case rule).
                switch ((rule.Type ?? string.Empty).Trim().ToLowerInvariant())
                {
                    case "sender":
                        if (string.Equals(senderEmail, rule.Pattern, StringComparison.OrdinalIgnoreCase))
                        {
                            return rule;
                        }
                        break;

                    case "domain":
                        if (string.Equals(senderDomain, rule.Pattern, StringComparison.OrdinalIgnoreCase))
                        {
                            return rule;
                        }
                        break;

                    case "category":
                        // Outlook categories are a comma-separated list. Compare each category for
                        // equality rather than a substring IndexOf, so pattern "Red" no longer matches
                        // category "Redirected" (or spans two comma-joined names).
                        if (CategoryMatches(categories, rule.Pattern))
                        {
                            return rule;
                        }
                        break;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns true when <paramref name="pattern"/> equals one of the comma-separated categories
        /// in <paramref name="categories"/> (case-insensitive, whitespace-trimmed).
        /// </summary>
        private static bool CategoryMatches(string categories, string pattern)
        {
            if (string.IsNullOrEmpty(categories) || string.IsNullOrEmpty(pattern))
            {
                return false;
            }

            foreach (string category in categories.Split(','))
            {
                if (string.Equals(category.Trim(), pattern.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Evaluates a list of AutoSlingRules against the given email metadata.
        /// Returns true if any enabled rule matches. Thin wrapper over <see cref="Match"/>.
        /// </summary>
        public virtual bool ShouldAutoSling(
            string senderEmail,
            string senderDomain,
            string categories,
            List<AutoSlingRule> rules)
        {
            return Match(senderEmail, senderDomain, categories, rules) != null;
        }
    }
}
