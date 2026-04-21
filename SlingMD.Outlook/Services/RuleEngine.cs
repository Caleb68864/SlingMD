using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class RuleEngine
    {
        /// <summary>
        /// Evaluates a list of AutoSlingRules against the given email metadata.
        /// Returns true if any enabled rule matches. First match wins. Case-insensitive.
        /// </summary>
        public virtual bool ShouldAutoSling(
            string senderEmail,
            string senderDomain,
            string categories,
            List<AutoSlingRule> rules)
        {
            if (rules == null || rules.Count == 0)
            {
                return false;
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

                switch (rule.Type)
                {
                    case "Sender":
                        if (string.Equals(senderEmail, rule.Pattern, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                        break;

                    case "Domain":
                        if (string.Equals(senderDomain, rule.Pattern, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                        break;

                    case "Category":
                        if (!string.IsNullOrEmpty(categories) &&
                            categories.IndexOf(rule.Pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return true;
                        }
                        break;
                }
            }

            return false;
        }
    }
}
