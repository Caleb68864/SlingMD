using System.Collections.Generic;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Evaluates whether a mail item should be auto-slung based on rules.
    /// Pure helper with no Outlook Interop dependencies.
    /// </summary>
    public class SlingDecisionEngine
    {
        private readonly RuleEngine _ruleEngine;

        /// <summary>
        /// Initializes a new instance with a custom RuleEngine (for testing).
        /// </summary>
        /// <param name="ruleEngine">The rule engine to use for matching.</param>
        public SlingDecisionEngine(RuleEngine ruleEngine)
        {
            _ruleEngine = ruleEngine ?? new RuleEngine();
        }

        /// <summary>
        /// Initializes a new instance with the default RuleEngine.
        /// </summary>
        public SlingDecisionEngine() : this(new RuleEngine())
        {
        }

        /// <summary>
        /// Decides whether a mail item should be auto-slung based on the provided rules.
        /// </summary>
        /// <param name="snapshot">A snapshot of the mail item properties.</param>
        /// <param name="rules">The list of auto-sling rules to evaluate.</param>
        /// <param name="watched">Optional watched folder configuration (reserved for future use).</param>
        /// <returns>A SlingDecision indicating whether to sling and which rule matched.</returns>
        public SlingDecision Decide(MailItemSnapshot snapshot, IReadOnlyList<AutoSlingRule> rules, WatchedFolder watched = null)
        {
            if (snapshot == null || rules == null || rules.Count == 0)
            {
                return new SlingDecision
                {
                    ShouldSling = false,
                    MatchedRule = null
                };
            }

            // Convert IReadOnlyList to List for RuleEngine.Match compatibility
            List<AutoSlingRule> ruleList = new List<AutoSlingRule>(rules);

            AutoSlingRule matchedRule = _ruleEngine.Match(
                snapshot.SenderEmail,
                snapshot.SenderDomain,
                snapshot.Categories,
                ruleList);

            return new SlingDecision
            {
                ShouldSling = matchedRule != null,
                MatchedRule = matchedRule
            };
        }
    }
}
