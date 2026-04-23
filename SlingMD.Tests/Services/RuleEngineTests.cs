using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class RuleEngineTests
    {
        private readonly RuleEngine _ruleEngine = new RuleEngine();

        [Fact]
        public void ShouldAutoSling_DomainRule_MatchingDomain_ReturnsTrue()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Domain", Pattern = "acme.com", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.True(result);
        }

        [Fact]
        public void ShouldAutoSling_DomainRule_NonMatchingDomain_ReturnsFalse()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Domain", Pattern = "other.com", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.False(result);
        }

        [Fact]
        public void ShouldAutoSling_SenderRule_MatchingSender_ReturnsTrue()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Sender", Pattern = "john@acme.com", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.True(result);
        }

        [Fact]
        public void ShouldAutoSling_SenderRule_NonMatchingSender_ReturnsFalse()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Sender", Pattern = "jane@acme.com", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.False(result);
        }

        [Fact]
        public void ShouldAutoSling_CategoryRule_MatchingCategory_ReturnsTrue()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Category", Pattern = "Important", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "Important", rules);

            Assert.True(result);
        }

        [Fact]
        public void ShouldAutoSling_CategoryRule_NonMatchingCategory_ReturnsFalse()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Category", Pattern = "Important", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "Normal", rules);

            Assert.False(result);
        }

        [Fact]
        public void ShouldAutoSling_DisabledRule_ReturnsFalse()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Sender", Pattern = "john@acme.com", Enabled = false }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.False(result);
        }

        [Fact]
        public void ShouldAutoSling_EmptyRulesList_ReturnsFalse()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>();

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.False(result);
        }

        [Fact]
        public void ShouldAutoSling_MultipleRules_FirstMatchWins()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Sender", Pattern = "john@acme.com", Enabled = true },
                new AutoSlingRule { Type = "Domain", Pattern = "acme.com", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            // Both match, but first match wins — result is still true
            Assert.True(result);
        }

        [Fact]
        public void ShouldAutoSling_CaseInsensitive_ReturnsTrue()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Sender", Pattern = "JOHN@ACME.COM", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.True(result);
        }

        [Fact]
        public void ShouldAutoSling_NullRules_ReturnsFalse()
        {
            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", null);

            Assert.False(result);
        }

        [Fact]
        public void ShouldAutoSling_DisabledRuleSkipped_EnabledRuleMatches()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Sender", Pattern = "john@acme.com", Enabled = false },
                new AutoSlingRule { Type = "Domain", Pattern = "acme.com", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            // First rule is disabled (skipped), second matches
            Assert.True(result);
        }

        [Fact]
        public void ShouldAutoSling_DomainRule_CaseInsensitive_ReturnsTrue()
        {
            List<AutoSlingRule> rules = new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Domain", Pattern = "ACME.COM", Enabled = true }
            };

            bool result = _ruleEngine.ShouldAutoSling("john@acme.com", "acme.com", "", rules);

            Assert.True(result);
        }
    }
}
