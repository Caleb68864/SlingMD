using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class SlingDecisionEngineTests
    {
        private readonly SlingDecisionEngine _engine = new SlingDecisionEngine();

        private static MailItemSnapshot Snapshot(string sender = "x@acme.com", string domain = "acme.com", string categories = "")
        {
            return new MailItemSnapshot { SenderEmail = sender, SenderDomain = domain, Categories = categories };
        }

        [Fact]
        public void Decide_NullSnapshot_ReturnsShouldSlingFalse()
        {
            SlingDecision d = _engine.Decide(null, new List<AutoSlingRule>());
            Assert.False(d.ShouldSling);
            Assert.Null(d.MatchedRule);
        }

        [Fact]
        public void Decide_NullRules_ReturnsShouldSlingFalse()
        {
            SlingDecision d = _engine.Decide(Snapshot(), null);
            Assert.False(d.ShouldSling);
            Assert.Null(d.MatchedRule);
        }

        [Fact]
        public void Decide_EmptyRules_ReturnsShouldSlingFalse()
        {
            SlingDecision d = _engine.Decide(Snapshot(), new List<AutoSlingRule>());
            Assert.False(d.ShouldSling);
            Assert.Null(d.MatchedRule);
        }

        [Fact]
        public void Decide_SenderMatch_ReturnsTrueAndRule()
        {
            AutoSlingRule rule = new AutoSlingRule { Type = "Sender", Pattern = "x@acme.com", Enabled = true };
            SlingDecision d = _engine.Decide(Snapshot("x@acme.com", "acme.com"), new List<AutoSlingRule> { rule });
            Assert.True(d.ShouldSling);
            Assert.Same(rule, d.MatchedRule);
        }

        [Fact]
        public void Decide_DomainMatch_ReturnsTrueAndRule()
        {
            AutoSlingRule rule = new AutoSlingRule { Type = "Domain", Pattern = "acme.com", Enabled = true };
            SlingDecision d = _engine.Decide(Snapshot("anyone@acme.com", "acme.com"), new List<AutoSlingRule> { rule });
            Assert.True(d.ShouldSling);
            Assert.Same(rule, d.MatchedRule);
        }

        [Fact]
        public void Decide_CategoryMatch_ReturnsTrueAndRule()
        {
            AutoSlingRule rule = new AutoSlingRule { Type = "Category", Pattern = "ToObsidian", Enabled = true };
            SlingDecision d = _engine.Decide(Snapshot(categories: "Work, ToObsidian, Important"), new List<AutoSlingRule> { rule });
            Assert.True(d.ShouldSling);
            Assert.Same(rule, d.MatchedRule);
        }

        [Fact]
        public void Decide_DisabledRule_DoesNotMatch()
        {
            AutoSlingRule rule = new AutoSlingRule { Type = "Sender", Pattern = "x@acme.com", Enabled = false };
            SlingDecision d = _engine.Decide(Snapshot("x@acme.com"), new List<AutoSlingRule> { rule });
            Assert.False(d.ShouldSling);
        }

        [Fact]
        public void Decide_NoMatchingRule_ReturnsFalse()
        {
            AutoSlingRule rule = new AutoSlingRule { Type = "Domain", Pattern = "different.com", Enabled = true };
            SlingDecision d = _engine.Decide(Snapshot("x@acme.com", "acme.com"), new List<AutoSlingRule> { rule });
            Assert.False(d.ShouldSling);
            Assert.Null(d.MatchedRule);
        }

        [Fact]
        public void Decide_FirstMatchWins()
        {
            AutoSlingRule first = new AutoSlingRule { Type = "Domain", Pattern = "acme.com", Enabled = true };
            AutoSlingRule second = new AutoSlingRule { Type = "Sender", Pattern = "x@acme.com", Enabled = true };
            SlingDecision d = _engine.Decide(Snapshot("x@acme.com", "acme.com"), new List<AutoSlingRule> { first, second });
            Assert.Same(first, d.MatchedRule);
        }

        [Fact]
        public void Decide_CaseInsensitive_OnSender()
        {
            AutoSlingRule rule = new AutoSlingRule { Type = "Sender", Pattern = "x@ACME.com", Enabled = true };
            SlingDecision d = _engine.Decide(Snapshot("X@acme.COM", "acme.com"), new List<AutoSlingRule> { rule });
            Assert.True(d.ShouldSling);
        }
    }
}
