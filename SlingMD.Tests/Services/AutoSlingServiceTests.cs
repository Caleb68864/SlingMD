using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class AutoSlingServiceTests
    {
        private static MailItemSnapshot Snapshot(string sender = "someone@acme.com", string domain = "acme.com", string categories = "")
        {
            return new MailItemSnapshot
            {
                SenderEmail = sender,
                SenderDomain = domain,
                Categories = categories
            };
        }

        private static List<AutoSlingRule> DomainRule(string domain = "acme.com")
        {
            return new List<AutoSlingRule>
            {
                new AutoSlingRule { Type = "Domain", Pattern = domain, Enabled = true }
            };
        }

        [Fact]
        public void EvaluateEligibility_AutoSlingDisabled_ReturnsDisabled()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: false,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(),
                rules: DomainRule());

            Assert.Equal(AutoSlingEligibility.Disabled, result.Outcome);
            Assert.Null(result.MatchedRule);
        }

        [Fact]
        public void EvaluateEligibility_AlreadyProcessed_ReturnsAlreadyProcessed()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: true,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(),
                rules: DomainRule());

            Assert.Equal(AutoSlingEligibility.AlreadyProcessed, result.Outcome);
        }

        [Fact]
        public void EvaluateEligibility_ShuttingDown_ReturnsShuttingDown()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: true,
                isAlreadyProcessed: false,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(),
                rules: DomainRule());

            Assert.Equal(AutoSlingEligibility.ShuttingDown, result.Outcome);
        }

        [Fact]
        public void EvaluateEligibility_SenderMatchesCurrentUser_ReturnsSelfSend()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(sender: "me@acme.com"),
                rules: DomainRule());

            Assert.Equal(AutoSlingEligibility.SelfSend, result.Outcome);
        }

        [Fact]
        public void EvaluateEligibility_SenderMatchesCurrentUser_CaseInsensitive_ReturnsSelfSend()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: "Me@Acme.COM",
                snapshot: Snapshot(sender: "me@acme.com"),
                rules: DomainRule());

            Assert.Equal(AutoSlingEligibility.SelfSend, result.Outcome);
        }

        [Fact]
        public void EvaluateEligibility_EmptyCurrentUserAddress_DoesNotTriggerSelfSend()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: string.Empty,
                snapshot: Snapshot(sender: "someone@acme.com"),
                rules: DomainRule());

            // Self-send guard must not fire when we can't identify the current user,
            // otherwise a missing CurrentUser.Address would silently disable auto-sling.
            Assert.Equal(AutoSlingEligibility.Sling, result.Outcome);
        }

        [Fact]
        public void EvaluateEligibility_NoMatchingRule_ReturnsNoMatch()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(sender: "stranger@other.com", domain: "other.com"),
                rules: DomainRule("acme.com"));

            Assert.Equal(AutoSlingEligibility.NoMatch, result.Outcome);
            Assert.Null(result.MatchedRule);
        }

        [Fact]
        public void EvaluateEligibility_MatchingRule_ReturnsSlingWithMatchedRule()
        {
            List<AutoSlingRule> rules = DomainRule("acme.com");

            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(sender: "someone@acme.com", domain: "acme.com"),
                rules: rules);

            Assert.Equal(AutoSlingEligibility.Sling, result.Outcome);
            Assert.Same(rules[0], result.MatchedRule);
        }

        [Fact]
        public void EvaluateEligibility_EmptyRules_ReturnsNoMatch()
        {
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: true,
                isShuttingDown: false,
                isAlreadyProcessed: false,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(),
                rules: new List<AutoSlingRule>());

            Assert.Equal(AutoSlingEligibility.NoMatch, result.Outcome);
        }

        [Fact]
        public void EvaluateEligibility_GuardsOrderedCorrectly_DisabledBeatsAlreadyProcessed()
        {
            // If both disabled and already-processed are true, Disabled wins (cheaper check).
            AutoSlingEligibilityResult result = AutoSlingService.EvaluateEligibility(
                enableAutoSling: false,
                isShuttingDown: true,
                isAlreadyProcessed: true,
                currentUserAddress: "me@acme.com",
                snapshot: Snapshot(sender: "me@acme.com"),
                rules: DomainRule());

            Assert.Equal(AutoSlingEligibility.Disabled, result.Outcome);
        }
    }
}
