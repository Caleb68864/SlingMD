using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Possible outcomes from evaluating whether a new mail item should be auto-slung.
    /// </summary>
    public enum AutoSlingEligibility
    {
        Disabled,
        AlreadyProcessed,
        ShuttingDown,
        SelfSend,
        NoMatch,
        Sling
    }

    /// <summary>
    /// Result of <see cref="AutoSlingService.EvaluateEligibility"/>. MatchedRule is only populated
    /// when Outcome is <see cref="AutoSlingEligibility.Sling"/>.
    /// </summary>
    public class AutoSlingEligibilityResult
    {
        public AutoSlingEligibility Outcome { get; set; }
        public AutoSlingRule MatchedRule { get; set; }
    }

    public class AutoSlingService
    {
        private readonly ObsidianSettings _settings;
        private readonly EmailProcessor _emailProcessor;
        private readonly NotificationService _notificationService;
        private readonly SlingDecisionEngine _slingDecisionEngine;
        private readonly EmailAddressParser _emailAddressParser;
        private readonly BoundedHashSet _processedEntryIds;

        private Application _outlookApp;
        private ApplicationEvents_11_NewMailExEventHandler _newMailExHandler;
        private volatile bool _shuttingDown;

        public AutoSlingService(
            ObsidianSettings settings,
            EmailProcessor emailProcessor,
            NotificationService notificationService,
            BoundedHashSet processedEntryIds = null)
        {
            _settings = settings;
            _emailProcessor = emailProcessor;
            _notificationService = notificationService;
            // A shared, caller-owned set dedupes across all monitors (an email that lands in both the
            // Inbox and a watched folder fires two monitors) and survives settings-save recreation.
            // Falls back to a private set when not supplied (e.g. unit tests).
            _processedEntryIds = processedEntryIds ?? new BoundedHashSet();
            _slingDecisionEngine = new SlingDecisionEngine();
            _emailAddressParser = new EmailAddressParser();
        }

        public void Start(Application outlookApp)
        {
            _outlookApp = outlookApp;
            _newMailExHandler = new ApplicationEvents_11_NewMailExEventHandler(OnNewMailEx);
            _outlookApp.NewMailEx += _newMailExHandler;
        }

        public void Shutdown()
        {
            _shuttingDown = true;
            try
            {
                if (_outlookApp != null && _newMailExHandler != null)
                {
                    _outlookApp.NewMailEx -= _newMailExHandler;
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"AutoSlingService.Shutdown failed: {ex.Message}");
            }
            finally
            {
                _newMailExHandler = null;
                _outlookApp = null;
            }
        }

        private async void OnNewMailEx(string entryIdCollection)
        {
            try
            {
                if (!_settings.EnableAutoSling)
                {
                    return;
                }

                if (string.IsNullOrEmpty(entryIdCollection))
                {
                    return;
                }

                string[] entryIds = entryIdCollection.Split(',');
                foreach (string rawEntryId in entryIds)
                {
                    string entryId = rawEntryId.Trim();
                    if (string.IsNullOrEmpty(entryId))
                    {
                        continue;
                    }

                    await ProcessSingleEmail(entryId);
                }
            }
            catch (System.Exception ex)
            {
                _notificationService.NotifyError("AutoSlingService: error handling new mail.", ex.Message);
            }
        }

        private async System.Threading.Tasks.Task ProcessSingleEmail(string entryId)
        {
            // This runs on the NewMailEx event for every inbound email while Auto-Sling is on, so
            // the MailItem and NameSpace COM objects obtained here MUST be released or they leak
            // continuously and unattended for the whole session.
            NameSpace session = null;
            MailItem mail = null;
            try
            {
                if (_processedEntryIds.Contains(entryId))
                {
                    return;
                }

                session = _outlookApp.Session;
                try
                {
                    mail = session.GetItemFromID(entryId) as MailItem;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Error($"AutoSlingService: could not retrieve item {entryId}: {ex.Message}");
                    return;
                }

                if (mail == null)
                {
                    return;
                }

                if (_shuttingDown)
                {
                    return;
                }

                // Self-send guard: skip emails where the sender is the current user
                string currentUserAddress = SafeComAction.Execute(
                    () => session.CurrentUser.Address,
                    "AutoSlingService.ProcessSingleEmail: self-send guard",
                    string.Empty);
                string senderEmailAddress = SafeComAction.Execute(
                    () => mail.SenderEmailAddress,
                    "AutoSlingService.ProcessSingleEmail: SenderEmailAddress",
                    string.Empty);
                if (!string.IsNullOrEmpty(currentUserAddress) &&
                    string.Equals(senderEmailAddress, currentUserAddress, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                string senderEmail = senderEmailAddress;
                string senderDomain = _emailAddressParser.Domain(senderEmail);
                string categories = SafeComAction.Execute(
                    () => mail.Categories,
                    "AutoSlingService.ProcessSingleEmail: Categories",
                    string.Empty);

                List<AutoSlingRule> rules = _settings.AutoSlingRules ?? new List<AutoSlingRule>();
                MailItemSnapshot snapshot = new MailItemSnapshot
                {
                    SenderEmail = senderEmail ?? string.Empty,
                    SenderDomain = senderDomain ?? string.Empty,
                    Categories = categories ?? string.Empty
                };
                AutoSlingEligibilityResult eligibility = EvaluateEligibility(
                    enableAutoSling: _settings.EnableAutoSling,
                    isShuttingDown: _shuttingDown,
                    isAlreadyProcessed: false, // already checked above; re-check here is redundant
                    currentUserAddress: currentUserAddress,
                    snapshot: snapshot,
                    rules: rules,
                    decisionEngine: _slingDecisionEngine);

                if (eligibility.Outcome != AutoSlingEligibility.Sling)
                {
                    return;
                }

                // Atomically reserve this id before the await. If another monitor (or a prior event
                // for the same id) already reserved it, Add returns false and we skip — the race-safe
                // cross-monitor duplicate guard. Runs synchronously before any await.
                if (!_processedEntryIds.Add(entryId))
                {
                    return;
                }

                await _emailProcessor.ProcessEmail(mail, contactMode: ContactInteractionMode.Automated, bulkMode: true);

                string subject = SafeComAction.Execute(() => mail.Subject, "AutoSlingService.ProcessSingleEmail: Subject", "Unknown");
                string ruleLabel = eligibility.MatchedRule != null ? $" [{eligibility.MatchedRule.Type}:{eligibility.MatchedRule.Pattern}]" : string.Empty;
                _notificationService.Notify($"Auto-slung email: {subject}{ruleLabel}");
            }
            catch (System.Exception ex)
            {
                _notificationService.NotifyError("AutoSlingService: error processing email.", ex.Message);
            }
            finally
            {
                if (mail != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                if (session != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
            }
        }

        /// <summary>
        /// Pure-logic eligibility check. Given the current state (flags, current-user address,
        /// already-processed status) and a mail snapshot + rules, returns the outcome.
        /// No Outlook COM dependencies — designed for unit testing.
        /// </summary>
        public static AutoSlingEligibilityResult EvaluateEligibility(
            bool enableAutoSling,
            bool isShuttingDown,
            bool isAlreadyProcessed,
            string currentUserAddress,
            MailItemSnapshot snapshot,
            IReadOnlyList<AutoSlingRule> rules,
            SlingDecisionEngine decisionEngine = null)
        {
            if (!enableAutoSling)
            {
                return new AutoSlingEligibilityResult { Outcome = AutoSlingEligibility.Disabled };
            }

            if (isAlreadyProcessed)
            {
                return new AutoSlingEligibilityResult { Outcome = AutoSlingEligibility.AlreadyProcessed };
            }

            if (isShuttingDown)
            {
                return new AutoSlingEligibilityResult { Outcome = AutoSlingEligibility.ShuttingDown };
            }

            if (snapshot != null
                && !string.IsNullOrEmpty(currentUserAddress)
                && string.Equals(snapshot.SenderEmail, currentUserAddress, StringComparison.OrdinalIgnoreCase))
            {
                return new AutoSlingEligibilityResult { Outcome = AutoSlingEligibility.SelfSend };
            }

            SlingDecisionEngine engine = decisionEngine ?? new SlingDecisionEngine();
            SlingDecision decision = engine.Decide(snapshot, rules);

            if (decision == null || !decision.ShouldSling)
            {
                return new AutoSlingEligibilityResult { Outcome = AutoSlingEligibility.NoMatch };
            }

            return new AutoSlingEligibilityResult
            {
                Outcome = AutoSlingEligibility.Sling,
                MatchedRule = decision.MatchedRule
            };
        }
    }
}
