using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class AutoSlingService
    {
        private readonly ObsidianSettings _settings;
        private readonly EmailProcessor _emailProcessor;
        private readonly NotificationService _notificationService;
        private readonly RuleEngine _ruleEngine;
        private readonly BoundedHashSet _processedEntryIds = new BoundedHashSet();

        private Application _outlookApp;
        private ApplicationEvents_11_NewMailExEventHandler _newMailExHandler;
        private volatile bool _shuttingDown;

        public AutoSlingService(
            ObsidianSettings settings,
            EmailProcessor emailProcessor,
            NotificationService notificationService)
        {
            _settings = settings;
            _emailProcessor = emailProcessor;
            _notificationService = notificationService;
            _ruleEngine = new RuleEngine();
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
            try
            {
                if (_processedEntryIds.Contains(entryId))
                {
                    return;
                }

                MailItem mail = null;
                try
                {
                    mail = _outlookApp.Session.GetItemFromID(entryId) as MailItem;
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
                    () => _outlookApp.Session.CurrentUser.Address,
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
                string senderDomain = ExtractDomain(senderEmail);
                string categories = SafeComAction.Execute(
                    () => mail.Categories,
                    "AutoSlingService.ProcessSingleEmail: Categories",
                    string.Empty);

                List<AutoSlingRule> rules = _settings.AutoSlingRules ?? new List<AutoSlingRule>();
                if (!_ruleEngine.ShouldAutoSling(senderEmail, senderDomain, categories, rules))
                {
                    return;
                }

                _processedEntryIds.Add(entryId);

                await _emailProcessor.ProcessEmail(mail);

                string subject = SafeComAction.Execute(() => mail.Subject, "AutoSlingService.ProcessSingleEmail: Subject", "Unknown");
                _notificationService.Notify($"Auto-slung email: {subject}");
            }
            catch (System.Exception ex)
            {
                _notificationService.NotifyError("AutoSlingService: error processing email.", ex.Message);
            }
        }

        private static string ExtractDomain(string email)
        {
            if (string.IsNullOrEmpty(email))
            {
                return string.Empty;
            }

            int atIndex = email.IndexOf('@');
            if (atIndex < 0 || atIndex == email.Length - 1)
            {
                return string.Empty;
            }

            return email.Substring(atIndex + 1);
        }
    }
}
