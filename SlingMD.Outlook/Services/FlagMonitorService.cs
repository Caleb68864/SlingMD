using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class FlagMonitorService
    {
        private readonly ObsidianSettings _settings;
        private readonly EmailProcessor _emailProcessor;
        private readonly NotificationService _notificationService;
        private readonly Dictionary<string, OlFlagStatus> _lastKnownFlagStatus;

        private MAPIFolder _inboxFolder;
        private Items _inboxItems;
        private volatile bool _shuttingDown;

        public FlagMonitorService(
            ObsidianSettings settings,
            EmailProcessor emailProcessor,
            NotificationService notificationService)
        {
            _settings = settings;
            _emailProcessor = emailProcessor;
            _notificationService = notificationService;
            _lastKnownFlagStatus = new Dictionary<string, OlFlagStatus>(StringComparer.OrdinalIgnoreCase);
        }

        public void Start(Application outlookApp)
        {
            try
            {
                _inboxFolder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                _inboxItems = _inboxFolder.Items;
                _inboxItems.ItemChange += OnItemChange;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"FlagMonitorService.Start failed: {ex.Message}");
            }
        }

        public void SignalShutdown()
        {
            _shuttingDown = true;
        }

        public void Stop()
        {
            try
            {
                if (_inboxItems != null)
                {
                    _inboxItems.ItemChange -= OnItemChange;
                    Marshal.ReleaseComObject(_inboxItems);
                    _inboxItems = null;
                }

                if (_inboxFolder != null)
                {
                    Marshal.ReleaseComObject(_inboxFolder);
                    _inboxFolder = null;
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"FlagMonitorService.Stop failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns true only when the flag transitions to olFlagMarked (i.e., user just flagged the item).
        /// Pure logic method — no COM dependency. Safe to call from unit tests.
        /// </summary>
        public static bool HasFlagTransitioned(OlFlagStatus oldFlagStatus, OlFlagStatus newFlagStatus)
        {
            return oldFlagStatus != OlFlagStatus.olFlagMarked &&
                   newFlagStatus == OlFlagStatus.olFlagMarked;
        }

        private async void OnItemChange(object item)
        {
            try
            {
                MailItem mail = item as MailItem;
                if (mail == null)
                {
                    return;
                }

                string entryId = SafeComAction.Execute(
                    () => mail.EntryID,
                    "FlagMonitorService.OnItemChange: EntryID",
                    string.Empty);
                if (string.IsNullOrEmpty(entryId))
                {
                    return;
                }

                OlFlagStatus currentFlagStatus = SafeComAction.Execute(
                    () => mail.FlagStatus,
                    "FlagMonitorService.OnItemChange: FlagStatus",
                    OlFlagStatus.olNoFlag);

                OlFlagStatus previousFlagStatus;
                if (!_lastKnownFlagStatus.TryGetValue(entryId, out previousFlagStatus))
                {
                    previousFlagStatus = OlFlagStatus.olNoFlag;
                }

                _lastKnownFlagStatus[entryId] = currentFlagStatus;

                if (!HasFlagTransitioned(previousFlagStatus, currentFlagStatus))
                {
                    return;
                }

                if (_shuttingDown)
                {
                    return;
                }

                await _emailProcessor.ProcessEmail(mail);

                try
                {
                    if (!string.IsNullOrEmpty(_settings.SentToObsidianCategory))
                    {
                        string existingCategories = mail.Categories ?? string.Empty;
                        if (existingCategories.IndexOf(_settings.SentToObsidianCategory, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            mail.Categories = string.IsNullOrEmpty(existingCategories)
                                ? _settings.SentToObsidianCategory
                                : existingCategories + ", " + _settings.SentToObsidianCategory;
                            mail.Save();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Error($"FlagMonitorService: failed to set category: {ex.Message}");
                }

                string subject = SafeComAction.Execute(() => mail.Subject, "FlagMonitorService.OnItemChange: Subject", "Unknown");
                _notificationService.Notify($"Auto-slung flagged email: {subject}");
            }
            catch (System.Exception ex)
            {
                _notificationService.NotifyError("FlagMonitorService: error processing flagged email.", ex.Message);
            }
        }
    }
}
