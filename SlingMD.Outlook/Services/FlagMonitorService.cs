using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class FlagMonitorService : IDisposable
    {
        private readonly ObsidianSettings _settings;
        private readonly EmailProcessor _emailProcessor;
        private readonly NotificationService _notificationService;
        private readonly Dictionary<string, OlFlagStatus> _lastKnownFlagStatus;

        // Upper bound on the flag-status history. Without a cap the dictionary grows by one entry per
        // distinct inbox item that ever raises ItemChange — an unbounded, session-long leak. When the
        // cap is hit we clear it: at worst a stale item re-evaluates its transition, and the shared
        // _processedEntryIds set still prevents an actual re-sling.
        private const int MaxTrackedFlagItems = 5000;
        // Shared, caller-owned processed-id set: dedupes flag-slings against the other monitors and
        // survives settings-save recreation (which clears _lastKnownFlagStatus and would otherwise
        // let a still-flagged item re-transition and be slung again).
        private readonly BoundedHashSet _processedEntryIds;

        private MAPIFolder _inboxFolder;
        private Items _inboxItems;
        private volatile bool _shuttingDown;

        public FlagMonitorService(
            ObsidianSettings settings,
            EmailProcessor emailProcessor,
            NotificationService notificationService,
            BoundedHashSet processedEntryIds = null)
        {
            _settings = settings;
            _emailProcessor = emailProcessor;
            _notificationService = notificationService;
            _lastKnownFlagStatus = new Dictionary<string, OlFlagStatus>(StringComparer.OrdinalIgnoreCase);
            _processedEntryIds = processedEntryIds ?? new BoundedHashSet();
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
        /// Releases the Outlook COM handles held by this service. Equivalent to calling
        /// <see cref="Stop"/>; implemented so the service can participate in <c>using</c> blocks
        /// and signal ownership of unmanaged Outlook handles to static analyzers.
        /// </summary>
        public void Dispose()
        {
            Stop();
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

                // Bound memory growth. All access here is on the Outlook STA thread (synchronous,
                // before the await), so no locking is needed.
                if (_lastKnownFlagStatus.Count >= MaxTrackedFlagItems && !_lastKnownFlagStatus.ContainsKey(entryId))
                {
                    _lastKnownFlagStatus.Clear();
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

                // Atomically reserve before processing so the same email isn't slung twice — by a
                // repeated ItemChange, by another monitor, or after a settings-save reset.
                if (!_processedEntryIds.Add(entryId))
                {
                    return;
                }

                await _emailProcessor.ProcessEmail(mail, contactMode: ContactInteractionMode.Automated, bulkMode: true);

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
