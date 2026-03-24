using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class FolderMonitorService
    {
        private readonly ObsidianSettings _settings;
        private readonly EmailProcessor _emailProcessor;
        private readonly NotificationService _notificationService;
        private readonly Application _outlookApp;

        // COM objects stored as class-level fields to prevent GC collection
        private Dictionary<string, MAPIFolder> _watchedFolderObjects;
        private Dictionary<string, Items> _watchedFolderItems;

        // Duplicate cache to prevent processing the same email twice
        private readonly BoundedHashSet _processedEntryIds = new BoundedHashSet();

        private volatile bool _shuttingDown;

        public FolderMonitorService(ObsidianSettings settings, EmailProcessor emailProcessor, NotificationService notificationService, Application outlookApp)
        {
            _settings = settings;
            _emailProcessor = emailProcessor;
            _notificationService = notificationService;
            _outlookApp = outlookApp;
            _watchedFolderObjects = new Dictionary<string, MAPIFolder>(StringComparer.OrdinalIgnoreCase);
            _watchedFolderItems = new Dictionary<string, Items>(StringComparer.OrdinalIgnoreCase);
        }

        public void StartWatching(List<WatchedFolder> folders)
        {
            if (folders == null)
            {
                return;
            }

            foreach (WatchedFolder watchedFolder in folders)
            {
                if (!watchedFolder.Enabled || string.IsNullOrWhiteSpace(watchedFolder.FolderPath))
                {
                    continue;
                }

                try
                {
                    MAPIFolder mapiFolder = ResolveFolderPath(watchedFolder.FolderPath);
                    if (mapiFolder == null)
                    {
                        Logger.Instance.Info(string.Format("FolderMonitorService: Could not resolve folder path '{0}' — skipping.", watchedFolder.FolderPath));
                        continue;
                    }

                    Items items = mapiFolder.Items;

                    items.ItemAdd += OnItemAdded;

                    string key = watchedFolder.FolderPath;
                    _watchedFolderObjects[key] = mapiFolder;
                    _watchedFolderItems[key] = items;

                    Logger.Instance.Info(string.Format("FolderMonitorService: Now watching folder '{0}'.", watchedFolder.FolderPath));
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Error(string.Format("FolderMonitorService: Failed to watch folder '{0}': {1}", watchedFolder.FolderPath, ex.Message));
                }
            }
        }

        public void SignalShutdown()
        {
            _shuttingDown = true;
        }

        public void StopWatching()
        {
            foreach (KeyValuePair<string, Items> pair in _watchedFolderItems)
            {
                try
                {
                    pair.Value.ItemAdd -= OnItemAdded;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"FolderMonitorService.StopWatching: could not unsubscribe ItemAdd: {ex.Message}");
                }

                try
                {
                    Marshal.ReleaseComObject(pair.Value);
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"FolderMonitorService.StopWatching: could not release Items COM object: {ex.Message}");
                }
            }

            foreach (KeyValuePair<string, MAPIFolder> pair in _watchedFolderObjects)
            {
                try
                {
                    Marshal.ReleaseComObject(pair.Value);
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"FolderMonitorService.StopWatching: could not release MAPIFolder COM object: {ex.Message}");
                }
            }

            _watchedFolderItems.Clear();
            _watchedFolderObjects.Clear();
            _watchedFolderItems = null;
            _watchedFolderObjects = null;
        }

        private async void OnItemAdded(object item)
        {
            try
            {
                MailItem mail = item as MailItem;
                if (mail == null)
                {
                    return;
                }

                if (_shuttingDown)
                {
                    return;
                }

                string entryId = SafeComAction.Execute(
                    () => mail.EntryID,
                    "FolderMonitorService.OnItemAdded: EntryID",
                    string.Empty);

                if (!string.IsNullOrEmpty(entryId) && _processedEntryIds.Contains(entryId))
                {
                    return;
                }

                // Self-send guard: skip emails sent by the current user to themselves
                if (IsSelfSent(mail))
                {
                    return;
                }

                if (!string.IsNullOrEmpty(entryId))
                {
                    _processedEntryIds.Add(entryId);
                }

                await _emailProcessor.ProcessEmail(mail);

                string subject = SafeComAction.Execute(
                    () => mail.Subject ?? string.Empty,
                    "FolderMonitorService.OnItemAdded: Subject",
                    string.Empty);

                _notificationService.Notify(string.Format("Auto-slung email: {0}", subject));
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error(string.Format("FolderMonitorService.OnItemAdded: {0}", ex.Message));
            }
        }

        private bool IsSelfSent(MailItem mail)
        {
            try
            {
                if (_outlookApp == null)
                {
                    return false;
                }

                Recipient currentUser = _outlookApp.Session.CurrentUser;
                if (currentUser == null)
                {
                    return false;
                }

                AddressEntry currentEntry = currentUser.AddressEntry;
                if (currentEntry == null)
                {
                    return false;
                }

                string currentAddress = string.Empty;
                try
                {
                    currentAddress = currentEntry.Address ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"FolderMonitorService.IsSelfSent: could not read current user address: {ex.Message}");
                }

                string senderAddress = string.Empty;
                try
                {
                    senderAddress = mail.SenderEmailAddress ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"FolderMonitorService.IsSelfSent: could not read sender address: {ex.Message}");
                }

                if (!string.IsNullOrEmpty(currentAddress) && !string.IsNullOrEmpty(senderAddress))
                {
                    return string.Equals(currentAddress, senderAddress, StringComparison.OrdinalIgnoreCase);
                }

                return false;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"FolderMonitorService.IsSelfSent: {ex.Message}");
                return false;
            }
        }

        private MAPIFolder ResolveFolderPath(string folderPath)
        {
            // Expected format: "\\AccountName\FolderName\SubFolder" or "AccountName\FolderName\SubFolder"
            // Navigate the Folders hierarchy from the session root
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return null;
            }

            Folders rootFolders = null;
            try
            {
                string normalized = folderPath.TrimStart('\\');
                string[] parts = normalized.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 0)
                {
                    return null;
                }

                rootFolders = _outlookApp.Session.Folders;
                MAPIFolder current = null;

                // Find the root store (account) matching parts[0]
                foreach (MAPIFolder storeFolder in rootFolders)
                {
                    if (string.Equals(storeFolder.Name, parts[0], StringComparison.OrdinalIgnoreCase))
                    {
                        current = storeFolder;
                    }
                    else
                    {
                        Marshal.ReleaseComObject(storeFolder);
                    }
                }

                if (current == null)
                {
                    return null;
                }

                // Navigate sub-folders
                for (int i = 1; i < parts.Length; i++)
                {
                    Folders subFolders = null;
                    MAPIFolder found = null;
                    try
                    {
                        subFolders = current.Folders;
                        foreach (MAPIFolder subFolder in subFolders)
                        {
                            if (string.Equals(subFolder.Name, parts[i], StringComparison.OrdinalIgnoreCase))
                            {
                                found = subFolder;
                            }
                            else
                            {
                                Marshal.ReleaseComObject(subFolder);
                            }
                        }
                    }
                    finally
                    {
                        if (subFolders != null)
                        {
                            Marshal.ReleaseComObject(subFolders);
                        }
                    }

                    if (found == null)
                    {
                        Marshal.ReleaseComObject(current);
                        return null;
                    }

                    // Release the previous current before moving to next level
                    Marshal.ReleaseComObject(current);
                    current = found;
                }

                return current;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"FolderMonitorService.ResolveFolderPath: {ex.Message}");
                return null;
            }
            finally
            {
                if (rootFolders != null)
                {
                    Marshal.ReleaseComObject(rootFolders);
                }
            }
        }
    }
}
