using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Ribbon;

namespace SlingMD.Outlook
{
    public partial class ThisAddIn
    {
        private ObsidianSettings _settings;
        private EmailProcessor _emailProcessor;
        private AppointmentProcessor _appointmentProcessor;
        private ContactProcessor _contactProcessor;
        private FileService _fileService;
        private NotificationService _notificationService;
        private FolderMonitorService _folderMonitorService;
        private AutoSlingService _autoSlingService;
        private FlagMonitorService _flagMonitorService;
        private SlingRibbon _ribbon;
        private Explorer _activeExplorer;
        private bool _startupComplete;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new SlingRibbon(this);
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            bool isFirstLaunchAfterInstall;

            _settings = LoadSettings(out isFirstLaunchAfterInstall);
            ValidateStartupHealth();
            _emailProcessor = new EmailProcessor(_settings);
            _appointmentProcessor = new AppointmentProcessor(_settings);
            _contactProcessor = new ContactProcessor(_settings);
            _fileService = new FileService(_settings);
            _notificationService = new NotificationService(_settings);

            if (isFirstLaunchAfterInstall && !_settings.HasShownSupportPrompt)
            {
                ShowFirstRunSupportPrompt();
            }

            try
            {
                _activeExplorer = Application.ActiveExplorer();
                if (_activeExplorer != null)
                {
                    _activeExplorer.SelectionChange += Explorer_SelectionChange;
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ThisAddIn.Startup: could not hook explorer selection change: {ex.Message}");
            }

            if (_settings.WatchedFolders != null && _settings.WatchedFolders.Count > 0)
            {
                _folderMonitorService = new FolderMonitorService(_settings, _emailProcessor, _notificationService, Application);
                _folderMonitorService.StartWatching(_settings.WatchedFolders);
            }

            _startupComplete = true;

            _autoSlingService = new AutoSlingService(_settings, _emailProcessor, _notificationService);
            _autoSlingService.Start(Application);

            if (_settings.EnableFlagToSling)
            {
                _flagMonitorService = new FlagMonitorService(_settings, _emailProcessor, _notificationService);
                _flagMonitorService.Start(Application);
            }

            _ribbon?.UpdateSlingButtonLabel(GetSelectedItemLabel());
        }

        private void Explorer_SelectionChange()
        {
            _ribbon?.UpdateSlingButtonLabel(GetSelectedItemLabel());
        }

        public string GetSelectedItemLabel()
        {
            if (!_startupComplete)
            {
                return "Sling";
            }

            try
            {
                Explorer explorer = Application.ActiveExplorer();
                if (explorer == null || explorer.Selection.Count == 0)
                {
                    return "Sling";
                }

                object selected = null;
                try
                {
                    selected = explorer.Selection[1];
                    if (selected is MailItem) return "Sling Email";
                    if (selected is AppointmentItem) return "Sling Appointment";
                    if (selected is ContactItem) return "Sling Contact";
                }
                finally
                {
                    if (selected != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(selected);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ThisAddIn.GetSelectedItemLabel: {ex.Message}");
            }

            return "Sling";
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _autoSlingService?.Shutdown();
            _flagMonitorService?.Stop();
            _folderMonitorService?.StopWatching();

            if (_activeExplorer != null)
            {
                _activeExplorer.SelectionChange -= Explorer_SelectionChange;
                _activeExplorer = null;
            }

            if (_settings != null)
            {
                _settings.Save();
            }
        }

        private void ValidateStartupHealth()
        {
            string vaultPath = _settings.GetFullVaultPath();
            string inboxPath = _settings.GetInboxPath();

            if (!System.IO.Directory.Exists(vaultPath))
            {
                Logger.Instance.Warning($"Startup health: vault path does not exist: {vaultPath}");
                MessageBox.Show(
                    $"SlingMD: Your Obsidian vault at \"{vaultPath}\" is not accessible.\n\n"
                    + "Exports will fail until you configure Settings.",
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!System.IO.Directory.Exists(inboxPath))
            {
                Logger.Instance.Info($"Startup health: inbox folder does not exist yet: {inboxPath}. Will be created on first export.");
            }

            Logger.Instance.Info($"Startup health: vault path OK: {vaultPath}");
        }

        private ObsidianSettings LoadSettings(out bool isFirstLaunchAfterInstall)
        {
            ObsidianSettings settings = new ObsidianSettings();
            isFirstLaunchAfterInstall = !settings.HasSavedSettings();
            settings.Load();
            return settings;
        }

        private void ShowFirstRunSupportPrompt()
        {
            SupportService.ShowBuyMeACoffeePrompt();
            _settings.HasShownSupportPrompt = true;

            try
            {
                _settings.Save();
            }
            catch (ArgumentException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
            catch (IOException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
        }

        private void ShowFirstRunStateSaveWarning(string errorMessage)
        {
            MessageBox.Show(
                "SlingMD showed the support prompt but could not save the first-run state."
                    + Environment.NewLine
                    + Environment.NewLine
                    + errorMessage,
                "SlingMD",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public async void ProcessSelection()
        {
            try
            {
                Explorer explorer = Application.ActiveExplorer();
                if (explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email, appointment, or contact first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                object selected = explorer.Selection[1];
                MailItem mail = selected as MailItem;
                AppointmentItem appointment = selected as AppointmentItem;
                ContactItem contact = selected as ContactItem;

                if (mail != null)
                {
                    await _emailProcessor.ProcessEmail(mail);
                }
                else if (appointment != null)
                {
                    await _appointmentProcessor.ProcessAppointment(appointment, bulkMode: false);
                }
                else if (contact != null)
                {
                    ContactProcessingResult contactResult = _contactProcessor.ProcessContact(contact);
                    if (contactResult == ContactProcessingResult.Success && _settings.LaunchObsidian)
                    {
                        _fileService.LaunchObsidian(_settings.VaultName, _settings.GetContactsPath());
                    }
                }
                else
                {
                    MessageBox.Show("Please select an email, appointment, or contact.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving item: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async void ProcessCurrentAppointment()
        {
            try
            {
                Inspector inspector = Application.ActiveInspector();
                if (inspector == null)
                {
                    MessageBox.Show("No item is currently open.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                AppointmentItem appointment = inspector.CurrentItem as AppointmentItem;
                if (appointment == null)
                {
                    MessageBox.Show("The open item is not an appointment.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (!appointment.Saved)
                {
                    DialogResult choice = MessageBox.Show(
                        "This appointment has unsaved changes. Save before exporting to Obsidian?",
                        "Unsaved Changes",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    if (choice == DialogResult.Cancel)
                    {
                        return;
                    }

                    if (choice == DialogResult.Yes)
                    {
                        appointment.Save();
                    }
                }

                await _appointmentProcessor.ProcessAppointment(appointment, bulkMode: false);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving appointment: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ProcessSelectedEmail()
        {
            ProcessSelection();
        }

        public void SlingAllContacts()
        {
            int saved = 0;
            int skipped = 0;
            int errors = 0;

            MAPIFolder contactsFolder = null;
            try
            {
                try
                {
                    contactsFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                    _contactProcessor.ProcessAddressBook(contactsFolder, out saved, out skipped, out errors);
                }
                finally
                {
                    if (contactsFolder != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(contactsFolder);
                    }
                }

                System.Collections.Generic.List<string> bulkErrors = _contactProcessor.GetBulkErrors();
                string summary = string.Format(
                    "Saved {0} contacts.\nSkipped: {1} (already exist or not a contact)\nErrors: {2}",
                    saved, skipped, errors);

                if (bulkErrors.Count > 0)
                {
                    summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
                }

                MessageBox.Show(
                    summary,
                    "Sling All Contacts",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (_settings.LaunchObsidian && saved > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, _settings.GetContactsPath());
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error exporting contacts: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void SaveTodaysAppointments()
        {
            int saved = 0;
            int skipped = 0;
            int errors = 0;
            int total = 0;

            try
            {
                Accounts accounts = Application.Session.Accounts;
                try
                {
                    foreach (Account account in accounts)
                    {
                        MAPIFolder calendar = null;
                        Items items = null;
                        Items restricted = null;
                        try
                        {
                            calendar = account.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                            items = calendar.Items;
                            items.IncludeRecurrences = true;
                            items.Sort("[Start]");

                            DateTime today = DateTime.Today;
                            DateTime tomorrow = today.AddDays(1);
                            string filter = string.Format(
                                "[Start] >= '{0}' AND [Start] < '{1}'",
                                today.ToString("g"),
                                tomorrow.ToString("g"));
                            restricted = items.Restrict(filter);

                            foreach (object item in restricted)
                            {
                                AppointmentItem appointment = item as AppointmentItem;
                                if (appointment == null) continue;

                                try
                                {
                                    total++;

                                    AppointmentProcessingResult result =
                                        await _appointmentProcessor.ProcessAppointment(
                                            appointment, bulkMode: true);

                                    switch (result)
                                    {
                                        case AppointmentProcessingResult.Success:
                                            saved++;
                                            break;
                                        case AppointmentProcessingResult.Skipped:
                                            skipped++;
                                            break;
                                        case AppointmentProcessingResult.Error:
                                            errors++;
                                            break;
                                    }
                                }
                                finally
                                {
                                    if (appointment != null)
                                    {
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appointment);
                                    }
                                }
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            errors++;
                        }
                        finally
                        {
                            if (restricted != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(restricted);
                            if (items != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                            if (calendar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(calendar);
                        }
                    }
                }
                finally
                {
                    if (accounts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accounts);
                }

                List<string> bulkErrors = _appointmentProcessor.GetBulkErrors();
                string summary = string.Format(
                    "Saved {0}/{1} appointments.\nSkipped: {2} (duplicates/cancelled)\nErrors: {3}",
                    saved, total, skipped, errors);

                if (bulkErrors.Count > 0)
                {
                    summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
                }

                MessageBox.Show(
                    summary,
                    "Save Today's Appointments",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (_settings.LaunchObsidian && saved > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, _settings.GetAppointmentsPath());
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error saving today's appointments: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void SaveAppointmentRange(DateTime start, DateTime end)
        {
            int saved = 0;
            int skipped = 0;
            int errors = 0;
            int total = 0;

            try
            {
                Accounts accounts = Application.Session.Accounts;
                try
                {
                    foreach (Account account in accounts)
                    {
                        MAPIFolder calendar = null;
                        Items items = null;
                        Items restricted = null;
                        try
                        {
                            calendar = account.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                            items = calendar.Items;
                            items.IncludeRecurrences = true;
                            items.Sort("[Start]");

                            DateTime rangeEnd = end.AddDays(1);
                            string filter = string.Format(
                                "[Start] >= '{0}' AND [Start] < '{1}'",
                                start.ToString("g"),
                                rangeEnd.ToString("g"));
                            restricted = items.Restrict(filter);

                            foreach (object item in restricted)
                            {
                                AppointmentItem appointment = item as AppointmentItem;
                                if (appointment == null) continue;

                                try
                                {
                                    total++;

                                    AppointmentProcessingResult result =
                                        await _appointmentProcessor.ProcessAppointment(
                                            appointment, bulkMode: true);

                                    switch (result)
                                    {
                                        case AppointmentProcessingResult.Success:
                                            saved++;
                                            break;
                                        case AppointmentProcessingResult.Skipped:
                                            skipped++;
                                            break;
                                        case AppointmentProcessingResult.Error:
                                            errors++;
                                            break;
                                    }
                                }
                                finally
                                {
                                    if (appointment != null)
                                    {
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appointment);
                                    }
                                }
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            errors++;
                        }
                        finally
                        {
                            if (restricted != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(restricted);
                            if (items != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                            if (calendar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(calendar);
                        }
                    }
                }
                finally
                {
                    if (accounts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accounts);
                }

                List<string> bulkErrors = _appointmentProcessor.GetBulkErrors();
                string summary = string.Format(
                    "Saved {0}/{1} appointments.\nSkipped: {2} (duplicates/cancelled)\nErrors: {3}\nDate range: {4} to {5}",
                    saved, total, skipped, errors,
                    start.ToString("d"), end.ToString("d"));

                if (bulkErrors.Count > 0)
                {
                    summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
                }

                MessageBox.Show(
                    summary,
                    "Save Appointment Range",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (_settings.LaunchObsidian && saved > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, _settings.GetAppointmentsPath());
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error saving appointments: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void CompleteThread()
        {
            try
            {
                Explorer explorer = Application.ActiveExplorer();
                if (explorer == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MailItem mail = explorer.Selection[1] as MailItem;
                if (mail == null)
                {
                    MessageBox.Show("Please select an email to complete its thread.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                TemplateService templateService = new TemplateService(_fileService);
                ThreadService threadService = new ThreadService(_fileService, templateService, _settings);
                string conversationId = threadService.GetConversationId(mail);

                ThreadCompletionService completionService = new ThreadCompletionService(_fileService, _settings);
                MAPIFolder inbox = null;
                List<MailItem> missingEmails;
                try
                {
                    inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    missingEmails = completionService.FindMissingEmails(conversationId, inbox, threadService.GetConversationId);
                }
                finally
                {
                    if (inbox != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                    }
                }

                if (missingEmails.Count == 0)
                {
                    MessageBox.Show("All emails in this thread have already been slung to Obsidian.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (ThreadCompletionDialog dialog = new ThreadCompletionDialog(missingEmails))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        List<MailItem> selected = dialog.SelectedEmails;
                        int processed = 0;
                        foreach (MailItem selectedMail in selected)
                        {
                            await _emailProcessor.ProcessEmail(selectedMail);
                            processed++;
                        }
                        MessageBox.Show($"Slung {processed} emails from this thread.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error completing thread: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ShowSettings()
        {
            try
            {
                using (SettingsForm form = new SettingsForm(_settings))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Signal shutdown to in-flight handlers before recreating
                        _autoSlingService?.Shutdown();
                        _flagMonitorService?.SignalShutdown();
                        _flagMonitorService?.Stop();
                        _folderMonitorService?.SignalShutdown();
                        _folderMonitorService?.StopWatching();

                        // Settings are automatically saved by the form
                        // Recreate processors with new settings
                        _emailProcessor = new EmailProcessor(_settings);
                        _appointmentProcessor = new AppointmentProcessor(_settings);
                        _contactProcessor = new ContactProcessor(_settings);
                        _fileService = new FileService(_settings);
                        _notificationService = new NotificationService(_settings);

                        // Restart folder monitoring with updated watched folders
                        _folderMonitorService = null;
                        if (_settings.WatchedFolders != null && _settings.WatchedFolders.Count > 0)
                        {
                            _folderMonitorService = new FolderMonitorService(_settings, _emailProcessor, _notificationService, Application);
                            _folderMonitorService.StartWatching(_settings.WatchedFolders);
                        }

                        // Restart auto-sling and flag monitor with updated settings
                        _autoSlingService = new AutoSlingService(_settings, _emailProcessor, _notificationService);
                        _autoSlingService.Start(Application);

                        _flagMonitorService = null;
                        if (_settings.EnableFlagToSling)
                        {
                            _flagMonitorService = new FlagMonitorService(_settings, _emailProcessor, _notificationService);
                            _flagMonitorService.Start(Application);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
