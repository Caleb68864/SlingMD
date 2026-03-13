using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
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
        private SlingRibbon _ribbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new SlingRibbon(this);
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            bool isFirstLaunchAfterInstall;

            _settings = LoadSettings(out isFirstLaunchAfterInstall);
            _emailProcessor = new EmailProcessor(_settings);
            _appointmentProcessor = new AppointmentProcessor(_settings);
            _contactProcessor = new ContactProcessor(_settings);
            _fileService = new FileService(_settings);

            if (isFirstLaunchAfterInstall && !_settings.HasShownSupportPrompt)
            {
                ShowFirstRunSupportPrompt();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Save settings when Outlook is closing
            if (_settings != null)
            {
                _settings.Save();
            }
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
                    _contactProcessor.ProcessContact(contact);
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

        public void ProcessSelectedContact()
        {
            try
            {
                Explorer explorer = Application.ActiveExplorer();
                if (explorer == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select a contact first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ContactItem contact = explorer.Selection[1] as ContactItem;
                if (contact == null)
                {
                    MessageBox.Show("The selected item is not a contact.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _contactProcessor.ProcessContact(contact);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error processing contact: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        public void ShowSettings()
        {
            try
            {
                using (SettingsForm form = new SettingsForm(_settings))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Settings are automatically saved by the form
                        // Recreate processors with new settings
                        _emailProcessor = new EmailProcessor(_settings);
                        _appointmentProcessor = new AppointmentProcessor(_settings);
                        _contactProcessor = new ContactProcessor(_settings);
                        _fileService = new FileService(_settings);
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
