using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public enum ContactProcessingResult
    {
        Success,
        Skipped,
        Error
    }

    /// <summary>
    /// Orchestrates the full life-cycle of turning a <see cref="ContactItem"/> into a properly
    /// formatted markdown note inside the user's Obsidian vault. Mirrors the design of
    /// <see cref="AppointmentProcessor"/> and reuses all shared services without any modifications to them.
    /// </summary>
    public class ContactProcessor
    {
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ContactService _contactService;

        private List<string> _bulkErrors = new List<string>();

        public List<string> GetBulkErrors()
        {
            List<string> errors = new List<string>(_bulkErrors);
            _bulkErrors.Clear();
            return errors;
        }

        public ContactProcessor(ObsidianSettings settings)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _contactService = new ContactService(_fileService, _templateService);
        }

        /// <summary>
        /// Converts the supplied <paramref name="contact"/> into a rich markdown note and writes it
        /// to the configured contacts folder in the Obsidian vault. When an existing note is found,
        /// the user is prompted before overwriting.
        /// </summary>
        /// <param name="contact">The contact item to export.</param>
        /// <returns>A <see cref="ContactProcessingResult"/> indicating the outcome.</returns>
        public ContactProcessingResult ProcessContact(ContactItem contact)
        {
            if (contact == null)
            {
                return ContactProcessingResult.Error;
            }

            try
            {
                // Vault path pre-check before any file writes
                string vaultPath = _settings.GetFullVaultPath();
                if (!System.IO.Directory.Exists(vaultPath))
                {
                    throw new System.IO.DirectoryNotFoundException(
                        $"Obsidian vault at \"{vaultPath}\" is not accessible. Check that the folder exists.");
                }

                ContactTemplateContext context = _contactService.ExtractContactData(contact);
                context.IncludeDetails = _settings?.ContactNoteIncludeDetails ?? true;

                bool noteExists = _contactService.ManagedContactNoteExists(context.ContactName);
                if (noteExists)
                {
                    DialogResult dialogResult = MessageBox.Show(
                        $"A contact note for \"{context.ContactName}\" already exists. Do you want to overwrite it?",
                        "Contact Note Exists",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (dialogResult == DialogResult.No)
                    {
                        return ContactProcessingResult.Skipped;
                    }
                }

                _contactService.CreateContactNote(context);
                return ContactProcessingResult.Success;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"ContactProcessor.ProcessContact failed: {ex.Message}");
                return ContactProcessingResult.Error;
            }
        }

        /// <summary>
        /// Iterates all items in <paramref name="contactsFolder"/>, exporting each
        /// <see cref="ContactItem"/> to the vault. Non-contact items are silently skipped.
        /// Errors for individual contacts are collected and available via <see cref="GetBulkErrors()"/>.
        /// </summary>
        /// <param name="contactsFolder">The Outlook MAPI folder containing contacts to export.</param>
        /// <param name="saved">Number of contacts successfully exported.</param>
        /// <param name="skipped">Number of items skipped (not a ContactItem or already exists).</param>
        /// <param name="errors">Number of contacts that failed to export.</param>
        public void ProcessAddressBook(MAPIFolder contactsFolder, out int saved, out int skipped, out int errors)
        {
            saved = 0;
            skipped = 0;
            errors = 0;

            if (contactsFolder == null)
            {
                return;
            }

            Items folderItems = null;
            int totalCount = 0;

            try
            {
                folderItems = contactsFolder.Items;
                totalCount = folderItems.Count;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"ContactProcessor.ProcessAddressBook: could not access folder items: {ex.Message}");
                _bulkErrors.Add($"Could not access folder items: {ex.Message}");
                return;
            }

            for (int i = 1; i <= totalCount; i++)
            {
                object rawItem = null;
                ContactItem contactItem = null;

                try
                {
                    rawItem = folderItems[i];
                    contactItem = rawItem as ContactItem;

                    if (contactItem == null)
                    {
                        skipped++;
                        continue;
                    }

                    int progressPercent = totalCount > 0 ? (int)((i / (double)totalCount) * 100) : 0;
                    Logger.Instance.Debug($"ContactProcessor: processing contact {i}/{totalCount} ({progressPercent}%)");

                    bool noteExists = false;
                    string contactName = string.Empty;
                    try
                    {
                        ContactTemplateContext context = _contactService.ExtractContactData(contactItem);
                        context.IncludeDetails = _settings?.ContactNoteIncludeDetails ?? true;
                        contactName = context.ContactName;

                        noteExists = _contactService.ManagedContactNoteExists(context.ContactName);
                        if (noteExists)
                        {
                            skipped++;
                            continue;
                        }

                        _contactService.CreateContactNote(context);
                        saved++;
                    }
                    catch (System.Exception ex)
                    {
                        string errorMsg = string.IsNullOrEmpty(contactName)
                            ? $"Contact {i}: {ex.Message}"
                            : $"Contact \"{contactName}\": {ex.Message}";

                        Logger.Instance.Error($"ContactProcessor.ProcessAddressBook: {errorMsg}");
                        _bulkErrors.Add(errorMsg);
                        errors++;
                    }
                }
                finally
                {
                    if (contactItem != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(contactItem);
                    }
                    else if (rawItem != null && System.Runtime.InteropServices.Marshal.IsComObject(rawItem))
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rawItem);
                    }
                }
            }

            if (folderItems != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(folderItems);
            }
        }
    }
}
