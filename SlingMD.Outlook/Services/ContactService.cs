using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Handles contact-related features like generating concise display names, creating/looking up
    /// contact notes inside the vault and building wiki-links for email addresses.
    /// </summary>
    public class ContactService
    {
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ObsidianSettings _settings;

        public ContactService(FileService fileService, TemplateService templateService)
        {
            _fileService = fileService;
            _templateService = templateService;
            _settings = fileService.GetSettings();
        }

        /// <summary>
        /// Returns a shortened version of <paramref name="fullName"/> that is better suited for filenames
        /// and note titles. Parenthesised suffixes are removed and first/last-name initials are applied.
        /// </summary>
        public string GetShortName(string fullName)
        {
            string cleanName = _fileService.CleanFileName(fullName);

            int parenIndex = cleanName.IndexOf('(');
            if (parenIndex > 0)
            {
                cleanName = cleanName.Substring(0, parenIndex).Trim();
            }

            string[] parts = cleanName.Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return "Unknown";
            }

            if (parts.Length == 1)
            {
                return parts[0].Length > 10 ? parts[0].Substring(0, 10) : parts[0];
            }

            string firstName = parts[0].Length > 10 ? parts[0].Substring(0, 10) : parts[0];
            string lastInitial = parts[parts.Length - 1].Substring(0, 1).ToUpper();
            return $"{firstName}{lastInitial}";
        }

        /// <summary>
        /// Attempts to resolve the SMTP address for the sender of <paramref name="mail"/>.
        /// Falls back to <see cref="MailItem.SenderEmailAddress"/> when the property accessor fails.
        /// </summary>
        public string GetSenderEmail(MailItem mail)
        {
            try
            {
                const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                return mail.PropertyAccessor.GetProperty(PrSmtpAddress);
            }
            catch
            {
                return mail.SenderEmailAddress;
            }
        }

        /// <summary>
        /// Builds a list of Obsidian wiki-links (e.g. <c>[[Jane Doe]]</c>) for the chosen recipient type.
        /// </summary>
        public List<string> BuildLinkedNames(Recipients recipients, OlMailRecipientType type)
        {
            List<string> names = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)type)
                    {
                        names.Add($"[[{recipient.Name}]]");
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return names;
        }

        /// <summary>
        /// Collects plain email addresses for recipients of the specified <paramref name="type"/>.
        /// </summary>
        public List<string> BuildEmailList(Recipients recipients, OlMailRecipientType type)
        {
            List<string> emails = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)type)
                    {
                        try
                        {
                            const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                            string email = recipient.PropertyAccessor.GetProperty(PrSmtpAddress);
                            if (!string.IsNullOrEmpty(email))
                            {
                                emails.Add(email);
                            }
                        }
                        catch
                        {
                            if (!string.IsNullOrEmpty(recipient.Address))
                            {
                                emails.Add(recipient.Address);
                            }
                        }
                    }
                }
                finally
                {
                    if (recipient != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }

            return emails;
        }

        /// <summary>
        /// Quick existence check for a contact note. Depending on user preference the entire vault may be
        /// searched in addition to the dedicated contacts folder.
        /// </summary>
        public bool ContactExists(string contactName)
        {
            try
            {
                string cleanName = _fileService.CleanFileName(contactName);
                string configuredFileName = BuildContactFileName(contactName);
                string contactsFolder = _settings.GetContactsPath();
                string configuredPath = Path.Combine(contactsFolder, configuredFileName + ".md");
                string legacyPath = Path.Combine(contactsFolder, cleanName + ".md");

                if (File.Exists(configuredPath) || File.Exists(legacyPath))
                {
                    return true;
                }

                if (_settings.SearchEntireVaultForContacts)
                {
                    string vaultPath = _settings.GetFullVaultPath();
                    string[] matchingConfiguredFiles = Directory.GetFiles(vaultPath, configuredFileName + ".md", SearchOption.AllDirectories);
                    if (matchingConfiguredFiles.Length > 0)
                    {
                        return true;
                    }

                    if (!string.Equals(configuredFileName, cleanName, StringComparison.OrdinalIgnoreCase))
                    {
                        string[] matchingLegacyFiles = Directory.GetFiles(vaultPath, cleanName + ".md", SearchOption.AllDirectories);
                        if (matchingLegacyFiles.Length > 0)
                        {
                            return true;
                        }
                    }

                    string[] allMarkdownFiles = Directory.GetFiles(vaultPath, "*.md", SearchOption.AllDirectories);
                    const int MaxFilesToSearch = 5000;
                    if (allMarkdownFiles.Length > MaxFilesToSearch)
                    {
                        return false;
                    }

                    string searchPattern = $"[[{contactName}]]";
                    foreach (string mdFile in allMarkdownFiles)
                    {
                        try
                        {
                            int linesRead = 0;
                            const int MaxLinesToRead = 100;
                            foreach (string line in File.ReadLines(mdFile))
                            {
                                if (line.Contains(searchPattern))
                                {
                                    return true;
                                }

                                linesRead++;
                                if (linesRead >= MaxLinesToRead)
                                {
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                return false;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Creates a stub markdown note for <paramref name="contactName"/> inside the configured contacts
        /// folder and populates it with a Dataview script that lists every email mentioning the contact.
        /// </summary>
        public void CreateContactNote(string contactName)
        {
            if (!_settings.EnableContactSaving)
            {
                return;
            }

            string contactsFolder = _settings.GetContactsPath();
            string fileNameNoExtension = BuildContactFileName(contactName);
            string filePath = Path.Combine(contactsFolder, fileNameNoExtension + ".md");
            _fileService.EnsureDirectoryExists(contactsFolder);

            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", contactName },
                { "type", "contact" },
                { "created", DateTime.Now.ToString("yyyy-MM-dd HH:mm") },
                { "tags", new List<string> { "contact" } }
            };

            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = metadata,
                ContactName = contactName,
                ContactShortName = GetShortName(contactName),
                Created = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                FileName = fileNameNoExtension + ".md",
                FileNameWithoutExtension = fileNameNoExtension
            };

            string content = _templateService.RenderContactContent(context);
            _fileService.WriteUtf8File(filePath, content);
        }

        private string BuildContactFileName(string contactName)
        {
            string cleanName = _fileService.CleanFileName(contactName);
            Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "ContactName", contactName ?? string.Empty },
                { "ContactShortName", GetShortName(contactName ?? string.Empty) },
                { "CleanContactName", cleanName }
            };

            return _templateService.RenderFilename(_settings.ContactFilenameFormat, replacements, cleanName);
        }
    }
}
