using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Handles contact-related features like generating concise display names, creating/looking up
    /// contact notes inside the vault and building wiki-links for email addresses.
    /// </summary>
    public class ContactService
    {
        private const string CommunicationHistoryHeading = "## Communication History";
        private const string LegacyEmailHistoryHeading = "## Email History";
        private const string NotesHeading = "## Notes";

        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ObsidianSettings _settings;
        private readonly ContactNameParser _contactNameParser;
        private readonly ContactLinkFormatter _contactLinkFormatter;
        private readonly DateFormatter _dateFormatter;
        private readonly IClock _clock;
        private static readonly MarkdownSectionFinder SectionFinder = new MarkdownSectionFinder();

        public ContactService(FileService fileService, TemplateService templateService, IClock clock = null)
        {
            _fileService = fileService;
            _templateService = templateService;
            _settings = fileService.GetSettings();
            _contactNameParser = new ContactNameParser();
            _contactLinkFormatter = new ContactLinkFormatter();
            _dateFormatter = new DateFormatter();
            _clock = clock ?? new SystemClock();
        }

        /// <summary>
        /// Populates the new ContactName fields (FirstName/LastName/etc.) on the given context
        /// by parsing the ContactName + Email through ContactNameParser. Idempotent.
        /// </summary>
        private void PopulateNameParts(ContactTemplateContext context)
        {
            ContactName parsed = _contactNameParser.Parse(context.ContactName, context.Email);
            context.FirstName = parsed.FirstName ?? string.Empty;
            context.LastName = parsed.LastName ?? string.Empty;
            context.MiddleName = parsed.MiddleName ?? string.Empty;
            context.Suffix = parsed.Suffix ?? string.Empty;
            context.FullName = parsed.FullName ?? string.Empty;
            context.DisplayName = parsed.DisplayName ?? string.Empty;
        }

        /// <summary>
        /// Formats a display name as a contact link using the configured <see cref="ObsidianSettings.ContactLinkFormat"/>.
        /// </summary>
        private string FormatContactLink(string displayName, string email)
        {
            ContactName parsed = _contactNameParser.Parse(displayName, email);
            string formatted = _contactLinkFormatter.Format(parsed, _settings.ContactLinkFormat);
            return string.IsNullOrEmpty(formatted) ? $"[[{displayName}]]" : formatted;
        }

        /// <summary>
        /// Returns a shortened version of <paramref name="fullName"/> that is better suited for filenames
        /// and note titles. Parenthesised suffixes are removed and first/last-name initials are applied.
        /// </summary>
        public string GetFilenameSafeShortName(string fullName)
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
                return mail.PropertyAccessor.GetProperty(MapiPropertyTags.PrSmtpAddress);
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
                        string email = null;
                        try
                        {
                            email = GetSMTPEmailAddress(recipient);
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Warning($"ContactService.BuildLinkedNames: GetSMTPEmailAddress failed for recipient '{recipient.Name}': {ex.Message}");
                        }
                        names.Add(FormatContactLink(recipient.Name, email));
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
                            string email = recipient.PropertyAccessor.GetProperty(MapiPropertyTags.PrSmtpAddress);
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
        /// Attempts to resolve the SMTP address for a meeting <paramref name="recipient"/>.
        /// Falls back to <see cref="Recipient.Address"/> when the property accessor fails.
        /// </summary>
        public string GetSMTPEmailAddress(Recipient recipient)
        {
            try
            {
                return recipient.PropertyAccessor.GetProperty(MapiPropertyTags.PrSmtpAddress) as string ?? recipient.Address;
            }
            catch
            {
                return recipient.Address;
            }
        }

        /// <summary>
        /// Builds a list of Obsidian wiki-links (e.g. <c>[[Jane Doe]]</c>) filtered by one or more
        /// <see cref="OlMeetingRecipientType"/> values.
        /// </summary>
        public List<string> BuildLinkedNames(Recipients recipients, params OlMeetingRecipientType[] types)
        {
            List<string> linkedNames = new List<string>();
            HashSet<int> typeSet = new HashSet<int>();
            foreach (OlMeetingRecipientType type in types)
            {
                typeSet.Add((int)type);
            }

            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (typeSet.Contains(recipient.Type))
                    {
                        string name = recipient.Name;
                        if (!string.IsNullOrEmpty(name))
                        {
                            string email = null;
                            try
                            {
                                email = GetSMTPEmailAddress(recipient);
                            }
                            catch (System.Exception ex)
                            {
                                Logger.Instance.Warning($"ContactService.BuildLinkedNames(meeting): GetSMTPEmailAddress failed for recipient '{name}': {ex.Message}");
                            }
                            linkedNames.Add(FormatContactLink(name, email));
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

            return linkedNames;
        }

        /// <summary>
        /// Collects plain email addresses for recipients matching the specified meeting role types.
        /// </summary>
        public List<string> BuildEmailList(Recipients recipients, IEnumerable<OlMeetingRecipientType> types)
        {
            List<string> emails = new List<string>();
            HashSet<int> typeSet = new HashSet<int>();
            foreach (OlMeetingRecipientType type in types)
            {
                typeSet.Add((int)type);
            }

            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (typeSet.Contains(recipient.Type))
                    {
                        try
                        {
                            string email = recipient.PropertyAccessor.GetProperty(MapiPropertyTags.PrSmtpAddress) as string;
                            if (!string.IsNullOrEmpty(email))
                            {
                                emails.Add(email);
                            }
                            else if (!string.IsNullOrEmpty(recipient.Address))
                            {
                                emails.Add(recipient.Address);
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
        /// Extracts conference room and equipment names from <see cref="OlMeetingRecipientType.olResource"/>
        /// recipients.
        /// </summary>
        public List<string> GetMeetingResourceData(Recipients recipients)
        {
            List<string> resources = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)OlMeetingRecipientType.olResource)
                    {
                        string name = recipient.Name;
                        if (!string.IsNullOrEmpty(name))
                        {
                            resources.Add(name);
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

            return resources;
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
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Warning($"ContactService.ContactExists: read lines failed for '{mdFile}': {ex.Message}");
                            continue;
                        }
                    }
                }

                return false;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ContactExists: search failed for '{contactName}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates or refreshes a contact note in the configured contacts folder.
        /// Existing notes keep user-authored content under <c>## Notes</c> while the managed
        /// communication history block is refreshed.
        /// </summary>
        public void CreateContactNote(string contactName)
        {
            if (!_settings.EnableContactSaving)
            {
                return;
            }

            string filePath = GetManagedContactNotePath(contactName);
            string fileNameNoExtension = Path.GetFileNameWithoutExtension(filePath);

            string created = _dateFormatter.Format(_clock.Now, _settings.ContactDateFormat);
            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", contactName },
                    { "type", "contact" },
                    { "created", created },
                    { "tags", new List<string> { "contact" } }
                },
                ContactName = contactName,
                ContactShortName = GetFilenameSafeShortName(contactName),
                Created = created,
                FileName = fileNameNoExtension + ".md",
                FileNameWithoutExtension = fileNameNoExtension,
                IncludeDetails = false
            };
            PopulateNameParts(context);

            CreateContactNote(context);
        }

        public void CreateContactNote(ContactTemplateContext context)
        {
            string filePath = GetManagedContactNotePath(context.ContactName);
            _fileService.EnsureDirectoryExists(_settings.GetContactsPath());

            string renderedContent = _templateService.RenderContactContent(context);

            if (!File.Exists(filePath))
            {
                _fileService.WriteUtf8File(filePath, renderedContent);
                return;
            }

            string existingContent = File.ReadAllText(filePath);
            if (context.IncludeDetails)
            {
                // Rich contact: preserve user-authored Notes section, refresh everything else
                string preservedNotes = ExtractUserNotesSection(existingContent);
                string updatedContent = ReplaceNotesSection(renderedContent, preservedNotes);
                _fileService.WriteUtf8File(filePath, updatedContent);
            }
            else
            {
                // Basic contact: preserve user content, refresh managed Communication History
                string managedSection = ExtractManagedCommunicationHistorySection(renderedContent);
                string updatedContent = MergeManagedSections(existingContent, managedSection);
                _fileService.WriteUtf8File(filePath, updatedContent);
            }
        }

        /// <summary>
        /// Extracts rich contact data from an Outlook <see cref="ContactItem"/> and returns a fully
        /// populated <see cref="ContactTemplateContext"/>. Each COM property read is individually
        /// try/caught to tolerate missing or restricted properties.
        /// </summary>
        public ContactTemplateContext ExtractContactData(ContactItem contact)
        {
            string fullName = string.Empty;
            try
            {
                fullName = contact.FullName ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read FullName failed: {ex.Message}");
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                try
                {
                    string lastName = contact.LastName ?? string.Empty;
                    string firstName = contact.FirstName ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(lastName) || !string.IsNullOrWhiteSpace(firstName))
                    {
                        fullName = string.IsNullOrWhiteSpace(firstName)
                            ? lastName
                            : string.IsNullOrWhiteSpace(lastName)
                                ? firstName
                                : $"{firstName} {lastName}";
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read FirstName/LastName failed: {ex.Message}");
                }
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                try
                {
                    fullName = contact.FileAs ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read FileAs failed: {ex.Message}");
                }
            }

            if (string.IsNullOrWhiteSpace(fullName))
            {
                fullName = "Unknown Contact";
            }

            string phone = string.Empty;
            try
            {
                phone = contact.BusinessTelephoneNumber ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read BusinessTelephoneNumber failed: {ex.Message}");
            }

            if (string.IsNullOrWhiteSpace(phone))
            {
                try
                {
                    phone = contact.MobileTelephoneNumber ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read MobileTelephoneNumber failed: {ex.Message}");
                }
            }

            if (string.IsNullOrWhiteSpace(phone))
            {
                try
                {
                    phone = contact.HomeTelephoneNumber ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read HomeTelephoneNumber failed: {ex.Message}");
                }
            }

            string email = string.Empty;
            try
            {
                email = contact.Email1Address ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read Email1Address failed: {ex.Message}");
            }

            string company = string.Empty;
            try
            {
                company = contact.CompanyName ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read CompanyName failed: {ex.Message}");
            }

            string jobTitle = string.Empty;
            try
            {
                jobTitle = contact.JobTitle ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read JobTitle failed: {ex.Message}");
            }

            string address = string.Empty;
            try
            {
                address = contact.BusinessAddress ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read BusinessAddress failed: {ex.Message}");
            }

            if (string.IsNullOrWhiteSpace(address))
            {
                try
                {
                    address = contact.HomeAddress ?? string.Empty;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ContactService.ExtractContactData: read HomeAddress failed: {ex.Message}");
                }
            }

            string birthday = string.Empty;
            try
            {
                DateTime birthdayDate = contact.Birthday;
                if (birthdayDate.Year != 4501)
                {
                    birthday = birthdayDate.ToString("yyyy-MM-dd");
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read Birthday failed: {ex.Message}");
            }

            string notes = string.Empty;
            try
            {
                notes = contact.Body ?? string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ContactService.ExtractContactData: read Body failed: {ex.Message}");
            }

            string cleanName = _fileService.CleanFileName(fullName);
            string fileNameNoExtension = _fileService.CleanFileName(fullName);

            string created = _dateFormatter.Format(_clock.Now, _settings.ContactDateFormat);
            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", fullName },
                { "type", "contact" },
                { "created", created },
                { "tags", new List<string> { "contact" } }
            };

            if (!string.IsNullOrWhiteSpace(company))
            {
                metadata["company"] = company;
            }

            if (!string.IsNullOrWhiteSpace(email))
            {
                metadata["email"] = email;
            }

            ContactTemplateContext result = new ContactTemplateContext
            {
                Metadata = metadata,
                ContactName = fullName,
                ContactShortName = GetFilenameSafeShortName(fullName),
                Created = created,
                FileName = fileNameNoExtension + ".md",
                FileNameWithoutExtension = fileNameNoExtension,
                Phone = phone,
                Email = email,
                Company = company,
                JobTitle = jobTitle,
                Address = address,
                Birthday = birthday,
                Notes = notes,
                IncludeDetails = true
            };
            PopulateNameParts(result);
            return result;
        }


        private static string ExtractUserNotesSection(string content)
        {
            int notesStart = FindSectionStart(content, NotesHeading);
            if (notesStart < 0)
            {
                return string.Empty;
            }

            return content.Substring(notesStart).TrimEnd();
        }

        private static string ReplaceNotesSection(string renderedContent, string preservedNotes)
        {
            int notesStart = FindSectionStart(renderedContent, NotesHeading);
            if (notesStart < 0)
            {
                if (string.IsNullOrWhiteSpace(preservedNotes))
                {
                    return renderedContent;
                }

                return renderedContent.TrimEnd() + Environment.NewLine + Environment.NewLine + preservedNotes + Environment.NewLine;
            }

            string prefix = renderedContent.Substring(0, notesStart).TrimEnd();
            string notesSection = string.IsNullOrWhiteSpace(preservedNotes) ? BuildEmptyNotesSection() : preservedNotes;
            return prefix + Environment.NewLine + Environment.NewLine + notesSection.TrimEnd() + Environment.NewLine;
        }

        public string GetManagedContactNotePath(string contactName)
        {
            string cleanName = _fileService.CleanFileName(contactName);
            string configuredFileName = BuildContactFileName(contactName);
            string contactsFolder = _settings.GetContactsPath();
            string configuredPath = Path.Combine(contactsFolder, configuredFileName + ".md");
            string legacyPath = Path.Combine(contactsFolder, cleanName + ".md");

            if (File.Exists(configuredPath))
            {
                return configuredPath;
            }

            if (File.Exists(legacyPath))
            {
                return legacyPath;
            }

            return configuredPath;
        }

        public bool ManagedContactNoteExists(string contactName)
        {
            return File.Exists(GetManagedContactNotePath(contactName));
        }

        private string BuildContactFileName(string contactName)
        {
            string cleanName = _fileService.CleanFileName(contactName);
            Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "ContactName", contactName ?? string.Empty },
                { "ContactShortName", GetFilenameSafeShortName(contactName ?? string.Empty) },
                { "CleanContactName", cleanName }
            };

            return _templateService.RenderFilename(_settings.ContactFilenameFormat, replacements, cleanName);
        }

        private static int FindSectionStart(string content, string heading, int startIndex = 0)
        {
            return SectionFinder.FindSectionStart(content, heading, startIndex);
        }

        private static string ExtractManagedCommunicationHistorySection(string content)
        {
            int historyStart = FindSectionStart(content, CommunicationHistoryHeading);
            if (historyStart < 0)
            {
                historyStart = FindSectionStart(content, LegacyEmailHistoryHeading);
            }

            if (historyStart < 0)
            {
                return string.Empty;
            }

            int notesStart = FindSectionStart(content, NotesHeading, historyStart);
            if (notesStart < 0)
            {
                return content.Substring(historyStart).TrimEnd();
            }

            return content.Substring(historyStart, notesStart - historyStart).TrimEnd();
        }

        private static string MergeManagedSections(string existingContent, string managedSection)
        {
            if (string.IsNullOrWhiteSpace(managedSection))
            {
                return existingContent;
            }

            int historyStart = FindSectionStart(existingContent, CommunicationHistoryHeading);
            if (historyStart < 0)
            {
                historyStart = FindSectionStart(existingContent, LegacyEmailHistoryHeading);
            }

            if (historyStart >= 0)
            {
                int notesStart = FindSectionStart(existingContent, NotesHeading, historyStart);
                string prefix = existingContent.Substring(0, historyStart).TrimEnd();
                string notesSection = notesStart >= 0 ? existingContent.Substring(notesStart).TrimStart() : BuildEmptyNotesSection();
                return JoinSections(prefix, managedSection, notesSection);
            }

            int standaloneNotesStart = FindSectionStart(existingContent, NotesHeading);
            if (standaloneNotesStart >= 0)
            {
                string prefix = existingContent.Substring(0, standaloneNotesStart).TrimEnd();
                string notesSection = existingContent.Substring(standaloneNotesStart).TrimStart();
                return JoinSections(prefix, managedSection, notesSection);
            }

            return JoinSections(existingContent.TrimEnd(), managedSection, BuildEmptyNotesSection());
        }

        private static string JoinSections(string prefix, string managedSection, string notesSection)
        {
            List<string> sections = new List<string>();
            if (!string.IsNullOrWhiteSpace(prefix))
            {
                sections.Add(prefix.TrimEnd());
            }

            sections.Add(managedSection.TrimEnd());
            sections.Add((string.IsNullOrWhiteSpace(notesSection) ? BuildEmptyNotesSection() : notesSection.TrimStart()).TrimEnd());
            return string.Join(Environment.NewLine + Environment.NewLine, sections) + Environment.NewLine;
        }

        private static string BuildEmptyNotesSection()
        {
            return NotesHeading + Environment.NewLine + Environment.NewLine;
        }
    }
}