using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Text;
using System.IO;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Handles contact-related features like generating concise display names, creating/looking up
    /// contact notes inside the vault and building wiki-links for email addresses.  All heavy file
    /// operations are delegated to <see cref="FileService"/> to keep the class testable.
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
        /// <param name="fullName">The display name coming from Outlook.</param>
        /// <returns>A condensed name, maximum 11 characters long.</returns>
        public string GetShortName(string fullName)
        {
            // Clean the name first
            string cleanName = _fileService.CleanFileName(fullName);
            
            // If name contains parentheses, take what's before them
            int parenIndex = cleanName.IndexOf('(');
            if (parenIndex > 0)
            {
                cleanName = cleanName.Substring(0, parenIndex).Trim();
            }

            // Split into parts
            string[] parts = cleanName.Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (parts.Length == 0) return "Unknown";
            if (parts.Length == 1) return parts[0].Length > 10 ? parts[0].Substring(0, 10) : parts[0];
            
            // For multiple parts, use first name and last name initial
            string firstName = parts[0].Length > 10 ? parts[0].Substring(0, 10) : parts[0];
            string lastInitial = parts[parts.Length - 1].Substring(0, 1).ToUpper();
            return $"{firstName}{lastInitial}";
        }

        /// <summary>
        /// Attempts to resolve the SMTP address for the sender of <paramref name="mail"/>.
        /// Falls back to <see cref="MailItem.SenderEmailAddress"/> when the property accessor fails.
        /// </summary>
        /// <param name="mail">Mail item being processed.</param>
        /// <returns>The best guess SMTP email address.</returns>
        public string GetSenderEmail(MailItem mail)
        {
            try
            {
                // Try to get SMTP address using PR_SMTP_ADDRESS property
                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                return mail.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
            }
            catch
            {
                // Fallback to SenderEmailAddress
                return mail.SenderEmailAddress;
            }
        }

        /// <summary>
        /// Builds a list of Obsidian wiki-links (e.g. <c>[[Jane Doe]]</c>) for the chosen recipient type.
        /// </summary>
        /// <param name="recipients">The full recipients collection.</param>
        /// <param name="type">Recipient classification (<c>To</c>, <c>Cc</c>, etc.).</param>
        /// <returns>A list which can directly be serialised into YAML front-matter.</returns>
        public List<string> BuildLinkedNames(Recipients recipients, OlMailRecipientType type)
        {
            var names = new List<string>();
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
                    // Release each Recipient COM object to prevent memory leaks
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
            var emails = new List<string>();
            foreach (Recipient recipient in recipients)
            {
                try
                {
                    if (recipient.Type == (int)type)
                    {
                        try
                        {
                            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                            string email = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
                            if (!string.IsNullOrEmpty(email))
                            {
                                emails.Add(email);
                            }
                        }
                        catch
                        {
                            // Fallback to Address property
                            if (!string.IsNullOrEmpty(recipient.Address))
                            {
                                emails.Add(recipient.Address);
                            }
                        }
                    }
                }
                finally
                {
                    // Release each Recipient COM object to prevent memory leaks
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
        /// <param name="contactName">Display name of the contact.</param>
        /// <returns><c>true</c> when a note already exists.</returns>
        public bool ContactExists(string contactName)
        {
            // Check if a note for this contact already exists
            try
            {
                string cleanName = _fileService.CleanFileName(contactName);
                
                // First check the dedicated contacts folder
                string contactsFolder = _settings.GetContactsPath();
                string filePath = Path.Combine(contactsFolder, $"{cleanName}.md");
                
                if (File.Exists(filePath))
                {
                    return true;
                }
                
                // If setting enabled, search the entire vault
                if (_settings.SearchEntireVaultForContacts)
                {
                    string vaultPath = _settings.GetFullVaultPath();

                    // Search for any markdown file with the contact's name in the title
                    // or with a [[ContactName]] link pattern

                    // Option 1: File name matches the contact name
                    string[] matchingFiles = Directory.GetFiles(vaultPath, $"{cleanName}.md", SearchOption.AllDirectories);
                    if (matchingFiles.Length > 0)
                    {
                        return true;
                    }

                    // Option 2: Search for markdown files with the contact name in wikilinks
                    // This is more expensive but necessary for a complete search
                    string[] allMarkdownFiles = Directory.GetFiles(vaultPath, "*.md", SearchOption.AllDirectories);

                    // Performance safeguard: Skip vault-wide search if there are too many files
                    // This prevents multi-second delays on large vaults (10,000+ notes)
                    const int MAX_FILES_TO_SEARCH = 5000;
                    if (allMarkdownFiles.Length > MAX_FILES_TO_SEARCH)
                    {
                        // Fall back to contacts folder only when vault is too large
                        return false;
                    }

                    // Prepare search patterns for the contact (exact match with brackets)
                    string searchPattern = $"[[{contactName}]]";

                    foreach (string mdFile in allMarkdownFiles)
                    {
                        try
                        {
                            // Performance optimization: Only read first 100 lines instead of entire file
                            // Contact mentions are typically in frontmatter or early in the note
                            int linesRead = 0;
                            const int MAX_LINES_TO_READ = 100;

                            foreach (string line in File.ReadLines(mdFile))
                            {
                                if (line.Contains(searchPattern))
                                {
                                    return true;
                                }

                                linesRead++;
                                if (linesRead >= MAX_LINES_TO_READ)
                                {
                                    break;
                                }
                            }
                        }
                        catch
                        {
                            // Skip files that can't be read
                            continue;
                        }
                    }
                }
                
                // If we get here, the contact doesn't exist
                return false;
            }
            catch (System.Exception)
            {
                // In case of any error, return false which will just treat it as a new contact
                return false;
            }
        }

        /// <summary>
        /// Creates a stub markdown note for <paramref name="contactName"/> inside the configured contacts
        /// folder and populates it with a dataview script that lists every email mentioning the contact.
        /// </summary>
        public void CreateContactNote(string contactName)
        {
            // Check if contact saving is enabled in settings
            if (!_settings.EnableContactSaving)
            {
                return; // Skip contact note creation if disabled
            }

            // Clean the contact name for file safety
            string cleanName = _fileService.CleanFileName(contactName);
            
            // Build the file path in the contacts folder within the vault
            string contactsFolder = _settings.GetContactsPath();
            string filePath = Path.Combine(contactsFolder, $"{cleanName}.md");

            // Ensure the contacts directory exists
            _fileService.EnsureDirectoryExists(contactsFolder);

            // Build the note content with frontmatter
            var metadata = new Dictionary<string, object>
            {
                { "title", contactName },
                { "type", "contact" },
                { "created", DateTime.Now.ToString("yyyy-MM-dd HH:mm") },
                { "tags", "contact" }
            };

            var content = new StringBuilder();
            content.Append(_templateService.BuildFrontMatter(metadata));
            content.AppendLine($"# {contactName}");
            content.AppendLine();
            content.AppendLine("## Communication History");
            content.AppendLine();
            content.AppendLine("```dataviewjs");
            content.AppendLine("// Find all emails where this person is mentioned");
            content.AppendLine("// Use title from frontmatter (original name) rather than file.name (cleaned name)");
            content.AppendLine("const contact = dv.current().title || dv.current().file.name;");
            content.AppendLine();
            content.AppendLine("// Helper to check if a field contains this contact");
            content.AppendLine("// Handles both Dataview Link objects and plain strings");
            content.AppendLine("function containsContact(field, contactName) {");
            content.AppendLine("    if (!field) return false;");
            content.AppendLine("    // Handle Dataview Link objects (have .path property)");
            content.AppendLine("    if (field.path) return field.path === contactName;");
            content.AppendLine("    // Handle string format - check for [[Name]] or exact match");
            content.AppendLine("    const str = String(field);");
            content.AppendLine("    return str.includes(`[[${contactName}]]`) || str === contactName;");
            content.AppendLine("}");
            content.AppendLine();
            content.AppendLine("// Helper to check arrays (to/cc fields can be arrays)");
            content.AppendLine("function checkArray(arr, contactName) {");
            content.AppendLine("    if (!arr) return false;");
            content.AppendLine("    if (!Array.isArray(arr)) return containsContact(arr, contactName);");
            content.AppendLine("    return arr.some(item => containsContact(item, contactName));");
            content.AppendLine("}");
            content.AppendLine();

            // Get the actual email tag from settings (use first tag from DefaultNoteTags, fallback to empty string to search all pages)
            string emailTag = (_settings.DefaultNoteTags != null && _settings.DefaultNoteTags.Count > 0)
                ? $"'#{_settings.DefaultNoteTags[0]}'"
                : "\"\"";

            content.AppendLine($"const emails = dv.pages({emailTag})");
            content.AppendLine("    .where(p => {");
            content.AppendLine("        return containsContact(p.from, contact) ||");
            content.AppendLine("               checkArray(p.to, contact) ||");
            content.AppendLine("               checkArray(p.cc, contact);");
            content.AppendLine("    })");
            content.AppendLine("    .sort(p => p.date, 'desc');");
            content.AppendLine();
            content.AppendLine("dv.table([\"Date\", \"Subject\", \"Type\"],");
            content.AppendLine("    emails.map(p => {");
            content.AppendLine("        // Determine message type");
            content.AppendLine("        const isFrom = containsContact(p.from, contact);");
            content.AppendLine("        const isTo = checkArray(p.to, contact);");
            content.AppendLine("        const type = isFrom ? \"From\" : isTo ? \"To\" : \"CC\";");
            content.AppendLine("        return [p.date, p.file.link, type];");
            content.AppendLine("    })");
            content.AppendLine(");");
            content.AppendLine("```");
            content.AppendLine();
            content.AppendLine("## Notes");
            content.AppendLine();

            // Write the file
            _fileService.WriteUtf8File(filePath, content.ToString());
        }
    }
} 