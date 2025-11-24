using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Provides helper routines for detecting email conversation threads, generating folder/note names, 
    /// manipulating files that belong to a thread and updating the thread summary note.  
    /// All publicly exposed members are safe for unit-testing and are free of any Outlook UI dependencies.
    /// </summary>
    public class ThreadService
    {
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ObsidianSettings _settings;

        public ThreadService(FileService fileService, TemplateService templateService, ObsidianSettings settings)
        {
            _fileService = fileService;
            _templateService = templateService;
            _settings = settings;
        }

        /// <summary>
        /// Derives a stable 20-character hash that uniquely identifies an Outlook conversation.
        /// The method tries several strategies (ConversationTopic, PR_CONVERSATION_INDEX, normalised subject)
        /// and falls back to a random GUID segment if everything else fails. Using 20 characters instead of 16
        /// significantly reduces the probability of hash collisions while remaining compact.
        /// </summary>
        /// <param name="mail">The <see cref="MailItem"/> for which to obtain the conversation id.</param>
        /// <returns>A 20-character hexadecimal string suitable for use as an identifier inside vault front-matter.</returns>
        public string GetConversationId(MailItem mail)
        {
            try
            {
                // Try to get the conversation topic first as it's most reliable for threading
                if (!string.IsNullOrEmpty(mail.ConversationTopic))
                {
                    string normalizedSubject = mail.ConversationTopic;
                    // Remove all variations of Re, Fwd, etc. and [EXTERNAL] tags
                    normalizedSubject = Regex.Replace(normalizedSubject, @"^(?:(?:Re|Fwd|FW|RE|FWD)[- :]|\[EXTERNAL\]|\s)+", "", RegexOptions.IgnoreCase);
                    // Also remove any "Re:" that might appear after [EXTERNAL]
                    normalizedSubject = Regex.Replace(normalizedSubject, @"^Re:\s+", "", RegexOptions.IgnoreCase);
                    return BitConverter.ToString(MD5.Create()
                        .ComputeHash(Encoding.UTF8.GetBytes(normalizedSubject)))
                        .Replace("-", "").Substring(0, 20);
                }

                // Try to get the conversation index using PR_CONVERSATION_INDEX property
                const string PR_CONVERSATION_INDEX = "http://schemas.microsoft.com/mapi/proptag/0x0071001F";
                byte[] conversationIndex = (byte[])mail.PropertyAccessor.GetProperty(PR_CONVERSATION_INDEX);

                if (conversationIndex != null && conversationIndex.Length >= 22)
                {
                    // The first 22 bytes of the conversation index identify the thread
                    // Use 20 characters for better uniqueness
                    return BitConverter.ToString(conversationIndex.Take(22).ToArray())
                        .Replace("-", "").Substring(0, 20);
                }

                // If both methods fail, use the normalized subject as last resort
                string subject = mail.Subject;
                subject = Regex.Replace(subject, @"^(?:(?:Re|Fwd|FW|RE|FWD)[- :]|\[EXTERNAL\]|\s)+", "", RegexOptions.IgnoreCase);
                subject = Regex.Replace(subject, @"^Re:\s+", "", RegexOptions.IgnoreCase);
                return BitConverter.ToString(MD5.Create()
                    .ComputeHash(Encoding.UTF8.GetBytes(subject)))
                    .Replace("-", "").Substring(0, 20);
            }
            catch
            {
                return Guid.NewGuid().ToString("N").Substring(0, 20);
            }
        }

        /// <summary>
        /// Returns a folder-friendly thread name (no timestamp)
        /// </summary>
        /// <param name="mail">The email that belongs to the thread.</param>
        /// <param name="cleanSubject">A sanitised subject line, typically produced by <see cref="FileService.CleanFileName(string)"/>.</param>
        /// <param name="firstSender">Short name of the first sender in the thread.</param>
        /// <param name="firstRecipient">Short name of the first recipient ("To") in the thread.</param>
        /// <returns>The folder-friendly name without any leading "0-" prefix.</returns>
        public string GetThreadFolderName(MailItem mail, string cleanSubject, string firstSender, string firstRecipient)
        {
            string threadSubject = !string.IsNullOrEmpty(mail.ConversationTopic)
                ? mail.ConversationTopic
                : mail.Subject;
            threadSubject = _fileService.CleanFileName(threadSubject);
            if (threadSubject.Length > 50)
            {
                threadSubject = threadSubject.Substring(0, 47) + "...";
            }
            firstSender = _fileService.CleanFileName(firstSender);
            firstRecipient = _fileService.CleanFileName(firstRecipient);
            return $"{threadSubject.Trim()}-{firstSender}-{firstRecipient}".Replace("--", "-");
        }

        /// <summary>
        /// Returns a note-friendly thread name (with timestamp)
        /// </summary>
        /// <param name="mail">The email that belongs to the thread.</param>
        /// <param name="cleanSubject">A sanitised subject line, typically produced by <see cref="FileService.CleanFileName(string)"/>.</param>
        /// <param name="firstSender">Short name of the first sender in the thread.</param>
        /// <param name="firstRecipient">Short name of the first recipient ("To") in the thread.</param>
        /// <returns>The folder-friendly name without any leading "0-" prefix.</returns>
        public string GetThreadNoteName(MailItem mail, string cleanSubject, string firstSender, string firstRecipient)
        {
            string threadSubject = !string.IsNullOrEmpty(mail.ConversationTopic)
                ? mail.ConversationTopic
                : mail.Subject;
            threadSubject = _fileService.CleanFileName(threadSubject);
            if (threadSubject.Length > 50)
            {
                threadSubject = threadSubject.Substring(0, 47) + "...";
            }
            firstSender = _fileService.CleanFileName(firstSender);
            firstRecipient = _fileService.CleanFileName(firstRecipient);
            string timestamp = mail.ReceivedTime.ToString("yyyy-MM-dd HHmm");
            return $"{timestamp}-{threadSubject.Trim()}-{firstSender}-{firstRecipient}".Replace("--", "-");
        }

        /// <summary>
        /// Generates or refreshes the <c>0-threadnote.md</c> summary file that lives inside a thread folder.
        /// </summary>
        /// <param name="threadFolderPath">Full path to the thread folder inside the vault inbox.</param>
        /// <param name="threadNotePath">Full path to the summary note file.</param>
        /// <param name="conversationId">Identifier created by <see cref="GetConversationId"/>.</param>
        /// <param name="threadNoteName">Base name (without <c>0-</c> prefix) for the summary file.</param>
        /// <param name="mail">The current email being processed – used only for the title.</param>
        public Task UpdateThreadNote(string threadFolderPath, string threadNotePath, string conversationId, string threadNoteName, MailItem mail)
        {
            var templateContent = _templateService.LoadTemplate("ThreadNoteTemplate.md") ??
                                _templateService.GetDefaultThreadNoteTemplate();

            // Clean the title using the same method as thread name
            string threadTitle = mail.ConversationTopic ?? mail.Subject;
            threadTitle = _fileService.CleanFileName(threadTitle);

            var replacements = new Dictionary<string, string>
            {
                { "title", threadTitle },
                { "threadId", conversationId }
            };

            string content = _templateService.ProcessTemplate(templateContent, replacements);
            _fileService.WriteUtf8File(threadNotePath, content);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Moves an email note that was originally written to the inbox into the designated thread folder.
        /// Any existing file with the same name will be overwritten.
        /// The method also appends a deterministic suffix based on the <c>internetMessageId</c> or <c>entryId</c>
        /// found in the note front-matter to avoid collisions between messages that share the same timestamp.
        /// Includes path length validation to prevent exceeding Windows' 260-character path limit by truncating
        /// the email ID suffix and/or base filename as necessary.
        /// </summary>
        /// <param name="emailPath">Absolute path of the markdown file to move.</param>
        /// <param name="threadFolderPath">Destination folder for the thread.</param>
        /// <returns>The absolute path of the moved file in its new location.</returns>
        public string MoveToThreadFolder(string emailPath, string threadFolderPath)
        {
            string fileName = Path.GetFileName(emailPath);
            string newFileName = fileName;
            // Add email id as a temporary suffix to avoid overwriting
            string emailId = null;
            bool inFrontMatter = false;
            foreach (var line in File.ReadLines(emailPath))
            {
                if (line.Trim() == "---")
                {
                    if (!inFrontMatter) { inFrontMatter = true; continue; }
                    else break;
                }
                if (inFrontMatter && line.Trim().StartsWith("internetMessageId:", StringComparison.OrdinalIgnoreCase))
                {
                    emailId = line.Trim().Substring("internetMessageId:".Length).Trim().Trim('"');
                    break;
                }
                if (inFrontMatter && line.Trim().StartsWith("entryId:", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(emailId))
                {
                    emailId = line.Trim().Substring("entryId:".Length).Trim().Trim('"');
                }
            }
            if (!string.IsNullOrEmpty(emailId))
            {
                // Sanitize emailId for filename (alphanumeric only)
                string safeId = new string(emailId.Where(char.IsLetterOrDigit).ToArray());
                string nameNoExt = Path.GetFileNameWithoutExtension(fileName);
                string ext = Path.GetExtension(fileName);

                // Windows path limit is 260 characters. Reserve 20 chars for safety margin.
                // Calculate how much space we have for the filename
                int maxPathLength = 240;
                int folderPathLength = threadFolderPath.Length + 1; // +1 for directory separator
                int availableForFilename = maxPathLength - folderPathLength - ext.Length;

                // Truncate safeId if necessary
                string fullFilename = $"{nameNoExt}-eid{safeId}";
                if (fullFilename.Length > availableForFilename)
                {
                    // Truncate the email ID, keeping at least the first 8 characters for uniqueness
                    int maxIdLength = availableForFilename - nameNoExt.Length - 4; // -4 for "-eid"
                    if (maxIdLength < 8)
                    {
                        // If even 8 chars won't fit, truncate the base name too
                        int maxBaseLength = availableForFilename - 12; // -12 for "-eid" + 8 char ID
                        if (maxBaseLength > 0)
                        {
                            nameNoExt = nameNoExt.Substring(0, Math.Min(nameNoExt.Length, maxBaseLength));
                        }
                        maxIdLength = 8;
                    }
                    safeId = safeId.Substring(0, Math.Min(safeId.Length, maxIdLength));
                }

                newFileName = $"{nameNoExt}-eid{safeId}{ext}";
            }

            string threadPath = Path.Combine(threadFolderPath, newFileName);

            // Final validation: ensure path doesn't exceed limit
            if (threadPath.Length > 240)
            {
                // Emergency truncation - this should rarely happen
                string nameNoExt = Path.GetFileNameWithoutExtension(newFileName);
                string ext = Path.GetExtension(newFileName);
                int maxFilenameLength = 240 - threadFolderPath.Length - ext.Length - 1;
                if (maxFilenameLength > 10)
                {
                    nameNoExt = nameNoExt.Substring(0, maxFilenameLength);
                    newFileName = $"{nameNoExt}{ext}";
                    threadPath = Path.Combine(threadFolderPath, newFileName);
                }
            }

            _fileService.EnsureDirectoryExists(threadFolderPath);
            if (File.Exists(threadPath))
            {
                File.Delete(threadPath);
            }
            File.Move(emailPath, threadPath);
            return threadPath;
        }

        /// <summary>
        /// Scans the inbox (and any existing thread folders) for notes that belong to the supplied conversation id.  
        /// Returns information that helps the caller decide whether the current email should be grouped into 
        /// an existing thread.
        /// </summary>
        /// <param name="conversationId">Identifier created by <see cref="GetConversationId"/>.</param>
        /// <param name="inboxPath">Absolute path to the vault inbox folder.</param>
        /// <returns>Tuple containing: <c>hasExistingThread</c>, the original thread name (if found), 
        /// the earliest email date in the thread and a count of matching messages.</returns>
        public (bool hasExistingThread, string earliestEmailThreadName, DateTime? earliestEmailDate, int emailCount) 
            FindExistingThread(string conversationId, string inboxPath)
        {
            bool hasExistingThread = false;
            DateTime? earliestEmailDate = null;
            string earliestEmailThreadName = null;
            int emailCount = 0;
            List<string> matchingFiles = new List<string>(); // Track matching files for debugging

            try 
            {
                // Get all markdown files from the inbox and subfolders
                var files = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
                
                // Search through each file for the thread ID
                foreach (var file in files)
                {
                    try
                    {
                        string emailContent = File.ReadAllText(file);
                        var threadIdMatch = Regex.Match(emailContent, @"threadId: ""([^""]+)""");
                        
                        // If this file belongs to the conversation thread
                        if (threadIdMatch.Success && threadIdMatch.Groups[1].Value == conversationId)
                        {
                            hasExistingThread = true;
                            emailCount++; // Increment email count for each matching email
                            matchingFiles.Add(file); // Add to our debugging list
                            
                            // Parse the date to find the earliest email
                            var dateMatch = Regex.Match(emailContent, @"date: (\d{4}-\d{2}-\d{2} \d{2}:\d{2})");
                            if (dateMatch.Success)
                            {
                                DateTime emailDate;
                                if (DateTime.TryParseExact(dateMatch.Groups[1].Value, "yyyy-MM-dd HH:mm", null, 
                                                         System.Globalization.DateTimeStyles.None, out emailDate))
                                {
                                    if (!earliestEmailDate.HasValue || emailDate < earliestEmailDate.Value)
                                    {
                                        earliestEmailDate = emailDate;
                                        
                                        // Check if this email is in a thread folder
                                        string directory = Path.GetDirectoryName(file);
                                        if (directory != inboxPath)
                                        {
                                            earliestEmailThreadName = Path.GetFileName(directory);
                                        }
                                        else
                                        {
                                            // Try to extract thread name components from frontmatter
                                            var subjectMatch = Regex.Match(emailContent, @"title: ""([^""]+)""");
                                            var fromMatch = Regex.Match(emailContent, @"from: ""[^""]*\[\[([^""]+)\]\]""");
                                            var toMatch = Regex.Match(emailContent, @"to:.*?\n\s*- ""[^""]*\[\[([^""]+)\]\]""", RegexOptions.Singleline);
                                            
                                            if (subjectMatch.Success && fromMatch.Success && toMatch.Success)
                                            {
                                                string subject = _fileService.CleanFileName(subjectMatch.Groups[1].Value);
                                                if (subject.Length > 50)
                                                {
                                                    subject = subject.Substring(0, 47) + "...";
                                                }
                                                string sender = fromMatch.Groups[1].Value;
                                                string recipient = toMatch.Groups[1].Value;
                                                earliestEmailThreadName = $"{subject}-{sender}-{recipient}";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (System.Exception)
                    {
                        // Skip files that can't be read
                        continue;
                    }
                }
                
                // If we found any matches, show more details for debugging
                if (emailCount > 0 && _settings.ShowThreadDebug)
                {
                    string filesList = string.Join("\n", matchingFiles);
                    System.Windows.Forms.MessageBox.Show(
                        $"Found {emailCount} existing emails with thread ID: {conversationId}\n\nFiles:\n{filesList}",
                        "Thread Match Details",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error searching for thread: {ex.Message}", "Thread Search Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }

            return (hasExistingThread, earliestEmailThreadName, earliestEmailDate, emailCount);
        }

        /// <summary>
        /// Renames every email note in a thread folder so that they receive a chronological <c>yyyy-MM-dd_HHmmss</c>
        /// prefix.  This guarantees a predictable sort order inside Obsidian and prevents filename collisions.
        /// Any temporary <c>-eid*</c> suffixes or outdated numeric suffixes are stripped before each file's
        /// actual base name is preserved. If multiple emails have identical timestamps, a counter suffix (_001, _002, etc.) is added.
        /// </summary>
        /// <param name="threadFolderPath">Full path to the thread folder.</param>
        /// <param name="baseName">Deprecated parameter (no longer used - kept for backward compatibility).</param>
        /// <param name="currentEmailPath">If supplied, returns the new path for the file that corresponds to the current email.</param>
        /// <returns>The new path for <c>currentEmailPath</c> if it was provided; otherwise <c>null</c>.</returns>
        public string ResuffixThreadNotes(string threadFolderPath, string baseName, string currentEmailPath = null)
        {
            if (!Directory.Exists(threadFolderPath)) return null;

            // Get all markdown files in the thread folder except the thread summary (starts with "0-")
            var files = Directory.GetFiles(threadFolderPath, "*.md", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("0-"))
                .ToList();
            var fileDates = new List<(string file, DateTime date)>();
            foreach (var file in files)
            {
                DateTime? date = null;
                bool inFrontMatter = false;
                foreach (var line in File.ReadLines(file))
                {
                    if (line.Trim() == "---")
                    {
                        if (!inFrontMatter) { inFrontMatter = true; continue; }
                        else break;
                    }
                    if (inFrontMatter && line.Trim().StartsWith("date:", StringComparison.OrdinalIgnoreCase))
                    {
                        var value = line.Trim().Substring("date:".Length).Trim().Trim('"');
                        if (DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out var parsed))
                            date = parsed;
                        else if (DateTime.TryParse(value, out var fallback))
                            date = fallback;
                        break;
                    }
                }
                if (date.HasValue)
                {
                    fileDates.Add((file, date.Value));
                }
            }
            // Order oldest first so lexicographical date prefix sorts correctly
            fileDates = fileDates.OrderBy(fd => fd.date).ToList();

            string newCurrentPath = null;
            int counter = 0;
            string lastGeneratedName = null;

            foreach (var fd in fileDates)
            {
                // Build new filename with date prefix (yyyy-MM-dd_HHmmss) followed by actual base name from file
                string nameNoExt = Path.GetFileNameWithoutExtension(fd.file);
                string ext = Path.GetExtension(fd.file);

                // Remove any email id suffix ( -eid{ID} ) or old numeric suffix ( -001 )
                nameNoExt = Regex.Replace(nameNoExt, "-eid[0-9A-Za-z]+$", "");
                nameNoExt = Regex.Replace(nameNoExt, "-\\d{3}$", "");

                // Remove any existing date prefix formats (yyyy-MM-dd_HHmmss, yyyy-MM-dd-HHmm, yyyy-MM-dd_HHmm)
                nameNoExt = Regex.Replace(nameNoExt, "^\\d{4}-\\d{2}-\\d{2}[_-]\\d{4,6}_?", "");

                // Include seconds in date prefix to reduce collision probability
                string datePrefix = fd.date.ToString("yyyy-MM-dd_HHmmss");
                string newName = $"{datePrefix}_{nameNoExt}{ext}";
                string newPath = Path.Combine(threadFolderPath, newName);

                // If we have a collision with the last generated name, add a counter suffix
                if (newName == lastGeneratedName)
                {
                    counter++;
                    newName = $"{datePrefix}_{nameNoExt}_{counter:D3}{ext}";
                    newPath = Path.Combine(threadFolderPath, newName);
                }
                else
                {
                    counter = 0;
                    lastGeneratedName = newName;
                }

                if (!fd.file.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                {
                    // Ensure no conflict before moving
                    if (File.Exists(newPath))
                    {
                        // If target exists and it's not the same file, add counter
                        counter++;
                        newName = $"{datePrefix}_{nameNoExt}_{counter:D3}{ext}";
                        newPath = Path.Combine(threadFolderPath, newName);
                    }
                    File.Move(fd.file, newPath);
                }

                if (currentEmailPath != null && fd.file.Equals(currentEmailPath, StringComparison.OrdinalIgnoreCase))
                {
                    newCurrentPath = newPath;
                }
            }
            return newCurrentPath;
        }
    }
} 