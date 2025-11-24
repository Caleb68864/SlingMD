using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;
using System.Linq;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Helpers;
using Logger = SlingMD.Outlook.Helpers.Logger;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Orchestrates the full life-cycle of turning an <see cref="MailItem"/> into a properly formatted
    /// markdown note inside the user's Obsidian vault. The processor coordinates the various helper
    /// services (file-, thread-, task- and contact-services) and honours all user settings.
    /// </summary>
    public class EmailProcessor
    {
        // Static lock dictionary to prevent race conditions when multiple emails target the same thread folder
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Threading.SemaphoreSlim> _threadFolderLocks
            = new System.Collections.Concurrent.ConcurrentDictionary<string, System.Threading.SemaphoreSlim>(StringComparer.OrdinalIgnoreCase);

        // Static cache to track processed email IDs and prevent O(n*m) file scanning on every email
        // Key: email ID (internetMessageId or entryId), Value: true (exists)
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, bool> _processedEmailIds
            = new System.Collections.Concurrent.ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private static DateTime _cacheLastBuilt = DateTime.MinValue;
        private static readonly object _cacheBuildLock = new object();

        // Compiled regex patterns for performance (used in CleanSubject and title formatting)
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);
        private static readonly Regex TrailingDashSpaceRegex = new Regex(@"[-\s]+$", RegexOptions.Compiled);
        private static readonly Regex ColonSpaceRegex = new Regex(@":\s*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex1 = new Regex(@"(?:Re_\s*)+(?:RE_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex2 = new Regex(@"(?:RE_\s*)+(?:Re_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex3 = new Regex(@"(?:Re_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex4 = new Regex(@"(?:RE_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex1 = new Regex(@"(?:Fw_\s*)+(?:FW_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex2 = new Regex(@"(?:FW_\s*)+(?:Fw_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex3 = new Regex(@"(?:Fw_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex4 = new Regex(@"(?:FW_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplySpaceRegex = new Regex(@"Re_\s+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardSpaceRegex = new Regex(@"Fw_\s+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;
        private readonly TaskService _taskService;
        private readonly ContactService _contactService;
        private readonly AttachmentService _attachmentService;

        public EmailProcessor(ObsidianSettings settings)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService, settings);
            _taskService = new TaskService(settings);
            _contactService = new ContactService(_fileService, _templateService);
            _attachmentService = new AttachmentService(settings, _fileService);
        }

        /// <summary>
        /// Converts the supplied <paramref name="mail"/> into a markdown note, creates the optional follow-up
        /// tasks and opens the resulting file in Obsidian (depending on settings). The method is asynchronous
        /// because it performs a number of I/O heavy operations (file moves, Outlook task creation, countdown
        /// dialog) that would otherwise block the Outlook UI thread.
        /// </summary>
        /// <param name="mail">The email that should be exported.</param>
        /// <param name="cancellationToken">Optional cancellation token to allow cancellation of long-running operations.</param>
        public async Task ProcessEmail(MailItem mail, System.Threading.CancellationToken cancellationToken = default(System.Threading.CancellationToken))
        {
            // Declare variables at method level so they're accessible throughout the method
            List<string> contactNames = new List<string>();
            string fileName = string.Empty;
            string fileNameNoExt = string.Empty;
            string filePath = string.Empty;
            string obsidianLinkPath = string.Empty;  // Added to store the path to use for Obsidian
            string conversationId = string.Empty;
            string threadNoteName = string.Empty;
            string threadFolderPath = string.Empty;
            string threadNotePath = string.Empty;
            bool shouldGroupThread = false;

            // Get task options first if needed
            if ((_settings.CreateOutlookTask || _settings.CreateObsidianTask) && _settings.AskForDates)
            {
                using (var form = new TaskOptionsForm(_settings.DefaultDueDays, _settings.DefaultReminderDays, _settings.DefaultReminderHour, _settings.UseRelativeReminder))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        _taskService.InitializeTaskSettings(form.DueDays, form.ReminderDays, form.ReminderHour, form.UseRelativeReminder);
                    }
                    else
                    {
                        _taskService.DisableTaskCreation();
                    }
                }
            }
            else
            {
                _taskService.InitializeTaskSettings();
            }

            // Collect all contact names - will be used later for contact creation
            // Add null check for SenderName
            if (!string.IsNullOrEmpty(mail.SenderName))
            {
                contactNames.Add(mail.SenderName);
            }

            // Properly release COM objects to prevent memory leaks
            Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        contactNames.Add(recipient.Name);
                    }
                    finally
                    {
                        // Release each Recipient COM object
                        if (recipient != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                        }
                    }
                }
            }
            finally
            {
                // Release the Recipients collection COM object
                if (recipients != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recipients);
                }
            }

            using (var status = new StatusService())
            {
                try
                {
                    status.UpdateProgress("Processing email...", 0);

                    // Check for cancellation
                    cancellationToken.ThrowIfCancellationRequested();

                    // Build note title using settings (with null safety)
                    string noteTitle = mail.Subject ?? "No Subject";
                    string senderClean = _contactService.GetShortName(mail.SenderName ?? "Unknown Sender");
                    string fileDateTime = mail.ReceivedTime.ToString("yyyy-MM-dd-HHmm");
                    string dateStr = mail.ReceivedTime.ToString("yyyy-MM-dd");
                    string subjectClean = CleanSubject(mail.Subject ?? "No Subject");

                    // Use settings for title format
                    string titleFormat = _settings.NoteTitleFormat ?? "{Subject} - {Date}";
                    bool includeDate = _settings.NoteTitleIncludeDate;
                    int maxLength = _settings.NoteTitleMaxLength > 0 ? _settings.NoteTitleMaxLength : 50;

                    // Prepare replacements
                    string formattedTitle = titleFormat
                        .Replace("{Subject}", subjectClean)
                        .Replace("{Sender}", senderClean)
                        .Replace("{Date}", includeDate ? dateStr : "");
                    // Remove double spaces and trim (using compiled regex)
                    formattedTitle = WhitespaceRegex.Replace(formattedTitle, " ").Trim();
                    // Remove trailing dash if date is omitted (using compiled regex)
                    formattedTitle = TrailingDashSpaceRegex.Replace(formattedTitle, "").Trim();
                    // Trim to max length
                    if (formattedTitle.Length > maxLength)
                        formattedTitle = formattedTitle.Substring(0, maxLength - 3) + "...";
                    noteTitle = formattedTitle;

                    // Email threading logic moved to its own method
                    (conversationId, threadNoteName, threadFolderPath, threadNotePath, shouldGroupThread, obsidianLinkPath, fileName, filePath, fileNameNoExt) =
                        GetThreadingInfo(mail, subjectClean, senderClean, fileDateTime, "");
                    if (shouldGroupThread)
                    {
                        status.UpdateProgress($"Email thread found: {threadNoteName}", 48);
                        // Remove 0- prefix from threadFolderPath if present
                        if (threadFolderPath.Contains($"0-{threadNoteName}"))
                        {
                            threadFolderPath = threadFolderPath.Replace($"0-{threadNoteName}", threadNoteName);
                            threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
                        }
                    }

                    status.UpdateProgress("Processing email metadata", 50);

                    // Extract real email IDs
                    var (realInternetMessageId, realEntryId) = ExtractEmailUniqueIds(mail);

                    // Get Recipients collection once and release it properly to prevent COM leaks
                    List<string> toLinked;
                    List<string> toEmails;
                    List<string> ccLinked;
                    List<string> ccEmails;

                    Recipients mailRecipients = null;
                    try
                    {
                        mailRecipients = mail.Recipients;
                        toLinked = _contactService.BuildLinkedNames(mailRecipients, OlMailRecipientType.olTo);
                        toEmails = _contactService.BuildEmailList(mailRecipients, OlMailRecipientType.olTo);
                        ccLinked = _contactService.BuildLinkedNames(mailRecipients, OlMailRecipientType.olCC);
                        ccEmails = _contactService.BuildEmailList(mailRecipients, OlMailRecipientType.olCC);
                    }
                    finally
                    {
                        // Release the Recipients collection COM object
                        if (mailRecipients != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mailRecipients);
                        }
                    }

                    // Build metadata for frontmatter
                    var metadata = new Dictionary<string, object>
                    {
                        { "title", noteTitle },
                        { "from", $"[[{mail.SenderName ?? "Unknown Sender"}]]" },
                        { "fromEmail", _contactService.GetSenderEmail(mail) },
                        { "to", toLinked },
                        { "toEmail", toEmails },
                        { "threadId", conversationId },
                        { "date", mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss") },
                        { "dailyNoteLink", $"[[{mail.ReceivedTime:yyyy-MM-dd}]]" },
                        { "internetMessageId", realInternetMessageId },
                        { "entryId", realEntryId },
                        { "tags", (_settings.DefaultNoteTags != null && _settings.DefaultNoteTags.Count > 0) ? new List<string>(_settings.DefaultNoteTags) : new List<string> { "FollowUp" } }
                    };

                    // Add CC if present
                    if (ccEmails.Count > 0)
                    {
                        metadata.Add("cc", ccLinked);
                        metadata.Add("ccEmail", ccEmails);
                    }

                    // Add threadNote if this is part of a thread and thread grouping is enabled
                    if (shouldGroupThread)
                    {
                        metadata.Add("threadNote", $"[[0-{threadNoteName}]]");
                    }

                    // Build content
                    var content = new System.Text.StringBuilder();
                    content.Append(_templateService.BuildFrontMatter(metadata));

                    // Add Obsidian task if enabled, using DefaultTaskTags
                    if (_settings.CreateObsidianTask && _taskService.ShouldCreateTasks)
                    {
                        var taskTags = (_settings.DefaultTaskTags != null && _settings.DefaultTaskTags.Count > 0)
                            ? _settings.DefaultTaskTags
                            : new List<string> { "FollowUp" };
                        content.Append(_taskService.GenerateObsidianTask(fileNameNoExt, taskTags));
                        content.Append("\n\n");
                    }

                    content.Append(mail.Body ?? string.Empty);

                    status.UpdateProgress("Writing note file", 75);

                    // Check for cancellation before file operations
                    cancellationToken.ThrowIfCancellationRequested();

                    // Check for duplicate email before writing the note
                    if (IsDuplicateEmail(_settings.GetInboxPath(), realInternetMessageId, realEntryId))
                    {
                        status.UpdateProgress("Duplicate email detected. Skipping note creation.", 100);
                        return;
                    }

                    if (shouldGroupThread)
                    {
                        // Get or create a semaphore for this thread folder to prevent race conditions
                        var semaphore = _threadFolderLocks.GetOrAdd(threadFolderPath, _ => new System.Threading.SemaphoreSlim(1, 1));

                        // Acquire the lock for this thread folder
                        await semaphore.WaitAsync();
                        try
                        {
                            // Write the new note for the current email to the thread folder with -eid{id} suffix
                            string emailId = !string.IsNullOrEmpty(realInternetMessageId) ? realInternetMessageId : realEntryId;
                            string safeId = new string(emailId.Where(char.IsLetterOrDigit).ToArray());
                            string baseName = $"{subjectClean}-{senderClean}";
                            string tempFileName = $"{baseName}-eid{safeId}.md";
                            string tempFilePath = Path.Combine(threadFolderPath, tempFileName);
                            _fileService.EnsureDirectoryExists(threadFolderPath);
                            _fileService.WriteUtf8File(tempFilePath, content.ToString());

                            // Add to cache to keep it synchronized
                            AddEmailToCache(realInternetMessageId, realEntryId);

                            filePath = tempFilePath;
                            fileName = tempFileName;
                            fileNameNoExt = Path.GetFileNameWithoutExtension(tempFileName);

                            // Move all existing emails for this thread to the thread folder
                            var mdFiles = Directory.GetFiles(_settings.GetInboxPath(), "*.md", SearchOption.TopDirectoryOnly);
                            foreach (var file in mdFiles)
                            {
                                // Read front matter to get threadId
                                bool inFrontMatter = false;
                                string foundThreadId = null;
                                foreach (var line in File.ReadLines(file))
                                {
                                    if (line.Trim() == "---")
                                    {
                                        if (!inFrontMatter)
                                        {
                                            inFrontMatter = true;
                                            continue;
                                        }
                                        else
                                        {
                                            // End of frontmatter
                                            break;
                                        }
                                    }
                                    if (inFrontMatter && line.Trim().StartsWith("threadId:", StringComparison.OrdinalIgnoreCase))
                                    {
                                        foundThreadId = line.Trim().Substring("threadId:".Length).Trim().Trim('"');
                                        break;
                                    }
                                }
                                if (!string.IsNullOrWhiteSpace(foundThreadId) && foundThreadId == conversationId)
                                {
                                    // Only move if not already in the thread folder
                                    if (Path.GetDirectoryName(file) != threadFolderPath)
                                    {
                                        _threadService.MoveToThreadFolder(file, threadFolderPath);
                                    }
                                }
                            }
                            // Resuffix all notes in the thread folder (except thread summary)
                            var updatedCurrentPath = _threadService.ResuffixThreadNotes(threadFolderPath, baseName, tempFilePath);

                            // Verify the file exists with retry logic
                            if (!string.IsNullOrWhiteSpace(updatedCurrentPath))
                            {
                                await WaitForFileAvailability(updatedCurrentPath);
                                filePath = updatedCurrentPath;
                                fileName = Path.GetFileName(updatedCurrentPath);
                                fileNameNoExt = Path.GetFileNameWithoutExtension(updatedCurrentPath);
                                obsidianLinkPath = $"{threadNoteName}/{fileNameNoExt}";
                            }

                            // Create or update the thread summary note
                            await _threadService.UpdateThreadNote(threadFolderPath, threadNotePath, conversationId, threadNoteName, mail);

                            // Process attachments if enabled (inside semaphore to prevent race condition with file resuffixing)
                            if (_settings.SaveInlineImages || _settings.SaveAllAttachments)
                            {
                                try
                                {
                                    status.UpdateProgress("Processing attachments", 77);
                                    var attachmentInfo = _attachmentService.ProcessAttachments(mail, filePath);

                                    if (attachmentInfo.SavedAttachments.Count > 0)
                                    {
                                        // Build attachment section
                                        var attachmentSection = new StringBuilder();
                                        attachmentSection.AppendLine("\n\n## Attachments\n");

                                        foreach (var attachment in attachmentInfo.SavedAttachments)
                                        {
                                            string wikilink = _attachmentService.GenerateWikilink(
                                                attachment.Filename,
                                                attachment.IsInline
                                            );
                                            attachmentSection.AppendLine(wikilink);
                                        }

                                        // Append to the existing note file
                                        _fileService.AppendToFile(filePath, attachmentSection.ToString());
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    Logger.Instance.Error($"Failed to process attachments for threaded email: {ex.Message}", ex);
                                    // Continue processing - don't fail the entire email if attachments fail
                                }
                            }
                        }
                        finally
                        {
                            // Always release the semaphore
                            semaphore.Release();
                        }
                    }
                    else
                    {
                        // Write the note as usual to the inbox
                        _fileService.WriteUtf8File(filePath, content.ToString());

                        // Add to cache to keep it synchronized
                        AddEmailToCache(realInternetMessageId, realEntryId);

                        // Process attachments if enabled
                        if (_settings.SaveInlineImages || _settings.SaveAllAttachments)
                        {
                            try
                            {
                                status.UpdateProgress("Processing attachments", 77);
                                var attachmentInfo = _attachmentService.ProcessAttachments(mail, filePath);

                                if (attachmentInfo.SavedAttachments.Count > 0)
                                {
                                    // Build attachment section
                                    var attachmentSection = new StringBuilder();
                                    attachmentSection.AppendLine("\n\n## Attachments\n");

                                    foreach (var attachment in attachmentInfo.SavedAttachments)
                                    {
                                        string wikilink = _attachmentService.GenerateWikilink(
                                            attachment.Filename,
                                            attachment.IsInline
                                        );
                                        attachmentSection.AppendLine(wikilink);
                                    }

                                    // Append to the existing note file
                                    _fileService.AppendToFile(filePath, attachmentSection.ToString());
                                }
                            }
                            catch (System.Exception ex)
                            {
                                Logger.Instance.Error($"Failed to process attachments: {ex.Message}", ex);
                                // Continue processing - don't fail the entire email if attachments fail
                            }
                        }
                    }

                    // Create Outlook task if enabled
                    if (_settings.CreateOutlookTask && _taskService.ShouldCreateTasks)
                    {
                        status.UpdateProgress("Creating Outlook task", 80);
                        await _taskService.CreateOutlookTask(mail);
                    }
                    
                    status.UpdateProgress("Completing email processing", 100);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error processing email: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
            // Process contacts outside the StatusService block
            // This ensures the progress window doesn't block the contact dialog
            if (_settings.EnableContactSaving && contactNames.Count > 0)
            {
                try
                {
                    // Remove duplicates and sort
                    contactNames = contactNames.Distinct().OrderBy(n => n).ToList();
                    
                    // Filter to only new contacts
                    var newContacts = new List<string>();
                    foreach (var contactName in contactNames)
                    {
                        if (!_contactService.ContactExists(contactName))
                        {
                            newContacts.Add(contactName);
                        }
                    }
                    
                    // Only show dialog if we have new contacts to create
                    if (newContacts.Count > 0)
                    {
                        // Show contact confirmation dialog
                        using (var dialog = new ContactConfirmationDialog(newContacts))
                        {
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                foreach (var contactName in dialog.SelectedContacts)
                                {
                                    // Create contact note for each selected contact
                                    _contactService.CreateContactNote(contactName);
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error processing contacts: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
            // Launch Obsidian if enabled
            if (_settings.LaunchObsidian)
            {
                try
                {
                    // Ensure delay happens after all file operations
                    if (_settings.ShowCountdown && _settings.ObsidianDelaySeconds > 0)
                    {
                        using (var countdown = new CountdownForm(_settings.ObsidianDelaySeconds))
                        {
                            countdown.ShowDialog();
                        }
                    }
                    else if (_settings.ObsidianDelaySeconds > 0)
                    {
                        await Task.Delay(_settings.ObsidianDelaySeconds * 1000);
                    }

                    // Always use the latest obsidianLinkPath (updated after resuffixing)
                    _fileService.LaunchObsidian(_settings.VaultName, obsidianLinkPath);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error launching Obsidian: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string CleanSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
                return string.Empty;

            string cleaned = subject;

            // Apply all cleanup patterns from settings
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                cleaned = Regex.Replace(cleaned, pattern, "", RegexOptions.IgnoreCase);
            }
            
            // Replace colons (with or without spaces) with underscores (using compiled regex)
            cleaned = ColonSpaceRegex.Replace(cleaned, "_");

            // Handle Re_ (Reply) prefixes (using compiled regex)
            // Remove redundant Re_ RE_ prefixes - keep only one "Re_"
            cleaned = ReplyPrefixRegex1.Replace(cleaned, "Re_");
            cleaned = ReplyPrefixRegex2.Replace(cleaned, "Re_");
            cleaned = ReplyPrefixRegex3.Replace(cleaned, "Re_");
            cleaned = ReplyPrefixRegex4.Replace(cleaned, "Re_");

            // Handle Fw_ (Forward) prefixes (using compiled regex)
            // Remove redundant Fw_ FW_ prefixes - keep only one "Fw_"
            cleaned = ForwardPrefixRegex1.Replace(cleaned, "Fw_");
            cleaned = ForwardPrefixRegex2.Replace(cleaned, "Fw_");
            cleaned = ForwardPrefixRegex3.Replace(cleaned, "Fw_");
            cleaned = ForwardPrefixRegex4.Replace(cleaned, "Fw_");

            // Ensure there are no spaces after prefixes (using compiled regex)
            cleaned = ReplySpaceRegex.Replace(cleaned, "Re_");
            cleaned = ForwardSpaceRegex.Replace(cleaned, "Fw_");

            return _fileService.CleanFileName(cleaned.Trim());
        }

        private string GetFirstRecipient(MailItem mail)
        {
            Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        if (recipient.Type == (int)OlMailRecipientType.olTo)
                        {
                            string recipientName = recipient.Name;
                            return recipientName;
                        }
                    }
                    finally
                    {
                        // Release each Recipient COM object
                        if (recipient != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                        }
                    }
                }
            }
            finally
            {
                // Release the Recipients collection COM object
                if (recipients != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recipients);
                }
            }
            return "Unknown";
        }

        /// <summary>
        /// Waits for a file to become available by checking its existence and attempting to open it.
        /// Uses exponential backoff retry logic with a maximum of 5 attempts over approximately 1 second.
        /// This replaces hard-coded delays and ensures file system operations have completed.
        /// </summary>
        /// <param name="filePath">The full path to the file to wait for.</param>
        private async Task WaitForFileAvailability(string filePath)
        {
            const int maxAttempts = 5;
            int attempt = 0;
            int delayMs = 50; // Start with 50ms

            while (attempt < maxAttempts)
            {
                try
                {
                    // Check if file exists and can be opened
                    if (File.Exists(filePath))
                    {
                        // Try to open the file briefly to ensure it's not locked
                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            // File is accessible
                            return;
                        }
                    }
                }
                catch (IOException)
                {
                    // File is locked or not yet available
                }
                catch (UnauthorizedAccessException)
                {
                    // Permission issue, but file exists
                    return;
                }

                // Exponential backoff: 50ms, 100ms, 200ms, 400ms, 800ms
                attempt++;
                if (attempt < maxAttempts)
                {
                    await Task.Delay(delayMs);
                    delayMs *= 2;
                }
            }

            // If we get here, file still isn't available but we've tried our best
            // Continue anyway - the worst case is the same as before
        }

        /// <summary>
        /// Extracts the InternetMessageID and EntryID from a MailItem.
        /// Returns (internetMessageId, entryId).
        /// </summary>
        private (string internetMessageId, string entryId) ExtractEmailUniqueIds(MailItem mail)
        {
            string entryId = mail.EntryID;
            string internetMessageId = null;
            try
            {
                // Try to get InternetMessageID via PropertyAccessor (works for most accounts)
                internetMessageId = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E") as string;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"PropertyAccessor method failed for InternetMessageID: {ex.Message}");
            }
            // Fallback to property if available
            if (string.IsNullOrEmpty(internetMessageId))
            {
                try
                {
                    internetMessageId = mail.GetType().GetProperty("InternetMessageID")?.GetValue(mail) as string;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"Reflection method failed for InternetMessageID: {ex.Message}");
                }
            }
            return (internetMessageId, entryId);
        }

        /// <summary>
        /// Returns thread-related info for an email, including paths and names.
        /// </summary>
        private (string conversationId, string threadNoteName, string threadFolderPath, string threadNotePath, bool shouldGroupThread, string obsidianLinkPath, string fileName, string filePath, string fileNameNoExt) GetThreadingInfo(MailItem mail, string subjectClean, string senderClean, string fileDateTime, string fileNameNoExt)
        {
            string conversationId = _threadService.GetConversationId(mail);
            string threadFolderName = _threadService.GetThreadFolderName(mail, subjectClean, senderClean, _contactService.GetShortName(GetFirstRecipient(mail)));
            string threadNoteName = _threadService.GetThreadNoteName(mail, subjectClean, senderClean, _contactService.GetShortName(GetFirstRecipient(mail)));
            string threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadFolderName);
            string threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
            var threadInfo = _threadService.FindExistingThread(conversationId, _settings.GetInboxPath());
            bool hasExistingThread = threadInfo.hasExistingThread;
            string earliestEmailThreadName = threadInfo.earliestEmailThreadName;
            int emailCount = threadInfo.emailCount;
            bool shouldGroupThread = hasExistingThread && _settings.GroupEmailThreads && emailCount >= 1;
            string fileName, filePath, obsidianLinkPath, fileNameNoExtResult;
            bool includeDate = _settings.NoteTitleIncludeDate;
            if (shouldGroupThread)
            {
                threadNoteName = earliestEmailThreadName ?? threadFolderName;
                threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
                if (includeDate)
                {
                    if (_settings.MoveDateToFrontInThread)
                    {
                        fileName = $"{fileDateTime}-{subjectClean}-{senderClean}.md";
                    }
                    else
                    {
                        fileName = $"{subjectClean}-{senderClean}-{fileDateTime}.md";
                    }
                }
                else
                {
                    // Suffix logic: gather all files, parse dates, sort, assign suffixes
                    string baseName = $"{subjectClean}-{senderClean}";
                    var files = Directory.Exists(threadFolderPath)
                        ? Directory.GetFiles(threadFolderPath, baseName + "*.md", SearchOption.TopDirectoryOnly)
                        : new string[0];
                    // List of (filename, date) pairs
                    var fileDates = new List<(string file, DateTime date, bool isCurrent)>();
                    DateTime thisDate = mail.ReceivedTime;
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
                            fileDates.Add((file, date.Value, false));
                        }
                    }
                    // Add the current email
                    fileDates.Add((null, thisDate, true));
                    // Sort by date
                    fileDates = fileDates.OrderBy(fd => fd.date).ToList();
                    // Assign suffixes
                    int idx = 1;
                    string suffix = null;
                    foreach (var fd in fileDates)
                    {
                        if (fd.isCurrent)
                        {
                            suffix = $"-{idx:D3}";
                            break;
                        }
                        idx++;
                    }
                    fileName = $"{baseName}{suffix}.md";
                }
                filePath = Path.Combine(threadFolderPath, fileName);
                fileNameNoExtResult = Path.GetFileNameWithoutExtension(fileName);
                obsidianLinkPath = $"{threadNoteName}/{fileNameNoExtResult}";
            }
            else
            {
                if (includeDate)
                {
                    fileName = $"{subjectClean}-{senderClean}-{fileDateTime}.md";
                }
                else
                {
                    fileName = $"{subjectClean}-{senderClean}.md";
                }
                filePath = Path.Combine(_settings.GetInboxPath(), fileName);
                fileNameNoExtResult = Path.GetFileNameWithoutExtension(fileName);
                obsidianLinkPath = fileNameNoExtResult;
            }
            return (conversationId, threadNoteName, threadFolderPath, threadNotePath, shouldGroupThread, obsidianLinkPath, fileName, filePath, fileNameNoExtResult);
        }

        /// <summary>
        /// Builds or rebuilds the email ID cache by scanning all markdown files in the inbox.
        /// Cache is rebuilt if it's older than 5 minutes or empty.
        /// This dramatically improves performance by avoiding O(n*m) file scanning on every email.
        /// </summary>
        private void EnsureEmailCacheIsBuilt(string inboxPath)
        {
            // Only rebuild cache if it's older than 5 minutes (to allow for external changes)
            // or if it's never been built
            if ((DateTime.Now - _cacheLastBuilt).TotalMinutes < 5 && _processedEmailIds.Count > 0)
            {
                return;
            }

            lock (_cacheBuildLock)
            {
                // Double-check after acquiring lock
                if ((DateTime.Now - _cacheLastBuilt).TotalMinutes < 5 && _processedEmailIds.Count > 0)
                {
                    return;
                }

                _processedEmailIds.Clear();

                var mdFiles = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
                foreach (var file in mdFiles)
                {
                    bool inFrontMatter = false;
                    foreach (var line in File.ReadLines(file))
                    {
                        if (line.Trim() == "---")
                        {
                            if (!inFrontMatter)
                            {
                                inFrontMatter = true;
                                continue;
                            }
                            else
                            {
                                // End of frontmatter
                                break;
                            }
                        }
                        if (inFrontMatter)
                        {
                            var trimmed = line.Trim();
                            if (trimmed.StartsWith("internetMessageId:", StringComparison.OrdinalIgnoreCase))
                            {
                                var value = trimmed.Substring("internetMessageId:".Length).Trim().Trim('"');
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    _processedEmailIds.TryAdd(value, true);
                                }
                            }
                            else if (trimmed.StartsWith("entryId:", StringComparison.OrdinalIgnoreCase))
                            {
                                var value = trimmed.Substring("entryId:".Length).Trim().Trim('"');
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    _processedEmailIds.TryAdd(value, true);
                                }
                            }
                        }
                    }
                }

                _cacheLastBuilt = DateTime.Now;
            }
        }

        /// <summary>
        /// Checks if an email with the given InternetMessageID or EntryID already exists.
        /// Uses an in-memory cache for O(1) lookups instead of scanning all files.
        /// Cache is automatically rebuilt every 5 minutes to stay synchronized with external changes.
        /// </summary>
        private bool IsDuplicateEmail(string inboxPath, string internetMessageId, string entryId)
        {
            // Ensure cache is built and up-to-date
            EnsureEmailCacheIsBuilt(inboxPath);

            // Fast O(1) lookup in cache
            if (!string.IsNullOrWhiteSpace(internetMessageId) && _processedEmailIds.ContainsKey(internetMessageId))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(entryId) && _processedEmailIds.ContainsKey(entryId))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Adds email IDs to the cache after successfully saving an email note.
        /// This keeps the cache synchronized without requiring a full rebuild.
        /// </summary>
        private void AddEmailToCache(string internetMessageId, string entryId)
        {
            if (!string.IsNullOrWhiteSpace(internetMessageId))
            {
                _processedEmailIds.TryAdd(internetMessageId, true);
            }
            if (!string.IsNullOrWhiteSpace(entryId))
            {
                _processedEmailIds.TryAdd(entryId, true);
            }
        }
    }
}