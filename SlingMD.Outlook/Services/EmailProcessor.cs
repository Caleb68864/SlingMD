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
using SlingMD.Outlook.Helpers;
using Logger = SlingMD.Outlook.Helpers.Logger;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Orchestrates the full life-cycle of turning an <see cref="MailItem"/> into a properly formatted
    /// markdown note inside the user's Obsidian vault. The processor coordinates the various helper
    /// services (file-, thread-, task- and contact-services) and honours all user settings.
    /// </summary>
    public class EmailProcessor
    {
        // Exponential-backoff parameters for WaitForFileAvailability.
        // 5 attempts × starting at 50 ms doubles to ~1550 ms total wait worst-case.
        private const int MaxFileWaitAttempts = 5;
        private const int InitialFileWaitDelayMs = 50;

        // Static lock dictionary to prevent race conditions when multiple emails target the same thread folder
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Threading.SemaphoreSlim> _threadFolderLocks
            = new System.Collections.Concurrent.ConcurrentDictionary<string, System.Threading.SemaphoreSlim>(StringComparer.OrdinalIgnoreCase);

        // Ids of the emails already in the vault, shared by every EmailProcessor (see ProcessedEmailCache).
        private static readonly ProcessedEmailCache _processedEmails = new ProcessedEmailCache();

        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;
        private readonly TaskService _taskService;
        private readonly ContactService _contactService;
        private readonly AttachmentService _attachmentService;
        private readonly EmailNoteEngine _engine;
        private readonly SubjectFilenameCleaner _subjectFilenameCleaner;
        private readonly UniqueFilenameResolver _uniqueFilenameResolver = new UniqueFilenameResolver();
        private readonly IClock _clock;

        /// <summary>
        /// Whether any attachment processing should run for an email given the current settings.
        /// Kept as a single predicate so the two call sites in <see cref="ProcessEmail"/> can't
        /// drift apart — omitting <see cref="ObsidianSettings.SaveRealAttachments"/> here was the
        /// cause of issue #12, where real attachments were never saved under the default settings.
        /// Once processing is enabled, <see cref="AttachmentService.ShouldSave"/> makes the
        /// per-attachment inline/real decision.
        /// </summary>
        internal static bool ShouldProcessAttachments(ObsidianSettings settings)
        {
            return settings != null
                && (settings.SaveRealAttachments || settings.SaveInlineImages || settings.SaveAllAttachments);
        }

        public EmailProcessor(ObsidianSettings settings, IClock clock = null)
            : this(
                settings,
                clock ?? new SystemClock(),
                fileService: null,
                templateService: null,
                threadService: null,
                taskService: null,
                contactService: null,
                attachmentService: null)
        {
        }

        /// <summary>
        /// Full-injection constructor for tests. Any argument passed as <c>null</c> falls back
        /// to the default production wiring. Production callers should use the single-arg
        /// constructor — this overload is only intended for unit tests that need to substitute
        /// collaborators (e.g. a fake FileService or an in-memory TemplateService).
        /// </summary>
        internal EmailProcessor(
            ObsidianSettings settings,
            IClock clock,
            FileService fileService,
            TemplateService templateService,
            ThreadService threadService,
            TaskService taskService,
            ContactService contactService,
            AttachmentService attachmentService)
        {
            _settings = settings;
            _clock = clock ?? new SystemClock();
            _fileService = fileService ?? new FileService(settings, new MessageBoxUserNotifier());
            _templateService = templateService ?? new TemplateService(_fileService);
            _threadService = threadService ?? new ThreadService(_fileService, _templateService, settings);
            _taskService = taskService ?? new TaskService(settings, _templateService, _clock);
            _contactService = contactService ?? new ContactService(_fileService, _templateService, _clock);
            _attachmentService = attachmentService ?? new AttachmentService(settings, _fileService, _clock);
            _subjectFilenameCleaner = new SubjectFilenameCleaner(settings, _fileService);
            _engine = new EmailNoteEngine(settings, _templateService, (name, email) => _contactService.FormatContactLink(name, email));
        }

        private static void RememberContactEmail(Dictionary<string, string> contactEmails, string name, string email)
        {
            if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(email) && email.IndexOf('@') > 0
                && !contactEmails.ContainsKey(name))
            {
                contactEmails[name] = email.Trim();
            }
        }

        private static string EmailFor(Dictionary<string, string> contactEmails, string name)
        {
            string email;
            return name != null && contactEmails.TryGetValue(name, out email) ? email : null;
        }

        /// <summary>
        /// Converts the supplied <paramref name="mail"/> into a markdown note, creates the optional follow-up
        /// tasks and opens the resulting file in Obsidian (depending on settings). The method is asynchronous
        /// because it performs a number of I/O heavy operations (file moves, Outlook task creation, countdown
        /// dialog) that would otherwise block the Outlook UI thread.
        /// </summary>
        /// <param name="mail">The email that should be exported.</param>
        /// <param name="cancellationToken">Optional cancellation token to allow cancellation of long-running operations.</param>
        /// <returns>
        /// <c>true</c> when the email's note was actually exported; <c>false</c> when it was skipped
        /// (duplicate) or the core export failed. Callers that don't care may ignore the result;
        /// batch callers (e.g. Complete-Thread) use it to report an honest success count.
        /// </returns>
        public async Task<bool> ProcessEmail(
            MailItem mail,
            System.Threading.CancellationToken cancellationToken = default(System.Threading.CancellationToken),
            ContactInteractionMode contactMode = ContactInteractionMode.Interactive,
            bool bulkMode = false)
        {
            EmailSlingResult result = await ProcessEmailWithResult(mail, cancellationToken, contactMode, bulkMode);
            return result.Outcome == EmailSlingOutcome.Saved;
        }

        /// <summary>
        /// Same as <see cref="ProcessEmail"/>, but says whether the email was saved, was already in the
        /// vault, or failed (and why), so batch callers can report each case.
        /// </summary>
        public async Task<EmailSlingResult> ProcessEmailWithResult(
            MailItem mail,
            System.Threading.CancellationToken cancellationToken = default(System.Threading.CancellationToken),
            ContactInteractionMode contactMode = ContactInteractionMode.Interactive,
            bool bulkMode = false)
        {
            // Declare variables at method level so they're accessible throughout the method
            List<string> contactNames = new List<string>();
            // Each contact's address, so a new contact note records it and an existing note listing it
            // is recognised under a different display name.
            Dictionary<string, string> contactEmails = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string fileName = string.Empty;
            string fileNameNoExt = string.Empty;
            string filePath = string.Empty;
            string obsidianLinkPath = string.Empty;  // Added to store the path to use for Obsidian
            string conversationId = string.Empty;
            string threadNoteName = string.Empty;
            string threadFolderPath = string.Empty;
            string threadNotePath = string.Empty;
            bool shouldGroupThread = false;
            bool coreExportSucceeded = false;
            string failureMessage = null;

            // Get task options first if needed. Never prompt in bulk mode — a batch of emails would
            // otherwise pop one modal date dialog per email.
            if (!bulkMode && (_settings.CreateOutlookTask || _settings.CreateObsidianTask) && _settings.AskForDates)
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
                RememberContactEmail(contactEmails, mail.SenderName, _contactService.GetSenderEmail(mail));
            }

            List<EmailRecipientData> recipientData = OutlookEmailReader.ReadRecipients(mail, _contactService);
            foreach (EmailRecipientData recipient in recipientData)
            {
                contactNames.Add(recipient.Name);
                RememberContactEmail(contactEmails, recipient.Name, recipient.Email);
            }

            using (var status = new StatusService(silent: bulkMode))
            {
                try
                {
                    status.UpdateProgress("Processing email...", 0);

                    // Vault path pre-check before any file writes
                    string vaultPath = _settings.GetFullVaultPath();
                    if (!System.IO.Directory.Exists(vaultPath))
                    {
                        throw new System.IO.DirectoryNotFoundException(
                            $"Obsidian vault at \"{vaultPath}\" is not accessible. Check that the folder exists.");
                    }

                    // Check for cancellation
                    cancellationToken.ThrowIfCancellationRequested();

                    // Everything the note is built from, read from the MailItem once.
                    EmailData email = OutlookEmailReader.Read(mail, _contactService, recipientData);

                    // Build note title using settings (with null safety)
                    string senderClean = _contactService.GetFilenameSafeShortName(email.SenderName ?? "Unknown Sender");
                    string fileDateTime = email.ReceivedTime.ToString("yyyy-MM-dd-HHmm");
                    string subjectClean = CleanSubject(email.Subject ?? "No Subject");
                    string noteTitle = _engine.BuildNoteTitle(subjectClean, senderClean, email.ReceivedTime);

                    // Email threading logic moved to its own method
                    (conversationId, threadNoteName, threadFolderPath, threadNotePath, shouldGroupThread, obsidianLinkPath, fileName, filePath, fileNameNoExt) =
                        GetThreadingInfo(mail, email, subjectClean, senderClean, fileDateTime, "");
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

                    string realInternetMessageId = email.InternetMessageId;
                    string realEntryId = email.EntryId;

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

                    Dictionary<string, object> metadata = BuildEmailMetadata(
                        noteTitle,
                        email.SenderName ?? "Unknown Sender",
                        email.SenderEmail,
                        toLinked,
                        toEmails,
                        conversationId,
                        email.ReceivedTime,
                        realInternetMessageId,
                        realEntryId,
                        ccLinked,
                        ccEmails,
                        shouldGroupThread,
                        threadNoteName
                    );

                    string threadNoteLink = string.Empty;
                    if (shouldGroupThread)
                    {
                        threadNoteLink = $"[[{Path.GetFileNameWithoutExtension(threadNotePath)}]]";
                        metadata["threadNote"] = threadNoteLink;
                    }

                    // Convert and clean once; the note may be re-rendered below after a rename.
                    string noteBody = _engine.ConvertBody(email, metadata);

                    string renderedContent = BuildEmailNoteContent(
                        email,
                        metadata,
                        noteTitle,
                        subjectClean,
                        senderClean,
                        noteBody,
                        fileNameNoExt,
                        conversationId,
                        threadNoteLink);

                    status.UpdateProgress("Writing note file", 75);

                    // Check for cancellation before file operations
                    cancellationToken.ThrowIfCancellationRequested();

                    // Check for duplicate email before writing the note
                    if (_processedEmails.IsDuplicate(_settings.GetInboxPath(), realInternetMessageId, realEntryId, _clock.Now))
                    {
                        status.UpdateProgress("Duplicate email detected. Skipping note creation.", 100);
                        if (!bulkMode)
                        {
                            // The progress window closes as soon as this returns, so say it where it stays visible.
                            ToastForm.ShowToast("SlingMD: this email is already in your vault.");
                        }
                        return new EmailSlingResult { Outcome = EmailSlingOutcome.Duplicate };
                    }

                    if (shouldGroupThread)
                    {
                        // Get or create a semaphore for this thread folder to prevent race conditions
                        var semaphore = _threadFolderLocks.GetOrAdd(threadFolderPath, _ => new System.Threading.SemaphoreSlim(1, 1));

                        // Acquire the lock for this thread folder
                        await semaphore.WaitAsync();
                        try
                        {
                            // Write the new note for the current email to the thread folder with a
                            // temporary -eid{hash} suffix. Both ids can be empty (unsaved item with no
                            // InternetMessageID); that still produces a usable, if less unique, "-eid" suffix.
                            string emailId = !string.IsNullOrEmpty(realInternetMessageId) ? realInternetMessageId : realEntryId;
                            string baseName = _engine.BuildBaseFileName(email.ReceivedTime, subjectClean, senderClean, fileDateTime, true);
                            string tempFileName = _threadService.GetTemporaryThreadNoteName(threadFolderPath, baseName, emailId) + ".md";
                            string tempFilePath = Path.Combine(threadFolderPath, tempFileName);
                            _fileService.EnsureDirectoryExists(threadFolderPath);
                            _fileService.WriteUtf8File(tempFilePath, renderedContent);

                            // Add to cache to keep it synchronized
                            _processedEmails.Add(realInternetMessageId, realEntryId);

                            filePath = tempFilePath;
                            fileName = tempFileName;
                            fileNameNoExt = Path.GetFileNameWithoutExtension(tempFileName);

                            // Move all existing emails for this thread to the thread folder. Notes
                            // grouped with their attachments live one level down, in a folder that
                            // shares their name.
                            var mdFiles = Directory.GetFiles(_settings.GetInboxPath(), "*.md", SearchOption.TopDirectoryOnly)
                                .Concat(Directory.GetDirectories(_settings.GetInboxPath())
                                    .Select(dir => Path.Combine(dir, Path.GetFileName(dir) + ".md"))
                                    .Where(File.Exists))
                                .ToArray();
                            foreach (var file in mdFiles)
                            {
                                // Read front matter to get threadId. One locked/unreadable neighbour
                                // must not abort the export of an email whose note is already written
                                // (mirrors the per-file guard in EnsureEmailCacheIsBuilt).
                                bool inFrontMatter = false;
                                string foundThreadId = null;
                                IEnumerable<string> lines;
                                try
                                {
                                    lines = File.ReadLines(file).ToList();
                                }
                                catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                                {
                                    Logger.Instance.Warning($"EmailProcessor: skipping unreadable inbox note '{file}': {ex.Message}");
                                    continue;
                                }
                                foreach (var line in lines)
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
                                        _threadService.MoveToThreadFolder(file, threadFolderPath, threadNoteName);
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
                                renderedContent = BuildEmailNoteContent(
                                    email,
                                    metadata,
                                    noteTitle,
                                    subjectClean,
                                    senderClean,
                                    noteBody,
                                    fileNameNoExt,
                                    conversationId,
                                    threadNoteLink);
                                _fileService.WriteUtf8File(filePath, renderedContent);
                            }
                            else
                            {
                                // Resuffix couldn't track the just-written note (e.g. unparsable
                                // frontmatter date). The file on disk is still the -eid temp name,
                                // but obsidianLinkPath retains GetThreadingInfo's non-eid value —
                                // point it at the real file so Obsidian doesn't open "not found".
                                obsidianLinkPath = $"{threadNoteName}/{fileNameNoExt}";
                            }

                            // Create or update the thread summary note
                            await _threadService.UpdateThreadNote(threadFolderPath, threadNotePath, conversationId, threadNoteName, mail);

                            // Process attachments if enabled (inside semaphore to prevent race condition with file resuffixing)
                            if (ShouldProcessAttachments(_settings))
                            {
                                try
                                {
                                    status.UpdateProgress("Processing attachments", 77);
                                    var attachmentInfo = _attachmentService.ProcessAttachments(mail, filePath);
                                    string noteContent = File.ReadAllText(filePath);
                                    string linkedContent = _attachmentService.ResolveInlineImageLinks(noteContent, attachmentInfo, filePath);
                                    if (linkedContent != noteContent)
                                    {
                                        _fileService.WriteUtf8File(filePath, linkedContent);
                                    }

                                    string attachmentSection = _attachmentService.BuildAttachmentSection(attachmentInfo, filePath, linkedContent);
                                    if (attachmentSection.Length > 0)
                                    {
                                        _fileService.AppendToFile(filePath, attachmentSection);
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
                        // A DIFFERENT email can produce the same subject/sender/minute filename
                        // (same-message re-slings never get here — IsDuplicateEmail returns above).
                        // Never overwrite an existing note: resolve to a "_N"-suffixed free name
                        // and re-render so the task self-link and FileName tokens match the final
                        // name (mirrors the threaded branch after resuffixing).
                        string resolvedPath = _uniqueFilenameResolver.Resolve(_settings.GetInboxPath(), fileName, File.Exists);
                        if (resolvedPath != null && !string.Equals(resolvedPath, filePath, StringComparison.OrdinalIgnoreCase))
                        {
                            filePath = resolvedPath;
                            fileName = Path.GetFileName(resolvedPath);
                            fileNameNoExt = Path.GetFileNameWithoutExtension(resolvedPath);
                            obsidianLinkPath = fileNameNoExt;
                            renderedContent = BuildEmailNoteContent(
                                email,
                                metadata,
                                noteTitle,
                                subjectClean,
                                senderClean,
                                noteBody,
                                fileNameNoExt,
                                conversationId,
                                threadNoteLink);
                        }

                        // Write the note as usual to the inbox
                        _fileService.WriteUtf8File(filePath, renderedContent);

                        // Add to cache to keep it synchronized
                        _processedEmails.Add(realInternetMessageId, realEntryId);

                        // Process attachments if enabled
                        if (ShouldProcessAttachments(_settings))
                        {
                            try
                            {
                                status.UpdateProgress("Processing attachments", 77);
                                var attachmentInfo = _attachmentService.ProcessAttachments(mail, filePath);
                                filePath = _attachmentService.GroupNoteWithAttachments(filePath);
                                string noteContent = File.ReadAllText(filePath);
                                string linkedContent = _attachmentService.ResolveInlineImageLinks(noteContent, attachmentInfo, filePath);
                                if (linkedContent != noteContent)
                                {
                                    _fileService.WriteUtf8File(filePath, linkedContent);
                                }

                                string attachmentSection = _attachmentService.BuildAttachmentSection(attachmentInfo, filePath, linkedContent);
                                if (attachmentSection.Length > 0)
                                {
                                    _fileService.AppendToFile(filePath, attachmentSection);
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
                    coreExportSucceeded = true;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Error($"EmailProcessor.ProcessEmail failed: {ex.Message}");
                    failureMessage = ex.Message;
                    if (bulkMode)
                    {
                        // Background monitors and batch slings share this path; a modal dialog per
                        // failed email would block the Outlook UI thread (and stack up one per
                        // inbound message when the vault is offline). Non-modal toast instead,
                        // honoring the same notification mode NotificationService uses.
                        if (_settings.AutoSlingNotificationMode == "Toast")
                        {
                            try { ToastForm.ShowToast($"SlingMD: email export failed — {ex.Message}", isError: true); }
                            catch (System.Exception toastEx)
                            {
                                // Best effort: the export failure itself is already logged.
                                Logger.Instance.Warning($"EmailProcessor: failure toast could not be shown: {toastEx.Message}");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Error processing email: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            if (!coreExportSucceeded)
            {
                return new EmailSlingResult { Outcome = EmailSlingOutcome.Failed, ErrorMessage = failureMessage };
            }

            // Process contacts outside the StatusService block
            // This ensures the progress window doesn't block the contact dialog
            if (_settings.EnableContactSaving && contactNames.Count > 0)
            {
                try
                {
                    // Remove duplicates and sort
                    contactNames = contactNames.Distinct().OrderBy(n => n).ToList();

                    // Refresh existing managed contact notes and collect truly new contacts.
                    var newContacts = new List<string>();
                    var managedContactsToRefresh = new List<string>();
                    foreach (var contactName in contactNames)
                    {
                        if (_contactService.ManagedContactNoteExists(contactName))
                        {
                            managedContactsToRefresh.Add(contactName);
                        }
                        else if (!_contactService.ContactExists(contactName)
                            && _contactService.FindContactNoteByEmail(EmailFor(contactEmails, contactName)) == null)
                        {
                            newContacts.Add(contactName);
                        }
                    }

                    foreach (var contactName in managedContactsToRefresh)
                    {
                        _contactService.CreateContactNote(contactName, EmailFor(contactEmails, contactName));
                    }

                    // For truly new contacts, try fuzzy matching first (when enabled)
                    var unfuzzyNewContacts = new List<string>();
                    string subjectForLog = string.Empty;
                    try
                    {
                        subjectForLog = mail.Subject ?? string.Empty;
                    }
                    catch (System.Exception ex) when (ex is System.Runtime.InteropServices.COMException
                        || ex is System.Runtime.InteropServices.InvalidComObjectException)
                    {
                        // The subject only labels the fuzzy-match log lines, so a failed read must not stop
                        // the contacts. It was read once already to build the note, so this means the item
                        // went away mid-sling (store offline, item moved or released).
                        Logger.Instance.Warning($"EmailProcessor: could not read the subject for the contact log: {ex.GetType().Name}: {ex.Message}");
                    }

                    foreach (var contactName in newContacts)
                    {
                        bool handled = _contactService.TryHandleFuzzyMatch(
                            contactName, EmailFor(contactEmails, contactName), contactMode, subjectForLog);
                        if (!handled)
                        {
                            unfuzzyNewContacts.Add(contactName);
                        }
                    }

                    // In automated mode, silently create all remaining new contacts
                    if (contactMode == ContactInteractionMode.Automated)
                    {
                        foreach (var contactName in unfuzzyNewContacts)
                        {
                            _contactService.CreateContactNote(contactName, EmailFor(contactEmails, contactName));
                        }
                    }
                    else if (unfuzzyNewContacts.Count > 0)
                    {
                        // Show contact confirmation dialog in interactive mode
                        using (var dialog = new ContactConfirmationDialog(unfuzzyNewContacts))
                        {
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                foreach (var contactName in dialog.SelectedContacts)
                                {
                                    _contactService.CreateContactNote(contactName, EmailFor(contactEmails, contactName));
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
            
            // Copy the saved note's path for pasting elsewhere. Skipped in bulk mode so background
            // and batch slings never overwrite what the user has on the clipboard.
            if (!bulkMode && _settings.CopyNotePathToClipboard && !string.IsNullOrEmpty(filePath))
            {
                try
                {
                    Clipboard.SetText(filePath);
                }
                catch (System.Exception ex) when (ex is System.Runtime.InteropServices.ExternalException || ex is System.Threading.ThreadStateException)
                {
                    Logger.Instance.Warning($"EmailProcessor: could not copy note path to clipboard: {ex.Message}");
                }
            }

            // Launch Obsidian if enabled. Suppressed in bulk mode — the batch caller launches once
            // at the end instead of once per email.
            if (!bulkMode && _settings.LaunchObsidian)
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

                    // Open by the note's path from the vault root. A thread-relative path such as
                    // "Thread/Note" only resolves when the inbox is the vault root itself.
                    _fileService.LaunchObsidian(_settings.VaultName, GetVaultRelativePath(filePath) ?? obsidianLinkPath);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error launching Obsidian: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (!bulkMode)
            {
                // Without Obsidian opening on the note, nothing else confirms the sling worked.
                ToastForm.ShowToast($"SlingMD: saved {GetVaultRelativePath(filePath) ?? Path.GetFileName(filePath)}");
            }

            if (!bulkMode && _settings.ShowNoteInExplorer)
            {
                _fileService.ShowInExplorer(filePath);
            }

            return new EmailSlingResult { Outcome = EmailSlingOutcome.Saved, NotePath = filePath };
        }

        private string CleanSubject(string subject) => _subjectFilenameCleaner.Clean(subject);

        /// <summary>
        /// Returns <paramref name="path"/> relative to the vault root with forward slashes, or
        /// <c>null</c> when it does not lie inside the vault.
        /// </summary>
        internal string GetVaultRelativePath(string path) => _engine.GetVaultRelativePath(path);

        /// <summary>
        /// Waits for a file to become available by checking its existence and attempting to open it.
        /// Uses exponential backoff retry logic (see <see cref="MaxFileWaitAttempts"/> and
        /// <see cref="InitialFileWaitDelayMs"/>). Replaces hard-coded delays.
        /// </summary>
        /// <param name="filePath">The full path to the file to wait for.</param>
        private async Task WaitForFileAvailability(string filePath)
        {
            int attempt = 0;
            int delayMs = InitialFileWaitDelayMs;

            while (attempt < MaxFileWaitAttempts)
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

                // Exponential backoff: InitialFileWaitDelayMs doubled on each retry.
                attempt++;
                if (attempt < MaxFileWaitAttempts)
                {
                    await Task.Delay(delayMs);
                    delayMs *= 2;
                }
            }

            // If we get here, file still isn't available but we've tried our best
            // Continue anyway - the worst case is the same as before
        }

        /// <summary>
        /// Applies <see cref="ObsidianSettings.BodyCleanupPatterns"/> to the Markdown body and, when a
        /// pattern removed text, records <c>bodyCleaned: true</c> in <paramref name="metadata"/> so the
        /// note is never mistaken for a verbatim copy of the email.
        /// </summary>
        internal string CleanNoteBody(string body, Dictionary<string, object> metadata) => _engine.CleanNoteBody(body, metadata);

        /// <summary>
        /// Returns thread-related info for an email, including paths and names.
        /// </summary>
        private string BuildEmailNoteContent(EmailData email, Dictionary<string, object> metadata, string noteTitle, string subjectClean, string senderClean, string noteBody, string fileNameNoExt, string conversationId, string threadNoteLink)
        {
            string taskBlock = string.Empty;
            if (_settings.CreateObsidianTask && _taskService.ShouldCreateTasks)
            {
                List<string> taskTags = _settings.DefaultTaskTags ?? new List<string>();
                taskBlock = _taskService.GenerateObsidianTask(fileNameNoExt, taskTags);
                if (!string.IsNullOrWhiteSpace(taskBlock))
                {
                    taskBlock += Environment.NewLine + Environment.NewLine;
                }
            }

            return _engine.RenderNote(email, metadata, noteTitle, subjectClean, senderClean, noteBody, fileNameNoExt, conversationId, threadNoteLink, taskBlock);
        }

        /// <summary>
        /// Returns thread-related info for an email, including paths and names.
        /// </summary>
        private (string conversationId, string threadNoteName, string threadFolderPath, string threadNotePath, bool shouldGroupThread, string obsidianLinkPath, string fileName, string filePath, string fileNameNoExt) GetThreadingInfo(MailItem mail, EmailData email, string subjectClean, string senderClean, string fileDateTime, string fileNameNoExt)
        {
            string conversationId = _threadService.GetConversationId(mail);
            string firstRecipient = _contactService.GetFilenameSafeShortName(email.FirstToRecipientName);
            string threadFolderName = _threadService.GetThreadFolderName(mail, subjectClean, senderClean, firstRecipient);
            string threadNoteName = _threadService.GetThreadNoteName(mail, subjectClean, senderClean, firstRecipient);
            string threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadFolderName);
            string threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
            var threadInfo = _threadService.FindExistingThread(conversationId, _settings.GetInboxPath());
            bool hasExistingThread = threadInfo.hasExistingThread;
            string earliestEmailThreadName = threadInfo.earliestEmailThreadName;
            int emailCount = threadInfo.emailCount;
            bool shouldGroupThread = _settings.GroupEmailThreads
                && ((hasExistingThread && emailCount >= 1) || _settings.AlwaysCreateThreadFolder);

            if (shouldGroupThread)
            {
                threadNoteName = earliestEmailThreadName ?? threadFolderName;
                threadFolderPath = Path.Combine(_settings.GetInboxPath(), threadNoteName);
                threadNotePath = Path.Combine(threadFolderPath, $"0-{threadNoteName}.md");
            }

            string fileNameNoExtResult = _engine.BuildBaseFileName(email.ReceivedTime, subjectClean, senderClean, fileDateTime, shouldGroupThread);
            if (!shouldGroupThread && _settings.GroupEmailThreads)
            {
                // Give a standalone note the name it would have inside a thread, derived from the
                // same front-matter date the thread rename pass reads. Moving it into a thread
                // folder later then keeps its name, so the Obsidian task link and links from other
                // notes don't break when a reply is slung.
                string frontMatterDate = _engine.FormatFrontMatterDate(email.ReceivedTime);
                if (ThreadService.TryParseNoteDate(frontMatterDate, out DateTime noteDate))
                {
                    fileNameNoExtResult = _threadService.GetThreadedNoteName(fileNameNoExtResult, noteDate, threadFolderPath);
                }
            }
            // Keep every path the note can produce inside the Windows path limits. Inside a thread
            // folder, ResuffixThreadNotes applies the same budget to the final name.
            fileNameNoExtResult = shouldGroupThread
                ? _fileService.FitNoteName(threadFolderPath, fileNameNoExtResult, ".md", false)
                : _fileService.FitNoteName(_settings.GetInboxPath(), fileNameNoExtResult, ".md", true);
            string fileName = fileNameNoExtResult + ".md";
            string filePath = shouldGroupThread
                ? Path.Combine(threadFolderPath, fileName)
                : Path.Combine(_settings.GetInboxPath(), fileName);
            string obsidianLinkPath = shouldGroupThread
                ? $"{threadNoteName}/{fileNameNoExtResult}"
                : fileNameNoExtResult;

            return (conversationId, threadNoteName, threadFolderPath, threadNotePath, shouldGroupThread, obsidianLinkPath, fileName, filePath, fileNameNoExtResult);
        }

        internal Dictionary<string, object> BuildEmailMetadata(
            string noteTitle,
            string senderName,
            string senderEmail,
            List<string> toLinked,
            List<string> toEmails,
            string conversationId,
            DateTime receivedTime,
            string realInternetMessageId,
            string realEntryId,
            List<string> ccLinked,
            List<string> ccEmails,
            bool shouldGroupThread,
            string threadNoteName)
        {
            return _engine.BuildMetadata(noteTitle, senderName, senderEmail, toLinked, toEmails, conversationId, receivedTime,
                realInternetMessageId, realEntryId, ccLinked, ccEmails, shouldGroupThread, threadNoteName);
        }
    }
}
