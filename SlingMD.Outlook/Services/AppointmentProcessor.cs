using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;
using Logger = SlingMD.Outlook.Helpers.Logger;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    public enum AppointmentProcessingResult
    {
        Success,
        Skipped,       // Duplicate or cancelled
        Error
    }

    /// <summary>
    /// Orchestrates the full life-cycle of turning an <see cref="AppointmentItem"/> into a properly
    /// formatted markdown note inside the user's Obsidian vault. Mirrors the design of
    /// <see cref="EmailProcessor"/> and reuses all shared services without any modifications to them.
    /// </summary>
    public class AppointmentProcessor
    {
        // Static cache for duplicate detection keyed on GlobalAppointmentID
        private static readonly ConcurrentDictionary<string, byte> _processedAppointmentIds
            = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        // Static lock dictionary to prevent race conditions when multiple recurring instances target the same thread folder
        private static readonly ConcurrentDictionary<string, SemaphoreSlim> _recurringFolderLocks
            = new ConcurrentDictionary<string, SemaphoreSlim>(StringComparer.OrdinalIgnoreCase);

        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;
        private readonly TaskService _taskService;
        private readonly ContactService _contactService;
        private readonly AttachmentService _attachmentService;
        private readonly AppointmentNoteEngine _noteEngine;
        private readonly SubjectFilenameCleaner _subjectFilenameCleaner;
        private readonly UniqueFilenameResolver _uniqueFilenameResolver = new UniqueFilenameResolver();
        private readonly IClock _clock;
        private readonly ReminderDueDateCalculator _reminderCalculator;
        private readonly IEmailPrompts _prompts;

        private List<string> _bulkErrors = new List<string>();
        private int _bulkAmbiguousCount;

        public List<string> GetBulkErrors()
        {
            List<string> errors = new List<string>(_bulkErrors);
            _bulkErrors.Clear();
            return errors;
        }

        /// <summary>
        /// Returns the number of ambiguous contact match events accumulated since the last call, then resets the counter.
        /// </summary>
        public int GetAndClearAmbiguousCount()
        {
            int count = _bulkAmbiguousCount;
            _bulkAmbiguousCount = 0;
            return count;
        }

        public AppointmentProcessor(ObsidianSettings settings, IClock clock = null, IEmailPrompts prompts = null)
        {
            _settings = settings;
            _prompts = prompts ?? new DialogEmailPrompts();
            _fileService = new FileService(settings, new MessageBoxUserNotifier());
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService, settings, new MessageBoxUserNotifier());
            _clock = clock ?? new SystemClock();
            _taskService = new TaskService(settings, _templateService, _clock);
            _contactService = new ContactService(_fileService, _templateService);
            _attachmentService = new AttachmentService(settings, _fileService);
            _noteEngine = new AppointmentNoteEngine(settings, _templateService, _contactService.NoteFileExists);
            _subjectFilenameCleaner = new SubjectFilenameCleaner(settings, _fileService);
            _reminderCalculator = new ReminderDueDateCalculator();
        }

        /// <summary>
        /// Converts the supplied <paramref name="appointment"/> into a markdown note and writes it to
        /// the configured appointments folder in the Obsidian vault.
        /// </summary>
        /// <param name="appointment">The calendar item to export.</param>
        /// <param name="bulkMode">When true the StatusService progress window is suppressed.</param>
        /// <param name="cancellationToken">Optional token to cancel long-running operations.</param>
        public async System.Threading.Tasks.Task<AppointmentProcessingResult> ProcessAppointment(
            AppointmentItem appointment,
            bool bulkMode = false,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            // --- Cancelled-appointment guard ---
            bool isCancelled = false;
            try
            {
                isCancelled = appointment.MeetingStatus == OlMeetingStatus.olMeetingCanceled;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not read MeetingStatus: {ex.Message}");
            }

            if (isCancelled && !_settings.SaveCancelledAppointments)
            {
                return AppointmentProcessingResult.Skipped;
            }

            // --- Series master guard ---
            OlRecurrenceState earlyRecurrenceState = OlRecurrenceState.olApptNotRecurring;
            try
            {
                earlyRecurrenceState = appointment.RecurrenceState;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not read RecurrenceState (early check): {ex.Message}");
            }

            if (earlyRecurrenceState == OlRecurrenceState.olApptMaster)
            {
                if (!bulkMode)
                {
                    System.Windows.Forms.DialogResult masterResult = System.Windows.Forms.MessageBox.Show(
                        "You've selected the recurring series master. Would you like to save the next upcoming instance instead?",
                        "Recurring Series",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (masterResult == System.Windows.Forms.DialogResult.Yes)
                    {
                        RecurrencePattern masterPattern = null;
                        AppointmentItem nextInstance = null;
                        try
                        {
                            masterPattern = appointment.GetRecurrencePattern();
                            nextInstance = masterPattern.GetOccurrence(DateTime.Today);
                            await ProcessAppointment(nextInstance, bulkMode, cancellationToken);
                        }
                        catch (System.Runtime.InteropServices.COMException comEx)
                        {
                            System.Windows.Forms.MessageBox.Show(
                                $"Could not find the next upcoming instance of this series: {comEx.Message}",
                                "Error",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Warning);
                        }
                        finally
                        {
                            if (nextInstance != null)
                            {
                                Marshal.ReleaseComObject(nextInstance);
                            }

                            if (masterPattern != null)
                            {
                                Marshal.ReleaseComObject(masterPattern);
                            }
                        }
                    }
                }

                // In bulk mode or after redirect: skip the series master itself
                return AppointmentProcessingResult.Skipped;
            }

            // --- Extract all metadata from the COM object up-front ---
            AppointmentData data = OutlookAppointmentReader.Read(appointment, _clock.Now);
            string subject = data.Subject;
            string location = data.Location;
            DateTime startTime = data.Start;
            string organizerName = data.OrganizerName;
            string globalAppointmentId = data.GlobalAppointmentId;
            List<string> requiredAttendees = data.RequiredAttendees;
            List<string> optionalAttendees = data.OptionalAttendees;

            // --- Duplicate detection ---
            if (!string.IsNullOrWhiteSpace(globalAppointmentId))
            {
                if (IsDuplicateAppointment(globalAppointmentId))
                {
                    return AppointmentProcessingResult.Skipped;
                }
            }

            // --- Vault path pre-check before any file writes ---
            string vaultPath = _settings.GetFullVaultPath();
            if (!System.IO.Directory.Exists(vaultPath))
            {
                throw new System.IO.DirectoryNotFoundException(
                    $"Obsidian vault at \"{vaultPath}\" is not accessible. Check that the folder exists.");
            }

            // --- Task creation flags ---
            bool createObsidianTask = _settings.AppointmentTaskCreation == "Obsidian"
                                   || _settings.AppointmentTaskCreation == "Both";
            bool createOutlookTask = _settings.AppointmentTaskCreation == "Outlook"
                                  || _settings.AppointmentTaskCreation == "Both";

            if (createObsidianTask || createOutlookTask)
            {
                if (!bulkMode && _settings.AskForDates)
                {
                    TaskDateChoice choice = _prompts.PromptTaskOptions(
                        _settings.DefaultDueDays, _settings.DefaultReminderDays,
                        _settings.DefaultReminderHour, _settings.UseRelativeReminder);
                    if (choice != null)
                    {
                        _taskService.InitializeTaskSettings(
                            choice.DueDays, choice.ReminderDays,
                            choice.ReminderHour, choice.UseRelativeReminder);
                    }
                    else
                    {
                        _taskService.DisableTaskCreation();
                        createObsidianTask = false;
                        createOutlookTask = false;
                    }
                }
                else
                {
                    _taskService.InitializeTaskSettings(
                        _settings.DefaultDueDays, _settings.DefaultReminderDays,
                        _settings.DefaultReminderHour, _settings.UseRelativeReminder);
                }
            }

            // --- Build the note title and file name ---
            string subjectClean = CleanSubject(subject);
            string dateStr = startTime.ToString("yyyy-MM-dd");
            string organizerShortName = _contactService.GetFilenameSafeShortName(
                string.IsNullOrWhiteSpace(organizerName) ? "Unknown Organizer" : organizerName);

            string noteTitle = _noteEngine.BuildNoteTitle(subjectClean, organizerShortName, startTime);

            string fileNameNoExt = _fileService.CleanFileName(noteTitle);
            if (string.IsNullOrWhiteSpace(fileNameNoExt))
            {
                fileNameNoExt = $"Appointment-{dateStr}";
            }

            string fileName = fileNameNoExt + ".md";
            string appointmentsPath = _settings.GetAppointmentsPath();
            string filePath = Path.Combine(appointmentsPath, fileName);

            // --- Determine whether this instance goes into a recurring-series thread folder ---
            bool isRecurring = _noteEngine.ShouldGroupAsRecurring(data.Recurrence);
            string recurringThreadFolderPath = null;
            string recurringThreadFolderName = null;
            string recurringInstanceFilePath = null;
            string recurringThreadNotePath = null;

            if (isRecurring)
            {
                recurringThreadFolderName = GetRecurringThreadFolderName(subjectClean);
                recurringThreadFolderPath = Path.Combine(appointmentsPath, recurringThreadFolderName);

                string dateStamp = startTime.ToString("yyyy-MM-dd");
                string instanceFileName = $"{dateStamp} - {subjectClean}.md";
                recurringInstanceFilePath = Path.Combine(recurringThreadFolderPath, instanceFileName);

                // Handle same-day collision by appending time suffix
                if (File.Exists(recurringInstanceFilePath))
                {
                    string timeSuffix = startTime.ToString("_HHmm");
                    instanceFileName = $"{dateStamp}{timeSuffix} - {subjectClean}.md";
                    recurringInstanceFilePath = Path.Combine(recurringThreadFolderPath, instanceFileName);
                }

                recurringThreadNotePath = Path.Combine(recurringThreadFolderPath, $"0-{recurringThreadFolderName}.md");

                // Update filePath/fileNameNoExt so Obsidian launch uses the correct path
                filePath = recurringInstanceFilePath;
                fileNameNoExt = Path.GetFileNameWithoutExtension(recurringInstanceFilePath);
            }

            // --- Build frontmatter ---
            Dictionary<string, object> metadata = _noteEngine.BuildMetadata(noteTitle, data, fileNameNoExt);

            // --- Generate Obsidian task block ---
            string taskBlock = string.Empty;
            if (createObsidianTask && _taskService.ShouldCreateTasks)
            {
                List<string> taskTags = _settings.AppointmentDefaultNoteTags ?? new List<string>();
                taskBlock = _taskService.GenerateObsidianTask(fileNameNoExt, taskTags);
            }

            // --- Render note content ---
            string renderedContent = _noteEngine.BuildNoteContent(
                metadata,
                noteTitle,
                subjectClean,
                data,
                dateStr,
                fileNameNoExt,
                taskBlock);

            // --- Write file and process attachments ---
            bool coreExportSucceeded = false;

            if (bulkMode)
            {
                // In bulk mode skip the StatusService UI
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (isRecurring)
                    {
                        SemaphoreSlim recurSemaphore = _recurringFolderLocks.GetOrAdd(recurringThreadFolderPath, _ => new SemaphoreSlim(1, 1));
                        await recurSemaphore.WaitAsync(cancellationToken);
                        try
                        {
                            _fileService.EnsureDirectoryExists(recurringThreadFolderPath);
                            _fileService.WriteUtf8File(recurringInstanceFilePath, renderedContent);
                            WriteRecurringThreadSummary(recurringThreadFolderPath, recurringThreadFolderName, recurringThreadNotePath, subjectClean);
                            CreateCompanionMeetingNote(recurringInstanceFilePath, noteTitle, organizerName,
                                string.Join(", ", requiredAttendees), dateStr, location, recurringThreadFolderPath);
                        }
                        finally
                        {
                            recurSemaphore.Release();
                        }
                    }
                    else
                    {
                        _fileService.EnsureDirectoryExists(appointmentsPath);
                        _fileService.WriteUtf8File(filePath, renderedContent);
                        CreateCompanionMeetingNote(filePath, noteTitle, organizerName,
                            string.Join(", ", requiredAttendees), dateStr, location, appointmentsPath);
                    }

                    if (!string.IsNullOrWhiteSpace(globalAppointmentId))
                    {
                        _processedAppointmentIds.TryAdd(globalAppointmentId, 0);
                    }

                    if (_settings.AppointmentSaveAttachments)
                    {
                        ProcessAppointmentAttachments(appointment, filePath);
                    }

                    coreExportSucceeded = true;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Error($"Error processing appointment (bulk): {ex.Message}", ex);
                    _bulkErrors.Add($"{subject}: {ex.Message}");
                    return AppointmentProcessingResult.Error;
                }
            }
            else
            {
                using (StatusService status = new StatusService(ProgressFormView.Open()))
                {
                    try
                    {
                        status.UpdateProgress("Processing appointment...", 0);
                        cancellationToken.ThrowIfCancellationRequested();

                        status.UpdateProgress("Building metadata", 25);

                        status.UpdateProgress("Writing note", 50);

                        if (isRecurring)
                        {
                            SemaphoreSlim recurSemaphore = _recurringFolderLocks.GetOrAdd(recurringThreadFolderPath, _ => new SemaphoreSlim(1, 1));
                            await recurSemaphore.WaitAsync(cancellationToken);
                            try
                            {
                                _fileService.EnsureDirectoryExists(recurringThreadFolderPath);
                                _fileService.WriteUtf8File(recurringInstanceFilePath, renderedContent);
                                WriteRecurringThreadSummary(recurringThreadFolderPath, recurringThreadFolderName, recurringThreadNotePath, subjectClean);
                                CreateCompanionMeetingNote(recurringInstanceFilePath, noteTitle, organizerName,
                                    string.Join(", ", requiredAttendees), dateStr, location, recurringThreadFolderPath);
                            }
                            finally
                            {
                                recurSemaphore.Release();
                            }
                        }
                        else
                        {
                            _fileService.EnsureDirectoryExists(appointmentsPath);
                            _fileService.WriteUtf8File(filePath, renderedContent);
                            CreateCompanionMeetingNote(filePath, noteTitle, organizerName,
                                string.Join(", ", requiredAttendees), dateStr, location, appointmentsPath);
                        }

                        if (!string.IsNullOrWhiteSpace(globalAppointmentId))
                        {
                            _processedAppointmentIds.TryAdd(globalAppointmentId, 0);
                        }

                        if (_settings.AppointmentSaveAttachments)
                        {
                            status.UpdateProgress("Processing attachments", 75);
                            ProcessAppointmentAttachments(appointment, filePath);
                        }

                        status.UpdateProgress("Complete", 100);
                        coreExportSucceeded = true;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(
                            $"Error processing appointment: {ex.Message}",
                            "SlingMD Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }

            if (!coreExportSucceeded)
            {
                return AppointmentProcessingResult.Error;
            }

            // --- Process contacts after core export (outside StatusService block so dialogs are not blocked) ---
            if (_settings.EnableContactSaving)
            {
                try
                {
                    // Collect attendee names from organizer, required, and optional (excluding resource/room attendees)
                    List<string> contactNames = new List<string>();

                    if (!string.IsNullOrWhiteSpace(organizerName))
                    {
                        contactNames.Add(organizerName.Replace("[[", "").Replace("]]", ""));
                    }

                    foreach (string name in requiredAttendees)
                    {
                        string cleaned = name.Replace("[[", "").Replace("]]", "");
                        if (!string.IsNullOrWhiteSpace(cleaned))
                        {
                            contactNames.Add(cleaned);
                        }
                    }

                    foreach (string name in optionalAttendees)
                    {
                        string cleaned = name.Replace("[[", "").Replace("]]", "");
                        if (!string.IsNullOrWhiteSpace(cleaned))
                        {
                            contactNames.Add(cleaned);
                        }
                    }

                    if (contactNames.Count > 0)
                    {
                        // Deduplicate and sort
                        contactNames = contactNames.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(n => n).ToList();

                        // Separate managed contacts (refresh) from truly new contacts (prompt)
                        List<string> newContacts = new List<string>();
                        List<string> managedContactsToRefresh = new List<string>();

                        foreach (string contactName in contactNames)
                        {
                            if (_contactService.ManagedContactNoteExists(contactName))
                            {
                                managedContactsToRefresh.Add(contactName);
                            }
                            else if (!_contactService.ContactExists(contactName))
                            {
                                newContacts.Add(contactName);
                            }
                        }

                        // Refresh existing managed contact notes
                        foreach (string contactName in managedContactsToRefresh)
                        {
                            _contactService.CreateContactNote(contactName);
                        }

                        // Determine contact interaction mode: bulk runs are always Automated
                        ContactInteractionMode contactMode = bulkMode
                            ? ContactInteractionMode.Automated
                            : ContactInteractionMode.Interactive;

                        // Try fuzzy matching for new contacts
                        List<string> unfuzzyNewContacts = new List<string>();
                        foreach (string contactName in newContacts)
                        {
                            bool handled = _contactService.TryHandleFuzzyMatch(
                                contactName, null, contactMode, subject);
                            if (!handled)
                            {
                                unfuzzyNewContacts.Add(contactName);
                            }
                        }

                        // In bulk (Automated) mode, accumulate the ambiguous count
                        if (bulkMode)
                        {
                            _bulkAmbiguousCount += _contactService.GetAndClearAmbiguousCount();
                        }

                        // In single mode, show dialog for remaining new contacts; in bulk mode, create silently
                        if (bulkMode)
                        {
                            foreach (string contactName in unfuzzyNewContacts)
                            {
                                _contactService.CreateContactNote(contactName);
                            }
                        }
                        else if (unfuzzyNewContacts.Count > 0)
                        {
                            foreach (string contactName in _prompts.ConfirmNewContacts(unfuzzyNewContacts))
                            {
                                _contactService.CreateContactNote(contactName);
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    if (!bulkMode)
                    {
                        MessageBox.Show(
                            $"Error processing contacts: {ex.Message}",
                            "SlingMD Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                    else
                    {
                        _bulkErrors.Add($"Contact processing error for '{subject}': {ex.Message}");
                    }
                }
            }

            // --- Create Outlook task if enabled ---
            if (createOutlookTask && _taskService.ShouldCreateTasks)
            {
                Microsoft.Office.Interop.Outlook.TaskItem task = null;
                try
                {
                    task = Globals.ThisAddIn.Application.CreateItem(OlItemType.olTaskItem)
                           as Microsoft.Office.Interop.Outlook.TaskItem;
                    if (task != null)
                    {
                        task.Subject = $"Follow up: {subject}";
                        task.Body = $"Follow up on appointment: {subject}\n"
                                  + $"Date: {startTime:yyyy-MM-dd HH:mm}\n"
                                  + $"Location: {location}";

                        TaskDueDates dates = _reminderCalculator.Calculate(_clock.Now, new TaskDueSettings
                        {
                            DefaultDueDays = _settings.DefaultDueDays,
                            UseRelativeReminder = _settings.UseRelativeReminder,
                            DefaultReminderDays = _settings.DefaultReminderDays,
                            DefaultReminderHour = _settings.DefaultReminderHour
                        });

                        task.DueDate = dates.DueDate;
                        task.ReminderSet = true;
                        task.ReminderTime = dates.ReminderDateTime;
                        task.Save();
                    }
                }
                catch (System.Exception ex)
                {
                    if (!bulkMode)
                    {
                        MessageBox.Show(
                            $"Could not create Outlook task: {ex.Message}",
                            "Task Creation Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                    else
                    {
                        Logger.Instance.Error($"Could not create Outlook task for appointment '{subject}': {ex.Message}", ex);
                    }
                }
                finally
                {
                    if (task != null)
                    {
                        Marshal.ReleaseComObject(task);
                    }
                }
            }

            // Launch Obsidian if enabled (only in non-bulk mode)
            if (!bulkMode && _settings.LaunchObsidian)
            {
                try
                {
                    if (_settings.ShowCountdown && _settings.ObsidianDelaySeconds > 0)
                    {
                        using (Forms.CountdownForm countdown = new Forms.CountdownForm(_settings.ObsidianDelaySeconds))
                        {
                            countdown.ShowDialog();
                        }
                    }
                    else if (_settings.ObsidianDelaySeconds > 0)
                    {
                        await System.Threading.Tasks.Task.Delay(_settings.ObsidianDelaySeconds * 1000);
                    }

                    // The full path (LaunchObsidian makes it vault-relative) opens the right note even
                    // when another note in the vault has the same name; the bare name is the fallback.
                    string vaultRoot = _settings.GetFullVaultPath();
                    bool insideVault = !string.IsNullOrEmpty(filePath) && !string.IsNullOrEmpty(vaultRoot)
                        && filePath.StartsWith(vaultRoot.TrimEnd('\\', '/') + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
                    _fileService.LaunchObsidian(_settings.VaultName, insideVault ? filePath : fileNameNoExt);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(
                        $"Error launching Obsidian: {ex.Message}",
                        "SlingMD Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }

            if (!bulkMode && _settings.ShowNoteInExplorer)
            {
                _fileService.ShowInExplorer(filePath);
            }

            return AppointmentProcessingResult.Success;
        }

        // --- Recurring meeting helpers ---

        private string GetRecurringThreadFolderName(string cleanedSubject)
        {
            // Used as a folder name: drop the trailing dot/space CleanFileName keeps and Windows drops.
            return FolderNameSanitizer.Sanitize(_fileService.CleanFileName(cleanedSubject));
        }

        private void WriteRecurringThreadSummary(
            string threadFolderPath,
            string threadFolderName,
            string threadNotePath,
            string title)
        {
            try
            {
                string threadContent = _noteEngine.BuildRecurringSummaryContent(threadFolderPath, threadFolderName, title);
                _fileService.WriteUtf8File(threadNotePath, threadContent);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"Failed to write recurring thread summary: {ex.Message}", ex);
            }
        }

        private void CreateCompanionMeetingNote(
            string appointmentNotePath,
            string appointmentNoteTitle,
            string organizer,
            string attendees,
            string date,
            string location,
            string outputFolder)
        {
            if (!_settings.CreateMeetingNotes)
            {
                return;
            }

            try
            {
                string appointmentFileNameNoExt = Path.GetFileNameWithoutExtension(appointmentNotePath);
                string meetingNotePath = Path.Combine(outputFolder, AppointmentNoteEngine.MeetingNoteFileName(appointmentFileNameNoExt));

                if (File.Exists(meetingNotePath))
                {
                    return;
                }

                string meetingNoteContent = _noteEngine.BuildMeetingNoteContent(
                    appointmentFileNameNoExt, appointmentNoteTitle, organizer, attendees, date, location);
                _fileService.WriteUtf8File(meetingNotePath, meetingNoteContent);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"Failed to create companion meeting note: {ex.Message}", ex);
            }
        }

        // --- Private helpers ---

        private bool IsDuplicateAppointment(string globalAppointmentId)
        {
            return !string.IsNullOrWhiteSpace(globalAppointmentId)
                   && _processedAppointmentIds.ContainsKey(globalAppointmentId);
        }

        private string CleanSubject(string subject) => _subjectFilenameCleaner.Clean(subject);

        /// <summary>
        /// Readable recurrence value for the note ("None", "Series", "Occurrence", "Exception") instead of
        /// Outlook's enum name (e.g. "olApptNotRecurring"). See <see cref="AppointmentNoteEngine.DescribeRecurrence"/>.
        /// </summary>
        internal static string DescribeRecurrence(OlRecurrenceState recurrenceState)
        {
            return AppointmentNoteEngine.DescribeRecurrence(OutlookAppointmentReader.ToRecurrence(recurrenceState));
        }

        private void ProcessAppointmentAttachments(AppointmentItem appointment, string noteFilePath)
        {
            Attachments attachments = null;
            try
            {
                attachments = appointment.Attachments;
                int attachmentCount = 0;
                try { attachmentCount = attachments.Count; }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"Could not read attachment count: {ex.Message}");
                    return;
                }

                if (attachmentCount == 0)
                {
                    return;
                }

                string targetFolder = Path.GetDirectoryName(noteFilePath);
                _fileService.EnsureDirectoryExists(targetFolder);

                List<SavedAttachment> savedAttachments = new List<SavedAttachment>();

                for (int i = 1; i <= attachmentCount; i++)
                {
                    Attachment attachment = null;
                    try
                    {
                        attachment = attachments[i];

                        string attachmentFileName = string.Empty;
                        try { attachmentFileName = attachment.FileName ?? string.Empty; }
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Debug($"Could not read attachment filename: {ex.Message}");
                        }

                        // Skip .ics calendar files
                        string extension = Path.GetExtension(attachmentFileName).ToLowerInvariant();
                        if (extension == ".ics")
                        {
                            continue;
                        }

                        // Save attachment to note folder
                        string safeFilename = _fileService.CleanFileName(attachmentFileName);
                        if (string.IsNullOrWhiteSpace(safeFilename))
                        {
                            string originalExt = Path.GetExtension(attachmentFileName);
                            safeFilename = string.IsNullOrEmpty(originalExt)
                                ? $"attachment-{i}.dat"
                                : $"attachment-{i}{originalExt}";
                        }

                        string fullPath = _uniqueFilenameResolver.Resolve(targetFolder, safeFilename, File.Exists);
                        if (fullPath == null)
                        {
                            Logger.Instance.Warning($"Failed to find unique filename for appointment attachment: {attachmentFileName}");
                            continue;
                        }
                        safeFilename = Path.GetFileName(fullPath);

                        try
                        {
                            attachment.SaveAsFile(fullPath);
                            savedAttachments.Add(new SavedAttachment
                            {
                                FullPath = fullPath,
                                IsInline = false
                            });
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Instance.Error($"Failed to save appointment attachment '{attachmentFileName}': {ex.Message}", ex);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Error($"Failed to process appointment attachment {i}: {ex.Message}", ex);
                    }
                    finally
                    {
                        if (attachment != null)
                        {
                            Marshal.ReleaseComObject(attachment);
                        }
                    }
                }

                if (savedAttachments.Count > 0)
                {
                    StringBuilder attachmentSection = new StringBuilder();
                    attachmentSection.AppendLine("\n\n## Attachments\n");

                    foreach (SavedAttachment saved in savedAttachments)
                    {
                        string wikilink = _attachmentService.GenerateWikilink(
                            saved.FullPath,
                            noteFilePath,
                            saved.IsInline);
                        attachmentSection.AppendLine(wikilink);
                    }

                    _fileService.AppendToFile(noteFilePath, attachmentSection.ToString());
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"Failed to process appointment attachments: {ex.Message}", ex);
            }
            finally
            {
                if (attachments != null)
                {
                    Marshal.ReleaseComObject(attachments);
                }
            }
        }
    }
}
