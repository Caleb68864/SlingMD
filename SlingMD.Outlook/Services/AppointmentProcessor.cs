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
        private readonly DateFormatter _dateFormatter;
        private readonly ContactNameParser _contactNameParser;
        private readonly ContactLinkFormatter _contactLinkFormatter;
        private readonly SubjectFilenameCleaner _subjectFilenameCleaner;
        private readonly NoteTitleBuilder _noteTitleBuilder;
        private readonly UniqueFilenameResolver _uniqueFilenameResolver = new UniqueFilenameResolver();
        private readonly IClock _clock;
        private readonly ReminderDueDateCalculator _reminderCalculator;

        private List<string> _bulkErrors = new List<string>();

        public List<string> GetBulkErrors()
        {
            List<string> errors = new List<string>(_bulkErrors);
            _bulkErrors.Clear();
            return errors;
        }

        public AppointmentProcessor(ObsidianSettings settings, IClock clock = null)
        {
            _settings = settings;
            _fileService = new FileService(settings);
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService, settings);
            _clock = clock ?? new SystemClock();
            _taskService = new TaskService(settings, _templateService, _clock);
            _contactService = new ContactService(_fileService, _templateService);
            _attachmentService = new AttachmentService(settings, _fileService);
            _dateFormatter = new DateFormatter();
            _contactNameParser = new ContactNameParser();
            _contactLinkFormatter = new ContactLinkFormatter();
            _subjectFilenameCleaner = new SubjectFilenameCleaner(settings, _fileService);
            _noteTitleBuilder = new NoteTitleBuilder();
            _reminderCalculator = new ReminderDueDateCalculator();
        }

        private string FormatPersonLink(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }
            ContactName parsed = _contactNameParser.Parse(name, null);
            string formatted = _contactLinkFormatter.Format(parsed, _settings.ContactLinkFormat);
            return string.IsNullOrEmpty(formatted) ? $"[[{name}]]" : formatted;
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
            string subject = string.Empty;
            string body = string.Empty;
            string location = string.Empty;
            DateTime startTime = _clock.Now;
            DateTime endTime = _clock.Now;
            string organizerName = string.Empty;
            string organizerEmail = string.Empty;
            string globalAppointmentId = string.Empty;
            OlRecurrenceState recurrenceState = OlRecurrenceState.olApptNotRecurring;
            List<string> requiredAttendees = new List<string>();
            List<string> optionalAttendees = new List<string>();
            List<string> resourceAttendees = new List<string>();

            try { subject = appointment.Subject ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Subject: {ex.Message}"); }

            try { body = appointment.Body ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Body: {ex.Message}"); }

            try { location = appointment.Location ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Location: {ex.Message}"); }

            try { startTime = appointment.Start; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read Start: {ex.Message}"); }

            try { endTime = appointment.End; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read End: {ex.Message}"); }

            try { recurrenceState = appointment.RecurrenceState; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read RecurrenceState: {ex.Message}"); }

            try { globalAppointmentId = appointment.GlobalAppointmentID ?? string.Empty; }
            catch (System.Exception ex) { Logger.Instance.Debug($"Could not read GlobalAppointmentID: {ex.Message}"); }

            // Extract organizer via AddressEntry COM object
            AddressEntry organizerEntry = null;
            try
            {
                organizerEntry = appointment.GetOrganizer();
                if (organizerEntry != null)
                {
                    try { organizerName = organizerEntry.Name ?? string.Empty; }
                    catch (System.Exception ex) { Logger.Instance.Debug($"Could not read organizer Name: {ex.Message}"); }

                    try { organizerEmail = organizerEntry.Address ?? string.Empty; }
                    catch (System.Exception ex) { Logger.Instance.Debug($"Could not read organizer Address: {ex.Message}"); }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not get organizer: {ex.Message}");
            }
            finally
            {
                if (organizerEntry != null)
                {
                    Marshal.ReleaseComObject(organizerEntry);
                }
            }

            // Extract recipients by type
            Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        string recipientName = string.Empty;
                        OlMeetingRecipientType recipientType = OlMeetingRecipientType.olRequired;

                        try { recipientName = recipient.Name ?? string.Empty; }
                        catch (System.Exception ex) { Logger.Instance.Debug($"Could not read recipient Name: {ex.Message}"); }

                        try { recipientType = (OlMeetingRecipientType)recipient.Type; }
                        catch (System.Exception ex) { Logger.Instance.Debug($"Could not read recipient Type: {ex.Message}"); }

                        if (!string.IsNullOrWhiteSpace(recipientName))
                        {
                            if (recipientType == OlMeetingRecipientType.olOptional)
                            {
                                optionalAttendees.Add(recipientName);
                            }
                            else if (recipientType == OlMeetingRecipientType.olResource)
                            {
                                resourceAttendees.Add(recipientName);
                            }
                            else
                            {
                                requiredAttendees.Add(recipientName);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Debug($"Could not process recipient: {ex.Message}");
                    }
                    finally
                    {
                        if (recipient != null)
                        {
                            Marshal.ReleaseComObject(recipient);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"Could not read Recipients: {ex.Message}");
            }
            finally
            {
                if (recipients != null)
                {
                    Marshal.ReleaseComObject(recipients);
                }
            }

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
                    using (Forms.TaskOptionsForm form = new Forms.TaskOptionsForm(
                        _settings.DefaultDueDays, _settings.DefaultReminderDays,
                        _settings.DefaultReminderHour, _settings.UseRelativeReminder))
                    {
                        System.Windows.Forms.DialogResult result = form.ShowDialog();
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            _taskService.InitializeTaskSettings(
                                form.DueDays, form.ReminderDays,
                                form.ReminderHour, form.UseRelativeReminder);
                        }
                        else
                        {
                            _taskService.DisableTaskCreation();
                            createObsidianTask = false;
                            createOutlookTask = false;
                        }
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
            string organizerShortName = _contactService.GetShortName(
                string.IsNullOrWhiteSpace(organizerName) ? "Unknown Organizer" : organizerName);

            string titleFormat = _settings.AppointmentNoteTitleFormat ?? "{Date} - {Subject}";
            int maxLength = _settings.AppointmentNoteTitleMaxLength > 0 ? _settings.AppointmentNoteTitleMaxLength : 50;

            Dictionary<string, string> titleTokens = new Dictionary<string, string>
            {
                { "Date", dateStr },
                { "Subject", subjectClean },
                { "Sender", organizerShortName }
            };
            string noteTitle = _noteTitleBuilder.BuildTrimmed(titleFormat, titleTokens, maxLength);

            string fileNameNoExt = _fileService.CleanFileName(noteTitle);
            if (string.IsNullOrWhiteSpace(fileNameNoExt))
            {
                fileNameNoExt = $"Appointment-{dateStr}";
            }

            string fileName = fileNameNoExt + ".md";
            string appointmentsPath = _settings.GetAppointmentsPath();
            string filePath = Path.Combine(appointmentsPath, fileName);

            // --- Determine whether this instance goes into a recurring-series thread folder ---
            bool isRecurring = ShouldGroupAsRecurring(recurrenceState);
            string recurringThreadFolderPath = null;
            string recurringThreadFolderName = null;
            string recurringInstanceFilePath = null;
            string recurringThreadNotePath = null;

            if (isRecurring)
            {
                recurringThreadFolderName = GetRecurringThreadFolderName(appointment, subjectClean);
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
            Dictionary<string, object> metadata = BuildAppointmentMetadata(
                noteTitle,
                organizerName,
                organizerEmail,
                requiredAttendees,
                optionalAttendees,
                resourceAttendees,
                location,
                startTime,
                endTime,
                recurrenceState,
                globalAppointmentId,
                fileNameNoExt);

            // --- Generate Obsidian task block ---
            string taskBlock = string.Empty;
            if (createObsidianTask && _taskService.ShouldCreateTasks)
            {
                List<string> taskTags = _settings.AppointmentDefaultNoteTags ?? new List<string>();
                taskBlock = _taskService.GenerateObsidianTask(fileNameNoExt, taskTags);
            }

            // --- Render note content ---
            string renderedContent = BuildAppointmentNoteContent(
                metadata,
                noteTitle,
                subjectClean,
                organizerName,
                organizerShortName,
                organizerEmail,
                dateStr,
                startTime,
                endTime,
                location,
                body,
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
                using (StatusService status = new StatusService())
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

                        // In single mode, show dialog for new contacts; in bulk mode, skip silently
                        if (!bulkMode && newContacts.Count > 0)
                        {
                            using (ContactConfirmationDialog dialog = new ContactConfirmationDialog(newContacts))
                            {
                                if (dialog.ShowDialog() == DialogResult.OK)
                                {
                                    foreach (string contactName in dialog.SelectedContacts)
                                    {
                                        _contactService.CreateContactNote(contactName);
                                    }
                                }
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

                    _fileService.LaunchObsidian(_settings.VaultName, fileNameNoExt);
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

            return AppointmentProcessingResult.Success;
        }

        // --- Recurring meeting helpers ---

        private bool ShouldGroupAsRecurring(OlRecurrenceState recurrenceState)
        {
            if (!_settings.GroupRecurringMeetings)
            {
                return false;
            }

            return recurrenceState == OlRecurrenceState.olApptOccurrence
                || recurrenceState == OlRecurrenceState.olApptException;
        }

        private string GetRecurringThreadFolderName(AppointmentItem appointment, string cleanedSubject)
        {
            RecurrencePattern pattern = null;
            try
            {
                pattern = appointment.GetRecurrencePattern();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Fall through to use cleaned subject only
            }
            finally
            {
                if (pattern != null)
                {
                    Marshal.ReleaseComObject(pattern);
                }
            }

            return _fileService.CleanFileName(cleanedSubject);
        }

        private void WriteRecurringThreadSummary(
            string threadFolderPath,
            string threadFolderName,
            string threadNotePath,
            string title)
        {
            try
            {
                ThreadTemplateContext threadContext = new ThreadTemplateContext
                {
                    Title = title,
                    ThreadId = threadFolderName,
                    FolderPath = threadFolderPath
                };

                string threadContent = _templateService.RenderThreadContent(threadContext);
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
                string meetingNoteFileName = appointmentFileNameNoExt + " - Meeting Notes.md";
                string meetingNotePath = Path.Combine(outputFolder, meetingNoteFileName);

                if (File.Exists(meetingNotePath))
                {
                    return;
                }

                string meetingNoteTitleNoExt = appointmentFileNameNoExt + " - Meeting Notes";
                string appointmentLink = $"[[{appointmentFileNameNoExt}]]";

                Dictionary<string, object> meetingMetadata = new Dictionary<string, object>
                {
                    { "title", meetingNoteTitleNoExt },
                    { "type", "Meeting Notes" },
                    { "appointment", appointmentLink },
                    { "date", date }
                };

                List<string> noteTags = _settings.AppointmentDefaultNoteTags;
                if (noteTags != null && noteTags.Count > 0)
                {
                    meetingMetadata.Add("tags", new List<string>(noteTags));
                }

                MeetingNoteTemplateContext context = new MeetingNoteTemplateContext
                {
                    Metadata = meetingMetadata,
                    AppointmentTitle = appointmentNoteTitle,
                    AppointmentLink = appointmentLink,
                    Organizer = organizer ?? string.Empty,
                    Attendees = attendees ?? string.Empty,
                    Date = date ?? string.Empty,
                    Location = location ?? string.Empty
                };

                string meetingNoteContent = _templateService.RenderMeetingNoteContent(context);
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

        private Dictionary<string, object> BuildAppointmentMetadata(
            string noteTitle,
            string organizerName,
            string organizerEmail,
            List<string> requiredAttendees,
            List<string> optionalAttendees,
            List<string> resourceAttendees,
            string location,
            DateTime startTime,
            DateTime endTime,
            OlRecurrenceState recurrenceState,
            string globalAppointmentId,
            string fileNameNoExt = null)
        {
            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", noteTitle },
                { "type", "Appointment" },
                { "organizer", FormatPersonLink(organizerName) },
                { "organizerEmail", organizerEmail ?? string.Empty },
                { "location", location ?? string.Empty },
                { "startDateTime", _dateFormatter.Format(startTime, _settings.AppointmentDateFormat) },
                { "endDateTime", _dateFormatter.Format(endTime, _settings.AppointmentDateFormat) },
                { "recurrence", recurrenceState.ToString() },
                { "globalAppointmentId", globalAppointmentId ?? string.Empty }
            };

            if (_settings.CreateMeetingNotes && !string.IsNullOrWhiteSpace(fileNameNoExt))
            {
                string meetingNotesTitle = fileNameNoExt + " - Meeting Notes";
                metadata.Add("meetingNotes", $"[[{meetingNotesTitle}]]");
            }

            if (requiredAttendees != null && requiredAttendees.Count > 0)
            {
                List<string> linkedRequired = requiredAttendees
                    .Select(name => FormatPersonLink(name))
                    .ToList();
                metadata.Add("attendees", linkedRequired);
            }

            if (optionalAttendees != null && optionalAttendees.Count > 0)
            {
                List<string> linkedOptional = optionalAttendees
                    .Select(name => FormatPersonLink(name))
                    .ToList();
                metadata.Add("optionalAttendees", linkedOptional);
            }

            if (resourceAttendees != null && resourceAttendees.Count > 0)
            {
                List<string> linkedResources = resourceAttendees
                    .Select(name => FormatPersonLink(name))
                    .ToList();
                metadata.Add("resources", linkedResources);
            }

            if (_settings.IncludeDailyNoteLink)
            {
                string dailyLinkFormat = _settings.DailyNoteLinkFormat ?? "[[yyyy-MM-dd]]";
                string innerFormat = dailyLinkFormat.Replace("[[", string.Empty).Replace("]]", string.Empty);
                string dailyNoteLink = startTime.ToString(innerFormat);
                metadata.Add("dailyNoteLink", "[[" + dailyNoteLink + "]]");
            }

            List<string> noteTags = _settings.AppointmentDefaultNoteTags;
            if (noteTags != null && noteTags.Count > 0)
            {
                metadata.Add("tags", new List<string>(noteTags));
            }

            return metadata;
        }

        private string BuildAppointmentNoteContent(
            Dictionary<string, object> metadata,
            string noteTitle,
            string subjectClean,
            string organizerName,
            string organizerShortName,
            string organizerEmail,
            string dateStr,
            DateTime startTime,
            DateTime endTime,
            string location,
            string body,
            string fileNameNoExt,
            string taskBlock = "")
        {
            string frontMatter = _templateService.BuildFrontMatter(metadata);

            StringBuilder content = new StringBuilder();
            content.Append(frontMatter);
            content.AppendLine($"# {noteTitle}");
            content.AppendLine();

            if (!string.IsNullOrWhiteSpace(taskBlock))
            {
                content.AppendLine(taskBlock);
                content.AppendLine();
            }

            if (!string.IsNullOrWhiteSpace(location))
            {
                content.AppendLine($"**Location:** {location}");
                content.AppendLine();
            }

            content.AppendLine($"**Start:** {_dateFormatter.Format(startTime, _settings.AppointmentDateFormat)}");
            content.AppendLine($"**End:** {_dateFormatter.Format(endTime, _settings.AppointmentDateFormat)}");
            content.AppendLine();

            if (!string.IsNullOrWhiteSpace(organizerName))
            {
                content.AppendLine($"**Organizer:** {FormatPersonLink(organizerName)}");
                content.AppendLine();
            }

            if (!string.IsNullOrWhiteSpace(body))
            {
                content.AppendLine("## Description");
                content.AppendLine();
                content.AppendLine(body.Trim());
                content.AppendLine();
            }

            if (_settings.CreateMeetingNotes)
            {
                content.AppendLine("## Meeting Notes");
                content.AppendLine();
            }

            return content.ToString();
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
