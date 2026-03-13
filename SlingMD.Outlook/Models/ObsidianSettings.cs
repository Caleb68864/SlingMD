using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace SlingMD.Outlook.Models
{
    public class ObsidianSettings
    {
        public string VaultName { get; set; } = "Logic";
        public string VaultBasePath { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Notes");
        public string InboxFolder { get; set; } = "Inbox";
        public string ContactsFolder { get; set; } = "Contacts";
        public bool EnableContactSaving { get; set; } = true;
        public bool SearchEntireVaultForContacts { get; set; } = false;
        public bool LaunchObsidian { get; set; } = true;
        public int ObsidianDelaySeconds { get; set; } = 1;
        public bool ShowCountdown { get; set; } = true;
        public bool CreateObsidianTask { get; set; } = true;
        public bool CreateOutlookTask { get; set; } = false;
        public int DefaultDueDays { get; set; } = 1;

        /// <summary>
        /// If true, DefaultReminderDays represents days before the due date.
        /// If false, DefaultReminderDays represents days from now (absolute).
        /// </summary>
        public bool UseRelativeReminder { get; set; } = false;

        /// <summary>
        /// Gets or sets the number of days for the reminder.
        /// If UseRelativeReminder is true: represents days before the due date.
        /// If UseRelativeReminder is false: represents days from now (absolute).
        /// </summary>
        public int DefaultReminderDays { get; set; } = 0;

        public int DefaultReminderHour { get; set; } = 9;
        public bool AskForDates { get; set; } = false;
        public bool GroupEmailThreads { get; set; } = true;
        public bool ShowDevelopmentSettings { get; set; } = false;
        public bool ShowThreadDebug { get; set; } = false;
        public bool HasShownSupportPrompt { get; set; } = false;

        /// <summary>
        /// Whether to include the dailyNoteLink field in frontmatter.
        /// </summary>
        public bool IncludeDailyNoteLink { get; set; } = true;

        /// <summary>
        /// Format for the daily note link in frontmatter.
        /// Default: [[yyyy-MM-dd]] which produces links like [[2024-01-15]].
        /// </summary>
        public string DailyNoteLinkFormat { get; set; } = "[[yyyy-MM-dd]]";

        /// <summary>
        /// Folder used when searching for user-provided templates.
        /// Relative paths are resolved from the Obsidian vault root.
        /// </summary>
        public string TemplatesFolder { get; set; } = "Templates";

        /// <summary>
        /// Template filename used for exported email notes.
        /// </summary>
        public string EmailTemplateFile { get; set; } = "EmailTemplate.md";

        /// <summary>
        /// Template filename used for contact notes.
        /// </summary>
        public string ContactTemplateFile { get; set; } = "ContactTemplate.md";

        /// <summary>
        /// Template filename used for inline Obsidian task lines.
        /// </summary>
        public string TaskTemplateFile { get; set; } = "TaskTemplate.md";

        /// <summary>
        /// Template filename used for thread summary notes.
        /// </summary>
        public string ThreadTemplateFile { get; set; } = "ThreadNoteTemplate.md";

        /// <summary>
        /// Template filename used for appointment notes.
        /// </summary>
        public string AppointmentTemplateFile { get; set; } = "AppointmentTemplate.md";

        /// <summary>
        /// Template filename used for meeting note stubs.
        /// </summary>
        public string MeetingNoteTemplateFile { get; set; } = "MeetingNoteTemplate.md";

        /// <summary>
        /// Optional filename format for email notes. Leave blank to preserve the legacy naming behavior.
        /// Supported tokens include {Subject}, {Sender}, {Date}, and {Timestamp}.
        /// </summary>
        public string EmailFilenameFormat { get; set; } = string.Empty;

        /// <summary>
        /// Filename format for contact notes.
        /// Supported tokens include {ContactName} and {ContactShortName}.
        /// </summary>
        public string ContactFilenameFormat { get; set; } = "{ContactName}";

        /// <summary>
        /// Whether to include contact details (phone, email, company, etc.) in the contact note.
        /// </summary>
        public bool ContactNoteIncludeDetails { get; set; } = true;

        /// <summary>
        /// Default tags to apply to the note's frontmatter.
        /// Leave empty to not include any tags.
        /// </summary>
        public List<string> DefaultNoteTags { get; set; } = CreateDefaultNoteTags();

        /// <summary>
        /// Default tags to apply to the Obsidian task (in the note body).
        /// </summary>
        public List<string> DefaultTaskTags { get; set; } = CreateDefaultTaskTags();

        /// <summary>
        /// Format for the note title. Use placeholders: {Subject}, {Sender}, {Date}.
        /// </summary>
        public string NoteTitleFormat { get; set; } = "{Subject} - {Date}";

        /// <summary>
        /// Maximum length for the note title. Titles longer than this will be trimmed with ellipsis.
        /// </summary>
        public int NoteTitleMaxLength { get; set; } = 50;

        /// <summary>
        /// Whether to include the date in the note title.
        /// </summary>
        public bool NoteTitleIncludeDate { get; set; } = true;

        public bool MoveDateToFrontInThread { get; set; } = true;

        /// <summary>
        /// Folder name for centralized attachment storage (relative to vault root).
        /// </summary>
        public string AttachmentsFolder { get; set; } = "Attachments";

        /// <summary>
        /// Determines where attachments are stored (same folder, subfolder per note, or centralized).
        /// </summary>
        public AttachmentStorageMode AttachmentStorageMode { get; set; } = AttachmentStorageMode.SameAsNote;

        /// <summary>
        /// Whether to save inline images from emails.
        /// </summary>
        public bool SaveInlineImages { get; set; } = true;

        /// <summary>
        /// Whether to save all email attachments (not just inline images).
        /// </summary>
        public bool SaveAllAttachments { get; set; } = false;

        /// <summary>
        /// Whether to use Obsidian wikilinks (![[image.png]]) or standard markdown (![image.png](image.png)).
        /// </summary>
        public bool UseObsidianWikilinks { get; set; } = true;

        // Appointment Settings

        /// <summary>
        /// Folder name for appointment notes (relative to vault root).
        /// </summary>
        public string AppointmentsFolder { get; set; } = "Appointments";

        /// <summary>
        /// Format for appointment note titles. Supported tokens: {Date}, {Subject}, {Sender}.
        /// </summary>
        public string AppointmentNoteTitleFormat { get; set; } = "{Date} - {Subject}";

        /// <summary>
        /// Maximum length for appointment note titles.
        /// </summary>
        public int AppointmentNoteTitleMaxLength { get; set; } = 50;

        /// <summary>
        /// Default tags to apply to appointment note frontmatter.
        /// </summary>
        public List<string> AppointmentDefaultNoteTags { get; set; } = new List<string> { "Appointment" };

        /// <summary>
        /// Whether to save attachments from appointment items.
        /// </summary>
        public bool AppointmentSaveAttachments { get; set; } = true;

        /// <summary>
        /// Whether to create a meeting notes section in appointment notes.
        /// </summary>
        public bool CreateMeetingNotes { get; set; } = true;

        /// <summary>
        /// Optional custom template path for meeting notes. Leave empty to use the default template.
        /// </summary>
        public string MeetingNoteTemplate { get; set; } = string.Empty;

        /// <summary>
        /// Whether to group recurring meeting occurrences into a shared folder.
        /// </summary>
        public bool GroupRecurringMeetings { get; set; } = true;

        /// <summary>
        /// Whether to save notes for cancelled appointments.
        /// </summary>
        public bool SaveCancelledAppointments { get; set; } = false;

        /// <summary>
        /// Task creation mode for appointments. Valid values: "None", "Obsidian", "Outlook", "Both".
        /// </summary>
        public string AppointmentTaskCreation { get; set; } = "None";

        public List<string> SubjectCleanupPatterns { get; set; } = CreateDefaultSubjectCleanupPatterns();

        private static List<string> CreateDefaultNoteTags()
        {
            return new List<string> { "FollowUp" };
        }

        private static List<string> CreateDefaultTaskTags()
        {
            return new List<string> { "FollowUp" };
        }

        private static List<string> CreateDefaultSubjectCleanupPatterns()
        {
            return new List<string>
            {
                @"^(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])*",
                @"(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])+",
                @"\[EXTERNAL\]\s*",
                @"\[Internal\]\s*",
                @"\[Confidential\]\s*",
                @"\[Secure\]\s*",
                @"\[Sensitive\]\s*",
                @"\[Private\]\s*",
                @"\[PHI\]\s*",
                @"\[Encrypted\]\s*",
                @"\[SPAM\]\s*",
                @"^\s+|\s+$",
                @"[-_\s]{2,}",
                @"^-+|-+$"
            };
        }

        public string GetFullVaultPath()
        {
            return Path.Combine(VaultBasePath, VaultName);
        }

        public string GetInboxPath()
        {
            return Path.Combine(GetFullVaultPath(), InboxFolder);
        }

        public string GetContactsPath()
        {
            return Path.Combine(GetFullVaultPath(), ContactsFolder);
        }

        public string GetAppointmentsPath()
        {
            return Path.Combine(GetFullVaultPath(), AppointmentsFolder);
        }

        public bool HasSavedSettings()
        {
            return File.Exists(GetSettingsPath());
        }

        public string GetTemplatesPath()
        {
            if (Path.IsPathRooted(TemplatesFolder))
            {
                return TemplatesFolder;
            }

            return Path.Combine(GetFullVaultPath(), TemplatesFolder);
        }

        /// <summary>
        /// Validates all settings before saving. Throws ArgumentException if any setting is invalid.
        /// </summary>
        private void Validate()
        {
            if (string.IsNullOrWhiteSpace(VaultName))
            {
                throw new ArgumentException("Vault name cannot be empty.");
            }

            if (string.IsNullOrWhiteSpace(VaultBasePath))
            {
                throw new ArgumentException("Vault base path cannot be empty.");
            }

            char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
            ValidateFolderName(InboxFolder, "Inbox folder", invalidFileNameChars);
            ValidateFolderName(ContactsFolder, "Contacts folder", invalidFileNameChars);
            ValidateOptionalRelativePathSegment(TemplatesFolder, "Templates folder", invalidFileNameChars);
            ValidateTemplateFileName(EmailTemplateFile, "Email template file", invalidFileNameChars);
            ValidateTemplateFileName(ContactTemplateFile, "Contact template file", invalidFileNameChars);
            ValidateTemplateFileName(TaskTemplateFile, "Task template file", invalidFileNameChars);
            ValidateTemplateFileName(ThreadTemplateFile, "Thread template file", invalidFileNameChars);
            ValidateFolderName(AttachmentsFolder, "Attachments folder", invalidFileNameChars);

            if (ObsidianDelaySeconds < 0 || ObsidianDelaySeconds > 60)
            {
                throw new ArgumentException("Obsidian delay must be between 0 and 60 seconds.");
            }

            if (DefaultDueDays < 0 || DefaultDueDays > 365)
            {
                throw new ArgumentException("Default due days must be between 0 and 365.");
            }

            if (DefaultReminderDays < 0 || DefaultReminderDays > 365)
            {
                throw new ArgumentException("Default reminder days must be between 0 and 365.");
            }

            if (DefaultReminderHour < 0 || DefaultReminderHour > 23)
            {
                throw new ArgumentException("Default reminder hour must be between 0 and 23.");
            }

            if (NoteTitleMaxLength < 10 || NoteTitleMaxLength > 500)
            {
                throw new ArgumentException("Note title max length must be between 10 and 500.");
            }

            if (SubjectCleanupPatterns != null)
            {
                foreach (string pattern in SubjectCleanupPatterns)
                {
                    try
                    {
                        System.Text.RegularExpressions.Regex.IsMatch("test", pattern);
                    }
                    catch (ArgumentException ex)
                    {
                        throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}");
                    }
                }
            }

            ValidateFolderName(AppointmentsFolder, "Appointments folder", invalidFileNameChars);

            if (AppointmentNoteTitleMaxLength < 10 || AppointmentNoteTitleMaxLength > 500)
            {
                throw new ArgumentException("Appointment note title max length must be between 10 and 500.");
            }

            string[] validTaskCreationValues = { "None", "Obsidian", "Outlook", "Both" };
            if (!System.Array.Exists(validTaskCreationValues, v => v == AppointmentTaskCreation))
            {
                throw new ArgumentException($"Invalid AppointmentTaskCreation value: {AppointmentTaskCreation}. Must be one of: None, Obsidian, Outlook, Both.");
            }
        }

        public void Save()
        {
            Validate();

            string json = JsonConvert.SerializeObject(this, Formatting.Indented);
            string settingsPath = GetSettingsPath();
            string settingsDir = Path.GetDirectoryName(settingsPath);
            if (!Directory.Exists(settingsDir))
            {
                Directory.CreateDirectory(settingsDir);
            }

            File.WriteAllText(settingsPath, json);
        }

        public void Load()
        {
            string settingsPath = GetSettingsPath();
            if (!File.Exists(settingsPath))
            {
                NormalizeLoadedSettings();
                return;
            }

            try
            {
                string json = File.ReadAllText(settingsPath);
                JsonConvert.PopulateObject(json, this, CreateJsonSerializerSettings());
            }
            catch (JsonException)
            {
                // Keep defaults if the saved file is malformed.
            }

            NormalizeLoadedSettings();
        }

        private static JsonSerializerSettings CreateJsonSerializerSettings()
        {
            return new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Ignore,
                ObjectCreationHandling = ObjectCreationHandling.Replace,
                Error = (sender, args) =>
                {
                    args.ErrorContext.Handled = true;
                }
            };
        }

        private void NormalizeLoadedSettings()
        {
            string defaultVaultBasePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Notes");

            VaultName = string.IsNullOrWhiteSpace(VaultName) ? "Logic" : VaultName;
            VaultBasePath = string.IsNullOrWhiteSpace(VaultBasePath) ? defaultVaultBasePath : VaultBasePath;
            InboxFolder = string.IsNullOrWhiteSpace(InboxFolder) ? "Inbox" : InboxFolder;
            ContactsFolder = string.IsNullOrWhiteSpace(ContactsFolder) ? "Contacts" : ContactsFolder;
            TemplatesFolder = string.IsNullOrWhiteSpace(TemplatesFolder) ? "Templates" : TemplatesFolder;
            EmailTemplateFile = string.IsNullOrWhiteSpace(EmailTemplateFile) ? "EmailTemplate.md" : EmailTemplateFile;
            ContactTemplateFile = string.IsNullOrWhiteSpace(ContactTemplateFile) ? "ContactTemplate.md" : ContactTemplateFile;
            TaskTemplateFile = string.IsNullOrWhiteSpace(TaskTemplateFile) ? "TaskTemplate.md" : TaskTemplateFile;
            ThreadTemplateFile = string.IsNullOrWhiteSpace(ThreadTemplateFile) ? "ThreadNoteTemplate.md" : ThreadTemplateFile;
            ContactFilenameFormat = string.IsNullOrWhiteSpace(ContactFilenameFormat) ? "{ContactName}" : ContactFilenameFormat;
            SubjectCleanupPatterns = SubjectCleanupPatterns ?? CreateDefaultSubjectCleanupPatterns();
            DefaultNoteTags = DefaultNoteTags ?? CreateDefaultNoteTags();
            DefaultTaskTags = DefaultTaskTags ?? CreateDefaultTaskTags();
            NoteTitleFormat = string.IsNullOrWhiteSpace(NoteTitleFormat) ? "{Subject} - {Date}" : NoteTitleFormat;
            DailyNoteLinkFormat = string.IsNullOrWhiteSpace(DailyNoteLinkFormat) ? "[[yyyy-MM-dd]]" : DailyNoteLinkFormat;
            AttachmentsFolder = string.IsNullOrWhiteSpace(AttachmentsFolder) ? "Attachments" : AttachmentsFolder;
            EmailFilenameFormat = EmailFilenameFormat ?? string.Empty;

            if (!Enum.IsDefined(typeof(AttachmentStorageMode), AttachmentStorageMode))
            {
                AttachmentStorageMode = AttachmentStorageMode.SameAsNote;
            }

            AppointmentsFolder = string.IsNullOrWhiteSpace(AppointmentsFolder) ? "Appointments" : AppointmentsFolder;
            AppointmentNoteTitleFormat = string.IsNullOrWhiteSpace(AppointmentNoteTitleFormat) ? "{Date} - {Subject}" : AppointmentNoteTitleFormat;
            AppointmentDefaultNoteTags = AppointmentDefaultNoteTags ?? new List<string> { "Appointment" };
            MeetingNoteTemplate = MeetingNoteTemplate ?? string.Empty;
            AppointmentTemplateFile = string.IsNullOrWhiteSpace(AppointmentTemplateFile) ? "AppointmentTemplate.md" : AppointmentTemplateFile;
            MeetingNoteTemplateFile = string.IsNullOrWhiteSpace(MeetingNoteTemplateFile) ? "MeetingNoteTemplate.md" : MeetingNoteTemplateFile;

            string[] validTaskCreationValues = { "None", "Obsidian", "Outlook", "Both" };
            if (string.IsNullOrWhiteSpace(AppointmentTaskCreation) || !System.Array.Exists(validTaskCreationValues, v => v == AppointmentTaskCreation))
            {
                AppointmentTaskCreation = "None";
            }
        }

        private static void ValidateFolderName(string value, string label, char[] invalidChars)
        {
            if (!string.IsNullOrWhiteSpace(value) && value.IndexOfAny(invalidChars) >= 0)
            {
                throw new ArgumentException($"{label} contains invalid characters: {value}");
            }
        }

        private static void ValidateOptionalRelativePathSegment(string value, string label, char[] invalidChars)
        {
            if (string.IsNullOrWhiteSpace(value) || Path.IsPathRooted(value))
            {
                return;
            }

            string[] pathSegments = value.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string pathSegment in pathSegments)
            {
                if (pathSegment == "." || pathSegment == "..")
                {
                    throw new ArgumentException($"{label} cannot contain relative traversal segments: {value}");
                }

                if (pathSegment.IndexOfAny(invalidChars) >= 0)
                {
                    throw new ArgumentException($"{label} contains invalid characters: {value}");
                }
            }
        }

        private static void ValidateTemplateFileName(string value, string label, char[] invalidChars)
        {
            if (string.IsNullOrWhiteSpace(value) || Path.IsPathRooted(value))
            {
                return;
            }

            string fileName = Path.GetFileName(value);
            if (string.IsNullOrWhiteSpace(fileName) || fileName.IndexOfAny(invalidChars) >= 0)
            {
                throw new ArgumentException($"{label} contains invalid characters: {value}");
            }
        }

        protected virtual string GetSettingsPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlingMD.Outlook", "ObsidianSettings.json");
        }
    }
}


