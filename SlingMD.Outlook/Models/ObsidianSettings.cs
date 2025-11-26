using System;
using System.Configuration;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
        public int DefaultDueDays { get; set; } = 1;  // Due tomorrow
        /// <summary>
        /// If true, DefaultReminderDays represents days before the due date.
        /// If false, DefaultReminderDays represents days from now (absolute).
        /// </summary>
        public bool UseRelativeReminder { get; set; } = false;
        /// <summary>
        /// Gets or sets the number of days for the reminder.
        /// If UseRelativeReminder is true: represents days before the due date
        /// If UseRelativeReminder is false: represents days from now (absolute)
        /// </summary>
        public int DefaultReminderDays { get; set; } = 0;  // Remind today
        public int DefaultReminderHour { get; set; } = 9;  // at 9am
        public bool AskForDates { get; set; } = false;
        public bool GroupEmailThreads { get; set; } = true;
        public bool ShowDevelopmentSettings { get; set; } = false;
        public bool ShowThreadDebug { get; set; } = false;
        /// <summary>
        /// Whether to include the dailyNoteLink field in frontmatter.
        /// </summary>
        public bool IncludeDailyNoteLink { get; set; } = true;
        /// <summary>
        /// Format for the daily note link in frontmatter. Use {date:format} for date placeholders.
        /// Default: [[yyyy-MM-dd]] which produces links like [[2024-01-15]]
        /// </summary>
        public string DailyNoteLinkFormat { get; set; } = "[[yyyy-MM-dd]]";
        /// <summary>
        /// Default tags to apply to the note's frontmatter.
        /// Leave empty to not include any tags.
        /// </summary>
        public List<string> DefaultNoteTags { get; set; } = new List<string> { "FollowUp" };
        /// <summary>
        /// Default tags to apply to the Obsidian task (in the note body).
        /// </summary>
        public List<string> DefaultTaskTags { get; set; } = new List<string> { "FollowUp" };
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

        public List<string> SubjectCleanupPatterns { get; set; } = new List<string>
        {
            // Remove all variations of Re/Fwd prefixes, including multiple occurrences
            @"^(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])*",  // Matches one or more prefixes at start
            @"(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])+",   // Matches prefixes anywhere in string
            // Common email tags
            @"\[EXTERNAL\]\s*",             // External email tags
            @"\[Internal\]\s*",             // Internal email tags
            @"\[Confidential\]\s*",         // Confidential tags
            @"\[Secure\]\s*",               // Secure email tags
            @"\[Sensitive\]\s*",            // Sensitive email tags
            @"\[Private\]\s*",              // Private email tags
            @"\[PHI\]\s*",                  // PHI email tags
            @"\[Encrypted\]\s*",            // Encrypted email tags
            @"\[SPAM\]\s*",                 // Spam tags
            // Cleanup
            @"^\s+|\s+$",                   // Leading/trailing whitespace
            @"[-_\s]{2,}",                  // Multiple separators into single hyphen
            @"^-+|-+$"                      // Leading/trailing hyphens
        };

        public string GetFullVaultPath()
        {
            return System.IO.Path.Combine(VaultBasePath, VaultName);
        }

        public string GetInboxPath()
        {
            return System.IO.Path.Combine(GetFullVaultPath(), InboxFolder);
        }

        public string GetContactsPath()
        {
            return Path.Combine(GetFullVaultPath(), ContactsFolder);
        }

        /// <summary>
        /// Validates all settings before saving. Throws ArgumentException if any setting is invalid.
        /// </summary>
        private void Validate()
        {
            // Validate vault name
            if (string.IsNullOrWhiteSpace(VaultName))
            {
                throw new ArgumentException("Vault name cannot be empty.");
            }

            // Validate vault base path
            if (string.IsNullOrWhiteSpace(VaultBasePath))
            {
                throw new ArgumentException("Vault base path cannot be empty.");
            }

            // Validate folder names don't contain invalid path characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            if (InboxFolder != null && InboxFolder.IndexOfAny(invalidChars) >= 0)
            {
                throw new ArgumentException($"Inbox folder name contains invalid characters: {InboxFolder}");
            }
            if (ContactsFolder != null && ContactsFolder.IndexOfAny(invalidChars) >= 0)
            {
                throw new ArgumentException($"Contacts folder name contains invalid characters: {ContactsFolder}");
            }

            // Validate numeric ranges
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

            // Validate attachment folder name
            if (AttachmentsFolder != null && AttachmentsFolder.IndexOfAny(invalidChars) >= 0)
            {
                throw new ArgumentException($"Attachments folder name contains invalid characters: {AttachmentsFolder}");
            }

            // Validate regex patterns
            if (SubjectCleanupPatterns != null)
            {
                foreach (var pattern in SubjectCleanupPatterns)
                {
                    try
                    {
                        // Test if the pattern compiles
                        System.Text.RegularExpressions.Regex.IsMatch("test", pattern);
                    }
                    catch (ArgumentException ex)
                    {
                        throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}");
                    }
                }
            }
        }

        public void Save()
        {
            // Validate before saving
            Validate();

            var settings = new Dictionary<string, object>
            {
                { "VaultName", VaultName },
                { "VaultBasePath", VaultBasePath },
                { "InboxFolder", InboxFolder },
                { "ContactsFolder", ContactsFolder },
                { "EnableContactSaving", EnableContactSaving },
                { "SearchEntireVaultForContacts", SearchEntireVaultForContacts },
                { "LaunchObsidian", LaunchObsidian },
                { "ObsidianDelaySeconds", ObsidianDelaySeconds },
                { "ShowCountdown", ShowCountdown },
                { "CreateObsidianTask", CreateObsidianTask },
                { "CreateOutlookTask", CreateOutlookTask },
                { "DefaultDueDays", DefaultDueDays },
                { "UseRelativeReminder", UseRelativeReminder },
                { "DefaultReminderDays", DefaultReminderDays },
                { "DefaultReminderHour", DefaultReminderHour },
                { "AskForDates", AskForDates },
                { "SubjectCleanupPatterns", SubjectCleanupPatterns },
                { "GroupEmailThreads", GroupEmailThreads },
                { "ShowDevelopmentSettings", ShowDevelopmentSettings },
                { "ShowThreadDebug", ShowThreadDebug },
                { "IncludeDailyNoteLink", IncludeDailyNoteLink },
                { "DailyNoteLinkFormat", DailyNoteLinkFormat },
                { "DefaultNoteTags", DefaultNoteTags },
                { "DefaultTaskTags", DefaultTaskTags },
                { "NoteTitleFormat", NoteTitleFormat },
                { "NoteTitleMaxLength", NoteTitleMaxLength },
                { "NoteTitleIncludeDate", NoteTitleIncludeDate },
                { "MoveDateToFrontInThread", MoveDateToFrontInThread },
                { "AttachmentsFolder", AttachmentsFolder },
                { "AttachmentStorageMode", AttachmentStorageMode },
                { "SaveInlineImages", SaveInlineImages },
                { "SaveAllAttachments", SaveAllAttachments },
                { "UseObsidianWikilinks", UseObsidianWikilinks }
            };

            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            
            // Ensure settings directory exists before saving
            string settingsPath = GetSettingsPath();
            string settingsDir = Path.GetDirectoryName(settingsPath);
            if (!Directory.Exists(settingsDir))
            {
                Directory.CreateDirectory(settingsDir);
            }
            
            File.WriteAllText(settingsPath, json);
        }

        /// <summary>
        /// Helper method to load a setting value using reflection.
        /// Reduces code duplication from 124 lines to ~20 lines.
        /// </summary>
        private void LoadSetting<T>(Dictionary<string, JToken> settings, string key)
        {
            if (settings.ContainsKey(key))
            {
                var property = GetType().GetProperty(key);
                if (property != null && property.CanWrite)
                {
                    try
                    {
                        T value = settings[key].ToObject<T>();
                        property.SetValue(this, value);
                    }
                    catch
                    {
                        // Skip invalid values
                    }
                }
            }
        }

        public void Load()
        {
            if (File.Exists(GetSettingsPath()))
            {
                string json = File.ReadAllText(GetSettingsPath());
                var settings = JsonConvert.DeserializeObject<Dictionary<string, JToken>>(json);

                // Use reflection-based loading to reduce code duplication
                LoadSetting<string>(settings, "VaultName");
                LoadSetting<string>(settings, "VaultBasePath");
                LoadSetting<string>(settings, "InboxFolder");
                LoadSetting<string>(settings, "ContactsFolder");
                LoadSetting<bool>(settings, "EnableContactSaving");
                LoadSetting<bool>(settings, "SearchEntireVaultForContacts");
                LoadSetting<bool>(settings, "LaunchObsidian");
                LoadSetting<int>(settings, "ObsidianDelaySeconds");
                LoadSetting<bool>(settings, "ShowCountdown");
                LoadSetting<bool>(settings, "CreateObsidianTask");
                LoadSetting<bool>(settings, "CreateOutlookTask");
                LoadSetting<int>(settings, "DefaultDueDays");
                LoadSetting<bool>(settings, "UseRelativeReminder");
                LoadSetting<int>(settings, "DefaultReminderDays");
                LoadSetting<int>(settings, "DefaultReminderHour");
                LoadSetting<bool>(settings, "AskForDates");
                LoadSetting<bool>(settings, "GroupEmailThreads");
                LoadSetting<bool>(settings, "ShowDevelopmentSettings");
                LoadSetting<bool>(settings, "ShowThreadDebug");
                LoadSetting<bool>(settings, "IncludeDailyNoteLink");
                LoadSetting<string>(settings, "DailyNoteLinkFormat");
                LoadSetting<string>(settings, "NoteTitleFormat");
                LoadSetting<int>(settings, "NoteTitleMaxLength");
                LoadSetting<bool>(settings, "NoteTitleIncludeDate");
                LoadSetting<bool>(settings, "MoveDateToFrontInThread");
                LoadSetting<string>(settings, "AttachmentsFolder");
                LoadSetting<AttachmentStorageMode>(settings, "AttachmentStorageMode");
                LoadSetting<bool>(settings, "SaveInlineImages");
                LoadSetting<bool>(settings, "SaveAllAttachments");
                LoadSetting<bool>(settings, "UseObsidianWikilinks");

                // Load list properties with additional validation
                LoadSetting<List<string>>(settings, "SubjectCleanupPatterns");
                LoadSetting<List<string>>(settings, "DefaultNoteTags");
                LoadSetting<List<string>>(settings, "DefaultTaskTags");
            }
        }

        protected virtual string GetSettingsPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlingMD.Outlook", "ObsidianSettings.json");
        }
    }
} 