using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// A single help entry describing one setting. Entries are the single source of truth
    /// for both hover tooltips in <see cref="SettingsForm"/> and the searchable browser
    /// rendered by <see cref="HelpForm"/>.
    /// </summary>
    public class HelpEntry
    {
        /// <summary>Unique identifier, e.g. "Contacts.ContactLinkFormat".</summary>
        public string Id { get; set; }

        /// <summary>Tab the setting lives on.</summary>
        public string Tab { get; set; }

        /// <summary>Human-readable title shown in tooltips and the help tree.</summary>
        public string Title { get; set; }

        /// <summary>One-line purpose — the first line shown in the tooltip.</summary>
        public string Summary { get; set; }

        /// <summary>Longer free-form description. Displayed under Summary in both surfaces.</summary>
        public string Description { get; set; }

        /// <summary>Optional token reference: name → explanation.</summary>
        public Dictionary<string, string> Tokens { get; set; }

        /// <summary>Optional list of worked examples (format string → rendered output pairs).</summary>
        public List<HelpExample> Examples { get; set; }

        /// <summary>Optional default value as a string.</summary>
        public string Default { get; set; }
    }

    public class HelpExample
    {
        public string Input { get; set; }
        public string Output { get; set; }
    }

    /// <summary>
    /// Central registry of help content for every setting exposed in <see cref="SettingsForm"/>.
    /// To add help for a new setting: add one <see cref="HelpEntry"/> here and reference it
    /// from the form via <see cref="Get(string)"/>.
    /// </summary>
    public static class SettingsHelp
    {
        private static readonly Dictionary<string, HelpEntry> _entries = BuildEntries();

        public static HelpEntry Get(string id)
        {
            HelpEntry entry;
            return _entries.TryGetValue(id, out entry) ? entry : null;
        }

        public static IEnumerable<HelpEntry> All()
        {
            return _entries.Values;
        }

        /// <summary>Case-insensitive substring search across title, summary, description, tokens, examples.</summary>
        public static IEnumerable<HelpEntry> Search(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return _entries.Values;
            }

            string q = query.Trim().ToLowerInvariant();
            List<HelpEntry> matches = new List<HelpEntry>();
            foreach (HelpEntry entry in _entries.Values)
            {
                if (Matches(entry, q))
                {
                    matches.Add(entry);
                }
            }
            return matches;
        }

        private static bool Matches(HelpEntry entry, string q)
        {
            if (Contains(entry.Title, q)) return true;
            if (Contains(entry.Summary, q)) return true;
            if (Contains(entry.Description, q)) return true;
            if (Contains(entry.Tab, q)) return true;
            if (Contains(entry.Default, q)) return true;
            if (entry.Tokens != null)
            {
                foreach (KeyValuePair<string, string> token in entry.Tokens)
                {
                    if (Contains(token.Key, q) || Contains(token.Value, q)) return true;
                }
            }
            if (entry.Examples != null)
            {
                foreach (HelpExample example in entry.Examples)
                {
                    if (Contains(example.Input, q) || Contains(example.Output, q)) return true;
                }
            }
            return false;
        }

        private static bool Contains(string haystack, string needle)
        {
            return !string.IsNullOrEmpty(haystack) && haystack.ToLowerInvariant().Contains(needle);
        }

        /// <summary>Renders a help entry as a multi-line tooltip string.</summary>
        public static string FormatAsTooltip(HelpEntry entry)
        {
            if (entry == null) return string.Empty;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(entry.Title);
            sb.AppendLine();
            if (!string.IsNullOrWhiteSpace(entry.Summary))
            {
                sb.AppendLine(entry.Summary);
            }
            if (!string.IsNullOrWhiteSpace(entry.Description))
            {
                sb.AppendLine();
                sb.AppendLine(entry.Description);
            }
            if (!string.IsNullOrWhiteSpace(entry.Default))
            {
                sb.AppendLine();
                sb.AppendLine("Default: " + entry.Default);
            }
            if (entry.Tokens != null && entry.Tokens.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Tokens:");
                int pad = entry.Tokens.Keys.Max(k => k.Length);
                foreach (KeyValuePair<string, string> token in entry.Tokens)
                {
                    sb.AppendLine("  " + token.Key.PadRight(pad) + "  " + token.Value);
                }
            }
            if (entry.Examples != null && entry.Examples.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Examples:");
                int pad = entry.Examples.Max(e => (e.Input ?? string.Empty).Length);
                foreach (HelpExample example in entry.Examples)
                {
                    sb.AppendLine("  " + (example.Input ?? string.Empty).PadRight(pad) + "  → " + example.Output);
                }
            }
            return sb.ToString().TrimEnd();
        }

        private static Dictionary<string, HelpEntry> BuildEntries()
        {
            List<HelpEntry> list = new List<HelpEntry>
            {
                // ===== General =====
                new HelpEntry
                {
                    Id = "General.VaultName",
                    Tab = "General",
                    Title = "Vault Name",
                    Summary = "Folder name of your Obsidian vault, placed under Vault Base Path.",
                    Description = "SlingMD combines Vault Base Path + Vault Name to find the vault root, which must already exist. Without Obsidian, any folder works: notes are saved there as plain Markdown files. With Obsidian, the name must match the vault's name in Obsidian so notes open there.",
                    Default = "Logic"
                },
                new HelpEntry
                {
                    Id = "General.VaultBasePath",
                    Tab = "General",
                    Title = "Vault Base Path",
                    Summary = "Parent directory that contains your Obsidian vault folder.",
                    Description = "Absolute filesystem path. Example: C:\\Users\\You\\Documents\\Notes. The vault folder (named by Vault Name) is expected to live inside this path."
                },
                new HelpEntry
                {
                    Id = "General.LaunchObsidian",
                    Tab = "General",
                    Title = "Launch Obsidian after saving",
                    Summary = "Opens the saved note in Obsidian automatically when the export finishes.",
                    Description = "Turn off if you prefer to review notes without Obsidian stealing focus, or if you don't use Obsidian and just want the Markdown files. The note is still saved either way."
                },
                new HelpEntry
                {
                    Id = "General.ShowNoteInExplorer",
                    Tab = "General",
                    Title = "Show the saved note in File Explorer",
                    Summary = "Opens File Explorer with the note selected after you sling an email or appointment.",
                    Description = "Handy without Obsidian, when the note is just a Markdown file in a folder. Works alongside Launch Obsidian. Batch slings and auto-sling don't open Explorer.",
                    Default = "Off"
                },
                new HelpEntry
                {
                    Id = "General.CopyNotePathToClipboard",
                    Tab = "General",
                    Title = "Copy note path to clipboard after saving",
                    Summary = "Puts the full file path of the saved note on the clipboard when you sling an email by hand, ready to paste elsewhere. Automatic and batch slings leave the clipboard alone. On by default."
                },
                new HelpEntry
                {
                    Id = "General.ShowCountdown",
                    Tab = "General",
                    Title = "Show countdown",
                    Summary = "Displays a small countdown dialog between save and launch.",
                    Description = "Only meaningful when Delay > 0. Useful if you want a visible pause before Obsidian opens so you can cancel."
                },
                new HelpEntry
                {
                    Id = "General.Delay",
                    Tab = "General",
                    Title = "Delay (seconds)",
                    Summary = "Pause in seconds between writing the note and launching Obsidian.",
                    Description = "Use a short delay (1–3s) if Obsidian's file watcher sometimes races ahead of the save."
                },
                new HelpEntry
                {
                    Id = "General.TemplatesFolder",
                    Tab = "General",
                    Title = "Templates Folder",
                    Summary = "Where SlingMD looks for user-provided template overrides.",
                    Description = "Vault-relative (e.g. \"Templates\") or an absolute path. SlingMD falls back to built-in templates when a file is missing.",
                    Default = "Templates"
                },
                new HelpEntry
                {
                    Id = "General.IncludeDailyNoteLink",
                    Tab = "General",
                    Title = "Include Daily Note Link",
                    Summary = "Adds a dailyNoteLink frontmatter field linking each exported note to its day.",
                    Description = "When enabled, the daily note link uses the format below. Useful for Daily Note plugins that surface related notes."
                },
                new HelpEntry
                {
                    Id = "General.DailyNoteLinkFormat",
                    Tab = "General",
                    Title = "Daily Note Link Format",
                    Summary = "Format string for the dailyNoteLink frontmatter field.",
                    Description = "Must resolve to a wikilink. Any .NET DateTime format placeholder works inside the [[...]] brackets.",
                    Default = "[[yyyy-MM-dd]]",
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "[[yyyy-MM-dd]]", Output = "[[2026-04-22]]" },
                        new HelpExample { Input = "[[Daily/yyyy-MM-dd]]", Output = "[[Daily/2026-04-22]]" },
                    }
                },

                // ===== Email =====
                new HelpEntry
                {
                    Id = "Email.InboxFolder",
                    Tab = "Email",
                    Title = "Inbox Folder",
                    Summary = "Vault subfolder where exported emails are written.",
                    Default = "Inbox"
                },
                new HelpEntry
                {
                    Id = "Email.PromptForFolderOnSling",
                    Tab = "Email",
                    Title = "Ask Where to Sling",
                    Summary = "Prompt for a destination subfolder each time you sling a single email.",
                    Description = "When on, slinging one email opens a folder picker listing subfolders under your Inbox folder, where you can also create a new one. Skip writes to the Inbox folder as usual; Cancel aborts the sling. Slinging multiple emails always prompts, regardless of this setting.",
                    Default = "Off"
                },
                new HelpEntry
                {
                    Id = "Email.NoteTitleFormat",
                    Tab = "Email",
                    Title = "Note Title Format",
                    Summary = "Controls the title at the top of each exported email note.",
                    Description = "Note title is separate from the filename on disk. A trailing separator is auto-trimmed when a token renders empty.",
                    Default = "{Subject} - {Date}",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{Subject}", "Cleaned subject line" },
                        { "{Sender}", "Sender short name" },
                        { "{Date}", "Email received date" },
                    },
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "{Subject} - {Date}", Output = "Status update - 2026-04-22" },
                        new HelpExample { Input = "{Sender}: {Subject}", Output = "JaneS: Status update" },
                    }
                },
                new HelpEntry
                {
                    Id = "Email.NoteTitleMaxLength",
                    Tab = "Email",
                    Title = "Max Title Length",
                    Summary = "Truncates the rendered title at this many characters (with ellipsis).",
                    Description = "File and folder names can end up shorter than this: SlingMD shortens them, at a word break, whenever the full path would pass the Windows path limit (240 characters). This mostly affects deep vault paths and emails inside thread folders.",
                    Default = "50"
                },
                new HelpEntry
                {
                    Id = "Email.NoteTitleIncludeDate",
                    Tab = "Email",
                    Title = "Include Date in Title",
                    Summary = "Controls whether the {Date} token renders.",
                    Description = "When off, {Date} resolves to empty and the trailing separator is trimmed automatically."
                },
                new HelpEntry
                {
                    Id = "Email.DefaultNoteTags",
                    Tab = "Email",
                    Title = "Default Note Tags",
                    Summary = "Comma-separated tags added to every exported email's frontmatter.",
                    Description = "Applied as YAML list under tags:. Leave blank to emit no tags. Tags are cleaned up for Obsidian: a leading # is removed and spaces become -.",
                    Default = "email (installs before 1.2.5: FollowUp)"
                },
                new HelpEntry
                {
                    Id = "Email.FoldQuotedReplies",
                    Tab = "Email",
                    Title = "Fold quoted replies",
                    Summary = "Shows only what a reply added; the quoted history underneath goes into a collapsed \"Earlier messages\" section at the end of the note.",
                    Description = "Recognises Outlook's From/Sent header, Gmail and Apple Mail quotes, -----Original Message----- and \"On … wrote:\". Nothing is removed: click the section in Obsidian to expand it. Forwards are always kept in full, and an email whose history can't be found is unchanged.",
                    Default = "Off"
                },
                new HelpEntry
                {
                    Id = "Email.SubjectCleanupPatterns",
                    Tab = "Email",
                    Title = "Subject Cleanup Patterns",
                    Summary = "Ordered regex patterns removed from the subject before it's used anywhere.",
                    Description = "Case-insensitive. Each pattern is stripped (replaced with empty). Built-in defaults strip Re:/Fwd: prefixes and common tags like [EXTERNAL]. Invalid regex rows are silently skipped."
                },
                new HelpEntry
                {
                    Id = "Email.EmailFilenameFormat",
                    Tab = "Email",
                    Title = "Email Filename Format",
                    Summary = "Template for the on-disk filename. Blank = legacy behavior.",
                    Description = "Filename-invalid characters are sanitized out. When grouping threads, the date may be moved to the front (see Threading tab).",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{Subject}", "Cleaned subject" },
                        { "{Sender}", "Sender short name" },
                        { "{Date}", "yyyy-MM-dd" },
                        { "{Timestamp}", "yyyy-MM-dd_HHmm" },
                    },
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "", Output = "(legacy default: Subject-Sender-timestamp)" },
                        new HelpExample { Input = "{Date}_{Subject}", Output = "2026-04-22_Status-update.md" },
                    }
                },
                new HelpEntry
                {
                    Id = "Email.EmailTemplateFile",
                    Tab = "Email",
                    Title = "Email Template File",
                    Summary = "Markdown template for exported email notes.",
                    Description = "Looked up under Templates Folder. If missing, SlingMD uses its built-in default.",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{{frontmatter}}", "The YAML properties block" },
                        { "{{body}}", "Email body, converted from HTML to Markdown" },
                        { "{{taskBlock}}", "The follow-up task line, when a task is created" },
                        { "{{subject}}", "Cleaned subject" },
                        { "{{noteTitle}}", "Rendered note title" },
                        { "{{senderName}}", "Sender display name" },
                        { "{{senderShortName}}", "Sender short name" },
                        { "{{senderEmail}}", "Sender email address" },
                        { "{{from}} / {{to}} / {{cc}}", "Contact links (see Contact Link Format)" },
                        { "{{fromEmail}} / {{toEmail}} / {{ccEmail}}", "Email addresses" },
                        { "{{date}} / {{timestamp}}", "Received date (Email Date Format)" },
                        { "{{threadNote}} / {{threadId}}", "Thread summary link and id" },
                        { "{{dailyNoteLink}}", "Daily note link, when enabled" },
                        { "{{tags}}", "Default note tags" },
                        { "{{fileName}} / {{fileNameNoExt}}", "The note's file name" },
                        { "{{<list>Csv}} / {{<list>Yaml}}", "Any list (to, cc, tags...) as comma-separated text or a YAML list" },
                    },
                    Default = "EmailTemplate.md"
                },
                new HelpEntry
                {
                    Id = "Email.EmailDateFormat",
                    Tab = "Email",
                    Title = "Email Date Format",
                    Summary = ".NET DateTime format string used for the date field in email frontmatter.",
                    Description = "Keep the T between date and time (ISO 8601) so Obsidian Properties and Dataview treat the value as a date and sort it correctly. Installs from before 1.2.5 keep their old default, yyyy-MM-dd HH:mm:ss, which Obsidian shows as text; change it here to switch. Existing notes are not rewritten.",
                    Default = "yyyy-MM-ddTHH:mm:ss",
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "yyyy-MM-ddTHH:mm:ss", Output = "2026-04-22T14:05:30" },
                        new HelpExample { Input = "yyyy-MM-dd HH:mm:ss", Output = "2026-04-22 14:05:30 (shown as text in Obsidian)" },
                        new HelpExample { Input = "MM/dd/yyyy h:mm tt", Output = "04/22/2026 2:05 PM" },
                    }
                },

                // ===== Appointments =====
                new HelpEntry
                {
                    Id = "Appointments.AppointmentsFolder",
                    Tab = "Appointments",
                    Title = "Appointments Folder",
                    Summary = "Vault subfolder for exported calendar appointments.",
                    Default = "Appointments"
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentNoteTitleFormat",
                    Tab = "Appointments",
                    Title = "Appointment Note Title Format",
                    Summary = "Title format for appointment notes.",
                    Default = "{Date} - {Subject}",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{Date}", "Start date (yyyy-MM-dd)" },
                        { "{Subject}", "Meeting subject" },
                        { "{Sender}", "Organizer short name" },
                    }
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentNoteTitleMaxLength",
                    Tab = "Appointments",
                    Title = "Max Title Length (Appointments)",
                    Summary = "Truncates rendered appointment titles at this many characters.",
                    Default = "50"
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentDefaultTags",
                    Tab = "Appointments",
                    Title = "Default Tags (Appointments)",
                    Summary = "Comma-separated tags applied to every appointment note.",
                    Default = "Appointment"
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentTemplateFile",
                    Tab = "Appointments",
                    Title = "Appointment Template File",
                    Summary = "Markdown template for appointment notes.",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{{frontmatter}}", "The YAML properties block" },
                        { "{{body}}", "Appointment body" },
                        { "{{taskBlock}}", "The task line, when a task is created" },
                        { "{{subject}} / {{noteTitle}}", "Subject and rendered title" },
                        { "{{organizer}} / {{organizerEmail}}", "Organizer link and address" },
                        { "{{attendees}}", "Required attendees" },
                        { "{{optionalAttendees}}", "Optional attendees" },
                        { "{{resources}}", "Rooms and resources" },
                        { "{{location}}", "Location" },
                        { "{{startDateTime}} / {{endDateTime}}", "Start and end (Appointment Date Format)" },
                        { "{{date}}", "Appointment date" },
                        { "{{recurrence}}", "Recurrence state" },
                        { "{{meetingNotes}}", "Link to the companion meeting note, when created" },
                        { "{{fileName}} / {{fileNameNoExt}}", "The note's file name" },
                    },
                    Default = "AppointmentTemplate.md"
                },
                new HelpEntry
                {
                    Id = "Appointments.MeetingNoteTemplate",
                    Tab = "Appointments",
                    Title = "Meeting Note Template",
                    Summary = "Optional path to a custom meeting-notes companion template.",
                    Description = "Leave empty to use the built-in template. Applies only when Create companion meeting notes is on.",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{{frontmatter}}", "The YAML properties block" },
                        { "{{appointmentTitle}}", "The appointment's title" },
                        { "{{appointmentLink}}", "Link back to the appointment note" },
                        { "{{organizer}}", "Organizer link" },
                        { "{{attendees}}", "Attendees" },
                        { "{{date}}", "Meeting date" },
                        { "{{location}}", "Location" },
                    }
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentDateFormat",
                    Tab = "Appointments",
                    Title = "Appointment Date Format",
                    Summary = ".NET DateTime format used for appointment start/end fields.",
                    Description = "Keep the T between date and time (ISO 8601) so Obsidian Properties treat the value as a date. Installs from before 1.2.5 keep their old default, yyyy-MM-dd HH:mm.",
                    Default = "yyyy-MM-ddTHH:mm"
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentTaskCreation",
                    Tab = "Appointments",
                    Title = "Task Creation Mode (Appointments)",
                    Summary = "Whether to create a follow-up task when saving an appointment.",
                    Description = "None: no task. Obsidian: inline task in the note. Outlook: an Outlook task. Both: both.",
                    Default = "None"
                },
                new HelpEntry
                {
                    Id = "Appointments.AppointmentSaveAttachments",
                    Tab = "Appointments",
                    Title = "Save attachments (Appointments)",
                    Summary = "Persists files attached to the calendar item (excluding .ics) alongside the note."
                },
                new HelpEntry
                {
                    Id = "Appointments.CreateMeetingNotes",
                    Tab = "Appointments",
                    Title = "Create companion meeting notes",
                    Summary = "Creates a linked meeting-notes stub when exporting an appointment."
                },
                new HelpEntry
                {
                    Id = "Appointments.GroupRecurringMeetings",
                    Tab = "Appointments",
                    Title = "Group recurring meeting instances",
                    Summary = "Puts instances of a recurring series into a shared folder with a summary note."
                },
                new HelpEntry
                {
                    Id = "Appointments.SaveCancelledAppointments",
                    Tab = "Appointments",
                    Title = "Save cancelled appointments",
                    Summary = "When off, cancelled meetings are skipped silently."
                },

                // ===== Contacts =====
                new HelpEntry
                {
                    Id = "Contacts.ContactsFolder",
                    Tab = "Contacts",
                    Title = "Contacts Folder",
                    Summary = "Vault subfolder where contact notes are created.",
                    Default = "Contacts"
                },
                new HelpEntry
                {
                    Id = "Contacts.EnableContactSaving",
                    Tab = "Contacts",
                    Title = "Enable Contact Saving",
                    Summary = "Master switch: when off, SlingMD does not create or update contact notes."
                },
                new HelpEntry
                {
                    Id = "Contacts.SearchEntireVaultForContacts",
                    Tab = "Contacts",
                    Title = "Search entire vault for contacts",
                    Summary = "Also look outside the Contacts folder when checking if a contact already exists.",
                    Description = "Helpful if your contacts live in a non-default location (e.g. People/Clients/). Slower on very large vaults."
                },
                new HelpEntry
                {
                    Id = "Contacts.ContactFilenameFormat",
                    Tab = "Contacts",
                    Title = "Contact Filename Format",
                    Summary = "Filename used when SlingMD creates a contact note.",
                    Default = "{ContactName}",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{ContactName}", "Full display name (e.g. \"John Smith\")" },
                        { "{ContactShortName}", "Filename-safe abbreviation (e.g. \"JohnS\")" },
                    },
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "{ContactName}", Output = "John Smith.md" },
                        new HelpExample { Input = "Contact - {ContactName}", Output = "Contact - John Smith.md" },
                    }
                },
                new HelpEntry
                {
                    Id = "Contacts.ContactTemplateFile",
                    Tab = "Contacts",
                    Title = "Contact Template File",
                    Summary = "Markdown template for contact notes.",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{{frontmatter}}", "The YAML properties block" },
                        { "{{contactName}} / {{contactShortName}}", "Full and short name" },
                        { "{{firstName}} / {{middleName}} / {{lastName}} / {{suffix}}", "Name parts" },
                        { "{{fullName}} / {{displayName}}", "Parsed full name and Outlook display name" },
                        { "{{phone}} / {{email}} / {{company}} / {{jobTitle}}", "Details (empty for contacts created from email)" },
                        { "{{address}} / {{birthday}} / {{notes}}", "More details (empty for contacts created from email)" },
                        { "{{created}}", "Creation date (Contact Date Format)" },
                        { "{{fileName}} / {{fileNameNoExt}}", "The note's file name" },
                    },
                    Default = "ContactTemplate.md"
                },
                new HelpEntry
                {
                    Id = "Contacts.ContactNoteIncludeDetails",
                    Tab = "Contacts",
                    Title = "Include contact details",
                    Summary = "Whether contact notes include the Outlook-extracted details section (phone, email, company, etc.)."
                },
                new HelpEntry
                {
                    Id = "Contacts.ContactLinkFormat",
                    Tab = "Contacts",
                    Title = "Contact Link Formats",
                    Summary = "Ordered list of formats tried when rendering contact mentions (from/to/cc/organizer/attendees). One format per line; first format whose rendered note exists in the vault wins. If none match, the first format is used and a new contact note is created.",
                    Description = "Use this when contact notes follow more than one filing convention — e.g. most people are filed as `[[{FullName}]]` but a few are filed by email address. Brackets are not required — anything around the tokens is emitted literally. List is tried top-down; put your most common format first.",
                    Default = "[[{FullName}]]\n[[{Email}]]",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{FullName}", "Full name (e.g. \"John A. Smith\")" },
                        { "{FirstName}", "Parsed first name" },
                        { "{LastName}", "Parsed last name" },
                        { "{MiddleName}", "Parsed middle name (empty if none)" },
                        { "{Suffix}", "Jr., Sr., III, PhD (empty if none)" },
                        { "{DisplayName}", "Outlook's original display string" },
                        { "{ShortName}", "Semantic short name (first name or full)" },
                        { "{Email}", "SMTP address" },
                        { "{FirstInitial}", "First letter of FirstName" },
                        { "{LastInitial}", "First letter of LastName" },
                    },
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "[[{FullName}]]", Output = "[[John Smith]]  (default wikilink)" },
                        new HelpExample { Input = "[[{LastName}, {FirstName}]]", Output = "[[Smith, John]]" },
                        new HelpExample { Input = "[[{FirstInitial}{LastName}]]", Output = "[[JSmith]]" },
                        new HelpExample { Input = "[[{FullName}|{Email}]]", Output = "[[John Smith|john@acme.com]]  (aliased)" },
                        new HelpExample { Input = "@{FullName}", Output = "@John Smith  (no wikilink)" },
                        new HelpExample { Input = "[{FullName}](mailto:{Email})", Output = "markdown mailto link" },
                    }
                },
                new HelpEntry
                {
                    Id = "Contacts.ContactDateFormat",
                    Tab = "Contacts",
                    Title = "Contact Date Format",
                    Summary = ".NET DateTime format used for the created field on contact notes.",
                    Default = "yyyy-MM-dd",
                    Examples = new List<HelpExample>
                    {
                        new HelpExample { Input = "yyyy-MM-dd", Output = "2026-04-22" },
                        new HelpExample { Input = "MMMM d, yyyy", Output = "April 22, 2026" },
                    }
                },
                new HelpEntry
                {
                    Id = "Contacts.EnableContactFuzzyMatching",
                    Tab = "Contacts",
                    Title = "Enable contact fuzzy matching",
                    Summary = "When off, contact matching uses exact filenames only — byte-identical to pre-1.3 behavior. Toggle on to match aliases and structural variants (honorifics, suffixes, parenthetical suffixes). Capped at 5000 files in vault-wide search.",
                },
                new HelpEntry
                {
                    Id = "Contacts.AutoSaveAliasOnMatchConfirmed",
                    Tab = "Contacts",
                    Title = "Auto-save alias when match confirmed",
                    Summary = "When the disambiguation dialog confirms a match, append the new name format as an alias on the matched note so future Slings of the same sender resolve as exact matches without prompting.",
                },
                new HelpEntry
                {
                    Id = "Contacts.BulkAmbiguousMatchLogPath",
                    Tab = "Contacts",
                    Title = "Bulk ambiguous-match log path",
                    Summary = "Vault-relative path where bulk operations record ambiguous matches that needed human review. Rendered as a wikilink in bulk run summaries.",
                    Default = "Logs/bulk-ambiguous-matches.md",
                },

                // ===== Tasks =====
                new HelpEntry
                {
                    Id = "Tasks.CreateObsidianTask",
                    Tab = "Tasks",
                    Title = "Create task in Obsidian note",
                    Summary = "Inserts an inline - [ ] task line at the top of every exported email."
                },
                new HelpEntry
                {
                    Id = "Tasks.CreateOutlookTask",
                    Tab = "Tasks",
                    Title = "Create task in Outlook",
                    Summary = "Also creates an Outlook task with due date and reminder.",
                    Description = "Useful if you want the task to surface in Outlook's To-Do bar, Teams, or To Do app."
                },
                new HelpEntry
                {
                    Id = "Tasks.AskForDates",
                    Tab = "Tasks",
                    Title = "Ask for dates each time",
                    Summary = "Shows a dialog on every sling to confirm due date and reminder.",
                    Description = "Disable if you want to use the defaults below without the extra click."
                },
                new HelpEntry
                {
                    Id = "Tasks.DueInDays",
                    Tab = "Tasks",
                    Title = "Due in Days",
                    Summary = "Default days from today until the task is due.",
                    Default = "1"
                },
                new HelpEntry
                {
                    Id = "Tasks.ReminderDays",
                    Tab = "Tasks",
                    Title = "Reminder Days",
                    Summary = "Default reminder offset. Can be days-from-now or days-before-due.",
                    Description = "Counted from today, or back from the due date when Reminder days are relative to the due date is on. 0 = no reminder."
                },
                new HelpEntry
                {
                    Id = "Tasks.UseRelativeReminder",
                    Tab = "Tasks",
                    Title = "Reminder days are relative to the due date",
                    Summary = "Counts Reminder Days back from the due date instead of forward from today.",
                    Description = "On: Due in 3 days with Reminder Days 1 reminds you 2 days from now (1 day before it is due). Off: Reminder Days 1 reminds you tomorrow, whatever the due date.",
                    Default = "Off"
                },
                new HelpEntry
                {
                    Id = "Tasks.ReminderHour",
                    Tab = "Tasks",
                    Title = "Reminder Hour",
                    Summary = "Hour of day (0–23) used for reminders.",
                    Default = "9"
                },
                new HelpEntry
                {
                    Id = "Tasks.DefaultTaskTags",
                    Tab = "Tasks",
                    Title = "Default Task Tags",
                    Summary = "Comma-separated tags rendered as #tags on the inline task line.",
                    Default = "FollowUp"
                },
                new HelpEntry
                {
                    Id = "Tasks.TaskTemplateFile",
                    Tab = "Tasks",
                    Title = "Task Template File",
                    Summary = "Template for the inline - [ ] task line (not the full note).",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{{noteLink}}", "Link to the note the task belongs to" },
                        { "{{noteName}}", "That note's name" },
                        { "{{tags}}", "Default task tags as #tags" },
                        { "{{createdDate}}", "Date the task was created" },
                        { "{{dueDate}}", "Due date" },
                        { "{{reminderDate}}", "Reminder date" },
                    },
                    Default = "TaskTemplate.md"
                },

                // ===== Threading =====
                new HelpEntry
                {
                    Id = "Threading.GroupEmailThreads",
                    Tab = "Threading",
                    Title = "Group email threads",
                    Summary = "Puts replies in the same thread into a shared folder with a summary note.",
                    Description = "Uses Outlook's conversation index when available, otherwise subject-based heuristics. Disable if you prefer a flat inbox layout."
                },
                new HelpEntry
                {
                    Id = "Threading.AlwaysCreateThreadFolder",
                    Tab = "Threading",
                    Title = "Always create a thread folder",
                    Summary = "Files every email into a thread folder with a summary note on its first sling, so a note never moves when a reply arrives later. Off by default.",
                    Description = "Without this, a single email stays loose in the inbox and is moved into a thread folder when a second email of the conversation is slung. Either way the note keeps the same file name, so links to it keep working."
                },
                new HelpEntry
                {
                    Id = "Threading.MoveDateToFrontInThread",
                    Tab = "Threading",
                    Title = "Move date to front in thread filename",
                    Summary = "Prefixes thread filenames with the date so they sort chronologically.",
                    Description = "Requires Include Date in Title. When off, filenames follow the Note Title Format order."
                },
                new HelpEntry
                {
                    Id = "Threading.ThreadTemplateFile",
                    Tab = "Threading",
                    Title = "Thread Template File",
                    Summary = "Template used for the thread summary note (the 0-Threadname.md file).",
                    Description = "The summary note is written once, when the thread folder is created, and never rewritten, so anything you add to it is kept. Template changes apply to new threads only.",
                    Tokens = new Dictionary<string, string>
                    {
                        { "{{title}}", "Thread title (conversation topic)" },
                        { "{{threadId}}", "Thread id shared by every note in the thread" },
                        { "{{folderPath}}", "Vault path of the thread folder, for Dataview queries" },
                    },
                    Default = "ThreadNoteTemplate.md"
                },

                // ===== Attachments =====
                new HelpEntry
                {
                    Id = "Attachments.StorageMode",
                    Tab = "Attachments",
                    Title = "Storage Location",
                    Summary = "Where attachments are saved relative to the note.",
                    Description = "Same folder: next to the .md file. Subfolder per note: in a sibling folder named after the note. Centralized: a single vault-level folder (see below).",
                    Default = "Same folder as note"
                },
                new HelpEntry
                {
                    Id = "Attachments.AttachmentsFolder",
                    Tab = "Attachments",
                    Title = "Centralized Folder Name",
                    Summary = "Vault subfolder used when Storage Location is Centralized.",
                    Description = "Files are organized into yyyy-MM subfolders inside this folder.",
                    Default = "Attachments"
                },
                new HelpEntry
                {
                    Id = "Attachments.SaveRealAttachments",
                    Tab = "Attachments",
                    Title = "Save real attachments",
                    Summary = "Persists files a person actually attached to the email (documents, images, etc.) while excluding signature graphics and other inline/embedded body images. Uses Outlook's MAPI signals to tell them apart."
                },
                new HelpEntry
                {
                    Id = "Attachments.SaveInlineImages",
                    Tab = "Attachments",
                    Title = "Save inline images",
                    Summary = "Persists images embedded in the HTML body (including signature graphics) and links them from the note. On by default.",
                    Description = "On by default for new installs only. If you upgraded from 1.2.4 or earlier, your saved setting is kept, which was off; turn it on to see email images in your notes."
                },
                new HelpEntry
                {
                    Id = "Attachments.SaveAllAttachments",
                    Tab = "Attachments",
                    Title = "Save all attachments",
                    Summary = "Persists every attachment (documents, zips, etc.), not just inline images."
                },
                new HelpEntry
                {
                    Id = "Attachments.UseObsidianWikilinks",
                    Tab = "Attachments",
                    Title = "Use Obsidian wikilinks",
                    Summary = "Links attachments as ![[file.png]] rather than standard markdown ![alt](file.png)."
                },
                new HelpEntry
                {
                    Id = "Attachments.ReuseDuplicateAttachments",
                    Tab = "Attachments",
                    Title = "Reuse duplicate attachments",
                    Summary = "Stores a file that is byte-for-byte identical to one already saved (signature logos, repeated forwards) once in <Attachments Folder>/Shared and links every later note to that copy. Off by default."
                },
                new HelpEntry
                {
                    Id = "Attachments.GroupNoteWithAttachments",
                    Tab = "Attachments",
                    Title = "Put note in its attachment folder",
                    Summary = "With 'Subfolder per note', moves an email note that has saved attachments into the folder named after it, so the note and its files sit together. Emails without attachments stay as single files, and notes inside a thread folder are not grouped. Off by default."
                },

                // ===== Auto-Sling =====
                new HelpEntry
                {
                    Id = "AutoSling.EnableAutoSling",
                    Tab = "Auto-Sling",
                    Title = "Enable Auto-Sling",
                    Summary = "Master switch for rule-based automatic slinging.",
                    Description = "When on, SlingMD evaluates every new email against the rules below and exports matching messages without you clicking anything."
                },
                new HelpEntry
                {
                    Id = "AutoSling.NotificationMode",
                    Tab = "Auto-Sling",
                    Title = "Notification Mode",
                    Summary = "How auto-sling activity is surfaced.",
                    Description = "Toast: small popup near the tray. Silent: no UI, log-only.",
                    Default = "Toast"
                },
                new HelpEntry
                {
                    Id = "AutoSling.EnableFlagToSling",
                    Tab = "Auto-Sling",
                    Title = "Enable Flag-to-Sling",
                    Summary = "Automatically slings any email you flag in Outlook.",
                    Description = "Works independently of the rules table. Combine with a category like \"Sent to Obsidian\" to track what's been slung."
                },
                new HelpEntry
                {
                    Id = "AutoSling.SentToObsidianCategory",
                    Tab = "Auto-Sling",
                    Title = "\"Sent to Obsidian\" Category",
                    Summary = "Outlook category applied to emails after SlingMD processes them.",
                    Description = "Useful as a visible marker in the Outlook inbox. Leave blank to skip categorization.",
                    Default = "Sent to Obsidian"
                },
                new HelpEntry
                {
                    Id = "AutoSling.Rules",
                    Tab = "Auto-Sling",
                    Title = "Auto-Sling Rules",
                    Summary = "Per-row: match on Sender address, Sender Domain, or Outlook Category.",
                    Description = "First matching rule wins. Enabled=false disables a rule without deleting it. Patterns are matched case-insensitively."
                },
                new HelpEntry
                {
                    Id = "AutoSling.WatchedFolders",
                    Tab = "Auto-Sling",
                    Title = "Watched Folders",
                    Summary = "Outlook folders SlingMD monitors for new items. Optional per-folder template override.",
                    Description = "Folder paths follow Outlook's hierarchy (e.g. \\\\account@example.com\\Inbox\\Work)."
                },

                // ===== Developer =====
                new HelpEntry
                {
                    Id = "Developer.ShowDevelopmentSettings",
                    Tab = "Developer",
                    Title = "Show development settings",
                    Summary = "Surfaces advanced developer-facing options elsewhere in the UI.",
                    Description = "Safe to leave on; not destructive. Hides some low-level fields by default to reduce clutter."
                },
                new HelpEntry
                {
                    Id = "Developer.ShowThreadDebug",
                    Tab = "Developer",
                    Title = "Show thread debug",
                    Summary = "Emits extra diagnostic frontmatter describing how thread grouping was resolved.",
                    Description = "Useful when debugging misgrouped conversations; harmless but noisy for normal use."
                },
            };

            Dictionary<string, HelpEntry> map = new Dictionary<string, HelpEntry>(StringComparer.Ordinal);
            foreach (HelpEntry entry in list)
            {
                map[entry.Id] = entry;
            }
            return map;
        }
    }
}
