using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Builds an appointment note from <see cref="AppointmentData"/>: title, front matter and the
    /// rendered note (custom template or the built-in layout). No COM or UI; the add-in's
    /// <c>AppointmentProcessor</c> reads the <c>AppointmentItem</c> with
    /// <c>OutlookAppointmentReader</c>, calls this, then does the file, recurring-folder, contact,
    /// attachment and task side effects. Moved here from <c>AppointmentProcessor</c> unchanged.
    /// </summary>
    public class AppointmentNoteEngine
    {
        private readonly ObsidianSettings _settings;
        private readonly TemplateService _templateService;
        private readonly Func<string, bool> _contactNoteExists;
        private readonly DateFormatter _dateFormatter = new DateFormatter();
        private readonly ContactNameParser _contactNameParser = new ContactNameParser();
        private readonly ContactLinkFormatter _contactLinkFormatter = new ContactLinkFormatter();
        private readonly NoteTitleBuilder _noteTitleBuilder = new NoteTitleBuilder();

        /// <param name="contactNoteExists">
        /// Whether a contact note with this file stem exists, for the contact link formats; the
        /// add-in passes <c>ContactService.NoteFileExists</c>.
        /// </param>
        public AppointmentNoteEngine(ObsidianSettings settings, TemplateService templateService, Func<string, bool> contactNoteExists)
        {
            _settings = settings;
            _templateService = templateService;
            _contactNoteExists = contactNoteExists ?? (stem => false);
        }

        /// <summary>A person as a link: the first contact link format that applies, else <c>[[name]]</c>.</summary>
        public string FormatPersonLink(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }
            ContactName parsed = _contactNameParser.Parse(name, null);
            string formatted = _contactLinkFormatter.Resolve(parsed, _settings.ContactLinkFormats, _contactNoteExists);
            return string.IsNullOrEmpty(formatted) ? $"[[{name}]]" : formatted;
        }

        /// <summary>
        /// The note title: <see cref="ObsidianSettings.AppointmentNoteTitleFormat"/> with {Date},
        /// {Subject} and {Sender} (the organizer's short name), trimmed to
        /// <see cref="ObsidianSettings.AppointmentNoteTitleMaxLength"/>.
        /// </summary>
        public string BuildNoteTitle(string subjectClean, string organizerShortName, DateTime startTime)
        {
            string dateStr = startTime.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
            string titleFormat = _settings.AppointmentNoteTitleFormat ?? "{Date} - {Subject}";
            int maxLength = _settings.AppointmentNoteTitleMaxLength > 0 ? _settings.AppointmentNoteTitleMaxLength : 50;

            Dictionary<string, string> titleTokens = new Dictionary<string, string>
            {
                { "Date", dateStr },
                { "Subject", subjectClean },
                { "Sender", organizerShortName }
            };
            return _noteTitleBuilder.BuildTrimmed(titleFormat, titleTokens, maxLength);
        }

        /// <summary>
        /// Whether the instance goes into a recurring-series folder: occurrences and exceptions, when
        /// <see cref="ObsidianSettings.GroupRecurringMeetings"/> is on.
        /// </summary>
        public bool ShouldGroupAsRecurring(AppointmentRecurrence recurrence)
        {
            if (!_settings.GroupRecurringMeetings)
            {
                return false;
            }

            return recurrence == AppointmentRecurrence.Occurrence
                || recurrence == AppointmentRecurrence.Exception;
        }

        /// <summary>
        /// Readable recurrence value for the note ("None", "Series", "Occurrence", "Exception") instead of
        /// Outlook's enum name (e.g. "olApptNotRecurring").
        /// </summary>
        public static string DescribeRecurrence(AppointmentRecurrence recurrence)
        {
            switch (recurrence)
            {
                case AppointmentRecurrence.Master:
                    return "Series";
                case AppointmentRecurrence.Occurrence:
                    return "Occurrence";
                case AppointmentRecurrence.Exception:
                    return "Exception";
                default:
                    return "None";
            }
        }

        /// <summary>
        /// The front matter, in this key order: title, type, calendar, organizer, organizerEmail,
        /// location, startDateTime, endDateTime, recurrence, globalAppointmentId, then meetingNotes,
        /// attendees, optionalAttendees, resources, dailyNoteLink, tags, meetingUrl and threadId when they
        /// apply.
        /// </summary>
        /// <param name="recurringThreadId">
        /// For an instance written into a recurring-series folder, the folder name: recorded as
        /// <c>threadId</c>, which the series summary's Dataview query matches on.
        /// </param>
        public Dictionary<string, object> BuildMetadata(string noteTitle, AppointmentData appointment, string fileNameNoExt = null, string recurringThreadId = null)
        {
            DateTime startTime = appointment.Start;
            DateTime endTime = appointment.End;

            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", noteTitle },
                { "type", "Appointment" },
                { "calendar", appointment.CalendarName ?? string.Empty },
                { "organizer", FormatPersonLink(appointment.OrganizerName) },
                { "organizerEmail", appointment.OrganizerEmail ?? string.Empty },
                { "location", appointment.Location ?? string.Empty },
                { "startDateTime", FormatAppointmentDate(startTime) },
                { "endDateTime", FormatAppointmentDate(endTime) },
                { "recurrence", DescribeRecurrence(appointment.Recurrence) },
                { "globalAppointmentId", appointment.GlobalAppointmentId ?? string.Empty }
            };

            if (_settings.CreateMeetingNotes && !string.IsNullOrWhiteSpace(fileNameNoExt))
            {
                string meetingNotesTitle = fileNameNoExt + " - Meeting Notes";
                metadata.Add("meetingNotes", $"[[{meetingNotesTitle}]]");
            }

            AddLinkedList(metadata, "attendees", appointment.RequiredAttendees);
            AddLinkedList(metadata, "optionalAttendees", appointment.OptionalAttendees);
            AddLinkedList(metadata, "resources", appointment.ResourceAttendees);

            if (_settings.IncludeDailyNoteLink)
            {
                string dailyLinkFormat = _settings.DailyNoteLinkFormat ?? "[[yyyy-MM-dd]]";
                string innerFormat = dailyLinkFormat.Replace("[[", string.Empty).Replace("]]", string.Empty);
                // A malformed user format falls back to the default instead of aborting the export
                // (matches the email path).
                string dailyNoteLink = _dateFormatter.FormatOrDefault(startTime, innerFormat, _dateFormatter.Format(startTime, "yyyy-MM-dd"));
                metadata.Add("dailyNoteLink", "[[" + dailyNoteLink + "]]");
            }

            List<string> noteTags = TagNormalizer.Normalize(_settings.AppointmentDefaultNoteTags);
            if (noteTags.Count > 0)
            {
                metadata.Add("tags", noteTags);
            }

            // A Teams/Zoom/Meet/Webex join link becomes a property, so it can be clicked from the note
            // or queried, instead of staying buried in the invitation text.
            string meetingUrl = MeetingLinkFinder.Find(appointment.Body);
            if (meetingUrl != null)
            {
                metadata["meetingUrl"] = meetingUrl;
            }

            if (!string.IsNullOrWhiteSpace(recurringThreadId))
            {
                metadata["threadId"] = recurringThreadId;
            }

            return metadata;
        }

        private void AddLinkedList(Dictionary<string, object> metadata, string key, List<string> names)
        {
            if (names != null && names.Count > 0)
            {
                List<string> linked = names
                    .Select(name => FormatPersonLink(name))
                    .ToList();
                metadata.Add(key, linked);
            }
        }

        private string FormatAppointmentDate(DateTime value)
        {
            return _dateFormatter.FormatOrDefault(value, _settings.AppointmentDateFormat, _dateFormatter.Format(value, ObsidianSettings.DefaultAppointmentDateFormat));
        }

        /// <summary>
        /// Whether the user has put an appointment template in the templates folder (or set an absolute
        /// path to one). Without one, notes keep the built-in layout.
        /// </summary>
        public bool HasCustomTemplate()
        {
            string templateFile = _settings.AppointmentTemplateFile;
            if (string.IsNullOrWhiteSpace(templateFile))
            {
                return false;
            }

            try
            {
                string path = Path.IsPathRooted(templateFile)
                    ? templateFile
                    : Path.Combine(_settings.GetTemplatesPath(), templateFile);
                return File.Exists(path);
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        private static string JoinMetadataList(Dictionary<string, object> metadata, string key)
        {
            object value;
            return metadata.TryGetValue(key, out value) && value is IEnumerable<string> list
                ? string.Join(", ", list)
                : string.Empty;
        }

        /// <summary>
        /// The note text: the user's appointment template when there is one, else the built-in layout
        /// (front matter, title, task block, location, join link, start/end, organizer, calendar,
        /// description, meeting-notes heading).
        /// </summary>
        public string BuildNoteContent(
            Dictionary<string, object> metadata,
            string noteTitle,
            string subjectClean,
            AppointmentData appointment,
            string dateStr,
            string fileNameNoExt,
            string taskBlock = "")
        {
            string organizerName = appointment.OrganizerName;
            string location = appointment.Location;
            string body = appointment.Body;

            // The Appointment Template File setting used to be ignored here. Honor it when the user has
            // provided a template, so the built-in layout is unchanged for everyone else.
            if (HasCustomTemplate())
            {
                return _templateService.RenderAppointmentContent(new AppointmentTemplateContext
                {
                    Metadata = metadata,
                    NoteTitle = noteTitle,
                    Subject = subjectClean,
                    Organizer = string.IsNullOrWhiteSpace(organizerName) ? string.Empty : FormatPersonLink(organizerName),
                    OrganizerEmail = appointment.OrganizerEmail ?? string.Empty,
                    Attendees = JoinMetadataList(metadata, "attendees"),
                    OptionalAttendees = JoinMetadataList(metadata, "optionalAttendees"),
                    Resources = JoinMetadataList(metadata, "resources"),
                    Location = location ?? string.Empty,
                    StartDateTime = metadata["startDateTime"] as string,
                    EndDateTime = metadata["endDateTime"] as string,
                    Recurrence = metadata["recurrence"] as string,
                    Date = dateStr,
                    Body = (body ?? string.Empty).Trim(),
                    TaskBlock = taskBlock ?? string.Empty,
                    FileName = fileNameNoExt + ".md",
                    FileNameWithoutExtension = fileNameNoExt
                });
            }

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

            object meetingUrl;
            if (metadata.TryGetValue("meetingUrl", out meetingUrl) && meetingUrl is string joinUrl)
            {
                content.AppendLine($"**Join:** [Join the meeting]({joinUrl})");
                content.AppendLine();
            }

            content.AppendLine($"**Start:** {FormatAppointmentDate(appointment.Start)}");
            content.AppendLine($"**End:** {FormatAppointmentDate(appointment.End)}");
            content.AppendLine();

            if (!string.IsNullOrWhiteSpace(organizerName))
            {
                content.AppendLine($"**Organizer:** {FormatPersonLink(organizerName)}");
                content.AppendLine();
            }

            object calendar;
            if (metadata.TryGetValue("calendar", out calendar) && calendar is string calendarName && !string.IsNullOrWhiteSpace(calendarName))
            {
                content.AppendLine($"**Calendar:** {calendarName}");
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
        /// <summary>File name of the companion meeting note for an appointment note.</summary>
        public static string MeetingNoteFileName(string appointmentFileNameNoExt)
        {
            return appointmentFileNameNoExt + " - Meeting Notes.md";
        }

        /// <summary>
        /// The companion meeting note (<see cref="ObsidianSettings.CreateMeetingNotes"/>): front matter
        /// linking back to the appointment note, rendered through the meeting-note template.
        /// </summary>
        public string BuildMeetingNoteContent(
            string appointmentFileNameNoExt,
            string appointmentNoteTitle,
            string organizer,
            string attendees,
            string date,
            string location)
        {
            string meetingNoteTitleNoExt = appointmentFileNameNoExt + " - Meeting Notes";
            string appointmentLink = $"[[{appointmentFileNameNoExt}]]";

            Dictionary<string, object> meetingMetadata = new Dictionary<string, object>
            {
                { "title", meetingNoteTitleNoExt },
                { "type", "Meeting Notes" },
                { "appointment", appointmentLink },
                { "date", date }
            };

            List<string> noteTags = TagNormalizer.Normalize(_settings.AppointmentDefaultNoteTags);
            if (noteTags.Count > 0)
            {
                meetingMetadata.Add("tags", noteTags);
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

            return _templateService.RenderMeetingNoteContent(context);
        }

        /// <summary>
        /// The <c>0-{folder}.md</c> summary note of a recurring series folder, rendered through the
        /// thread template. Its Dataview query lists the notes whose <c>threadId</c> is
        /// <paramref name="threadFolderName"/> in the folder, given vault-relative as Dataview needs.
        /// </summary>
        public string BuildRecurringSummaryContent(string threadFolderPath, string threadFolderName, string title)
        {
            ThreadTemplateContext threadContext = new ThreadTemplateContext
            {
                Title = title,
                ThreadId = threadFolderName,
                FolderPath = ToVaultRelativeFolder(threadFolderPath)
            };

            return _templateService.RenderThreadContent(threadContext);
        }

        /// <summary>
        /// <paramref name="folderPath"/> relative to the vault root with forward slashes; a path outside
        /// the vault (or already relative) is returned with forward slashes only.
        /// </summary>
        private string ToVaultRelativeFolder(string folderPath)
        {
            string path = folderPath ?? string.Empty;
            string vaultRoot = null;
            try
            {
                vaultRoot = _settings.GetFullVaultPath().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch (ArgumentException)
            {
            }

            if (!string.IsNullOrEmpty(vaultRoot)
                && path.Length > vaultRoot.Length
                && path.StartsWith(vaultRoot, StringComparison.OrdinalIgnoreCase)
                && (path[vaultRoot.Length] == Path.DirectorySeparatorChar || path[vaultRoot.Length] == Path.AltDirectorySeparatorChar))
            {
                path = path.Substring(vaultRoot.Length + 1);
            }

            return path.Replace('\\', '/');
        }
    }
}
