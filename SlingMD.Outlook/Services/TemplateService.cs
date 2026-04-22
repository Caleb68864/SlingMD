using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class EmailTemplateContext
    {
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
        public string NoteTitle { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string SenderShortName { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
        public string Timestamp { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string TaskBlock { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FileNameWithoutExtension { get; set; } = string.Empty;
        public string ThreadNote { get; set; } = string.Empty;
        public string ThreadId { get; set; } = string.Empty;
    }

    public class ContactTemplateContext
    {
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
        public string ContactName { get; set; } = string.Empty;
        public string ContactShortName { get; set; } = string.Empty;
        public string Created { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FileNameWithoutExtension { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Company { get; set; } = string.Empty;
        public string JobTitle { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
        public string Birthday { get; set; } = string.Empty;
        public string Notes { get; set; } = string.Empty;
        public bool IncludeDetails { get; set; } = true;

        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string MiddleName { get; set; } = string.Empty;
        public string Suffix { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
    }

    public class TaskTemplateContext
    {
        public string NoteLink { get; set; } = string.Empty;
        public string NoteName { get; set; } = string.Empty;
        public string Tags { get; set; } = string.Empty;
        public string CreatedDate { get; set; } = string.Empty;
        public string ReminderDate { get; set; } = string.Empty;
        public string DueDate { get; set; } = string.Empty;
    }

    public class ThreadTemplateContext
    {
        public string Title { get; set; } = string.Empty;
        public string ThreadId { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
    }

    public class AppointmentTemplateContext
    {
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
        public string NoteTitle { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Organizer { get; set; } = string.Empty;
        public string OrganizerEmail { get; set; } = string.Empty;
        public string Attendees { get; set; } = string.Empty;
        public string OptionalAttendees { get; set; } = string.Empty;
        public string Resources { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public string StartDateTime { get; set; } = string.Empty;
        public string EndDateTime { get; set; } = string.Empty;
        public string Recurrence { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string TaskBlock { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FileNameWithoutExtension { get; set; } = string.Empty;
    }

    public class MeetingNoteTemplateContext
    {
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
        public string AppointmentTitle { get; set; } = string.Empty;
        public string AppointmentLink { get; set; } = string.Empty;
        public string Organizer { get; set; } = string.Empty;
        public string Attendees { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
    }

    /// <summary>
    /// Handles loading, rendering and construction of markdown templates. The service supports both
    /// repo-shipped defaults and user-provided overrides stored in the configured templates folder.
    /// </summary>
    public class TemplateService
    {
        private readonly FileService _fileService;
        private readonly ObsidianSettings _settings;
        private readonly SlingMD.Outlook.Services.Formatting.DateFormatter _dateFormatter;

        // Cache template contents by path. Keyed on (path, lastWriteTimeUtc) so edits to a template
        // file — by the user or by an update — invalidate automatically.
        private readonly Dictionary<string, CachedTemplate> _templateCache =
            new Dictionary<string, CachedTemplate>(StringComparer.OrdinalIgnoreCase);
        private readonly object _templateCacheLock = new object();

        private struct CachedTemplate
        {
            public DateTime LastWriteTimeUtc;
            public string Content;
        }

        public TemplateService(FileService fileService)
        {
            _fileService = fileService;
            _settings = fileService.GetSettings();
            _templatePathResolver = new SlingMD.Outlook.Services.Formatting.TemplatePathResolver();
            _dateFormatter = new SlingMD.Outlook.Services.Formatting.DateFormatter();
        }

        private readonly SlingMD.Outlook.Services.Formatting.TemplatePathResolver _templatePathResolver;

        /// <summary>
        /// Attempts to locate <paramref name="templateName"/> in the configured template folder, the vault,
        /// and the application deployment folders. The first hit is returned as raw text.
        /// Results are cached per-file and invalidated on <see cref="File.GetLastWriteTimeUtc"/> change.
        /// </summary>
        public string LoadTemplate(string templateName)
        {
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return null;
            }

            if (Path.IsPathRooted(templateName) && File.Exists(templateName))
            {
                return ReadTemplateCached(templateName);
            }

            List<string> candidatePaths = BuildTemplateCandidatePaths(templateName);
            foreach (string path in candidatePaths)
            {
                if (File.Exists(path))
                {
                    return ReadTemplateCached(path);
                }
            }

            return null;
        }

        /// <summary>
        /// Reads a template file, returning cached content when the file's last-write timestamp
        /// matches the cache entry. Any I/O error falls through to a fresh read.
        /// </summary>
        private string ReadTemplateCached(string path)
        {
            DateTime currentMtime;
            try
            {
                currentMtime = File.GetLastWriteTimeUtc(path);
            }
            catch (System.Exception)
            {
                // If we can't stat the file, bypass the cache and let ReadAllText surface the real error.
                return File.ReadAllText(path);
            }

            lock (_templateCacheLock)
            {
                if (_templateCache.TryGetValue(path, out CachedTemplate cached)
                    && cached.LastWriteTimeUtc == currentMtime)
                {
                    return cached.Content;
                }
            }

            string content = File.ReadAllText(path);

            lock (_templateCacheLock)
            {
                _templateCache[path] = new CachedTemplate
                {
                    LastWriteTimeUtc = currentMtime,
                    Content = content
                };
            }

            return content;
        }

        /// <summary>
        /// Naive string replacement renderer that swaps out <c>{{key}}</c> placeholders with the values
        /// supplied in <paramref name="replacements"/>.
        /// </summary>
        public string ProcessTemplate(string templateContent, Dictionary<string, string> replacements)
        {
            if (string.IsNullOrEmpty(templateContent))
            {
                return string.Empty;
            }

            string result = templateContent;
            if (replacements == null)
            {
                return result;
            }

            foreach (KeyValuePair<string, string> replacement in replacements)
            {
                string value = replacement.Value ?? string.Empty;
                result = result.Replace($"{{{{{replacement.Key}}}}}", value);
            }

            return result;
        }

        /// <summary>
        /// Produces a YAML front-matter block from the supplied dictionary. Lists are automatically
        /// serialised as YAML arrays and <see cref="DateTime"/> values use the <c>yyyy-MM-dd HH:mm</c>
        /// format.
        /// </summary>
        public virtual string BuildFrontMatter(Dictionary<string, object> metadata)
        {
            StringBuilder frontMatter = new StringBuilder();
            frontMatter.AppendLine("---");

            foreach (KeyValuePair<string, object> item in metadata)
            {
                if (item.Value == null)
                {
                    continue;
                }

                IEnumerable<string> stringEnumerable = item.Value as IEnumerable<string>;
                if (item.Key == "tags" && stringEnumerable != null && !(item.Value is string))
                {
                    WriteYamlList(frontMatter, item.Key, stringEnumerable);
                }
                else if (item.Value is string stringValue)
                {
                    frontMatter.AppendLine($"{item.Key}: \"{EscapeYamlDoubleQuotedScalar(stringValue)}\"");
                }
                else if (item.Value is DateTime dateTimeValue)
                {
                    // Respect EmailDateFormat setting (matches the {{date}} / {{timestamp}} contract).
                    string format = _settings?.EmailDateFormat;
                    if (string.IsNullOrWhiteSpace(format))
                    {
                        format = "yyyy-MM-dd HH:mm:ss";
                    }
                    frontMatter.AppendLine($"{item.Key}: \"{EscapeYamlDoubleQuotedScalar(_dateFormatter.Format(dateTimeValue, format))}\"");
                }
                else if (stringEnumerable != null && !(item.Value is string))
                {
                    WriteYamlList(frontMatter, item.Key, stringEnumerable);
                }
                else
                {
                    frontMatter.AppendLine($"{item.Key}: {item.Value}");
                }
            }

            frontMatter.AppendLine("---");
            frontMatter.AppendLine();
            return frontMatter.ToString();
        }

        public string RenderEmailContent(EmailTemplateContext context)
        {
            string templateContent = LoadConfiguredTemplate(_settings.EmailTemplateFile, "EmailTemplate.md");
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = GetDefaultEmailTemplate();
            }

            Dictionary<string, string> replacements = BuildMetadataReplacements(context.Metadata);
            AddReplacement(replacements, "noteTitle", context.NoteTitle);
            AddReplacement(replacements, "subject", context.Subject);
            AddReplacement(replacements, "senderName", context.SenderName);
            AddReplacement(replacements, "senderShortName", context.SenderShortName);
            AddReplacement(replacements, "senderEmail", context.SenderEmail);
            AddReplacement(replacements, "date", context.Date);
            AddReplacement(replacements, "timestamp", context.Timestamp);
            AddReplacement(replacements, "body", context.Body);
            AddReplacement(replacements, "taskBlock", context.TaskBlock);
            AddReplacement(replacements, "fileName", context.FileName);
            AddReplacement(replacements, "fileNameNoExt", context.FileNameWithoutExtension);
            AddReplacement(replacements, "threadNote", context.ThreadNote);
            AddReplacement(replacements, "threadId", context.ThreadId);

            return ProcessTemplate(templateContent, replacements);
        }

        public string RenderContactContent(ContactTemplateContext context)
        {
            // Always try to load the user's configured ContactTemplate file. The previous gate
            // skipped lookup for the common case where users drop their template in the default filename.
            string templateContent = LoadTemplate(_settings.ContactTemplateFile);

            // Fall back to built-in defaults based on IncludeDetails
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = context.IncludeDetails
                    ? GetDefaultRichContactTemplate()
                    : GetDefaultContactTemplate();
            }

            Dictionary<string, string> replacements = BuildMetadataReplacements(context.Metadata);
            AddReplacement(replacements, "contactName", context.ContactName);
            AddReplacement(replacements, "contactShortName", context.ContactShortName);
            AddReplacement(replacements, "created", context.Created);
            AddReplacement(replacements, "fileName", context.FileName);
            AddReplacement(replacements, "fileNameNoExt", context.FileNameWithoutExtension);
            AddReplacement(replacements, "phone", context.Phone);
            AddReplacement(replacements, "email", context.Email);
            AddReplacement(replacements, "company", context.Company);
            AddReplacement(replacements, "jobTitle", context.JobTitle);
            AddReplacement(replacements, "address", context.Address);
            AddReplacement(replacements, "birthday", context.Birthday);
            AddReplacement(replacements, "notes", context.Notes);
            AddReplacement(replacements, "firstName", context.FirstName);
            AddReplacement(replacements, "lastName", context.LastName);
            AddReplacement(replacements, "middleName", context.MiddleName);
            AddReplacement(replacements, "suffix", context.Suffix);
            AddReplacement(replacements, "fullName", context.FullName);
            AddReplacement(replacements, "displayName", context.DisplayName);

            return ProcessTemplate(templateContent, replacements);
        }

        public string RenderTaskLine(TaskTemplateContext context)
        {
            string templateContent = LoadConfiguredTemplate(_settings.TaskTemplateFile, "TaskTemplate.md");
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = GetDefaultTaskTemplate();
            }

            Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "noteLink", context.NoteLink ?? string.Empty },
                { "noteName", context.NoteName ?? string.Empty },
                { "tags", context.Tags ?? string.Empty },
                { "createdDate", context.CreatedDate ?? string.Empty },
                { "reminderDate", context.ReminderDate ?? string.Empty },
                { "dueDate", context.DueDate ?? string.Empty }
            };

            return ProcessTemplate(templateContent, replacements);
        }

        public string RenderThreadContent(ThreadTemplateContext context)
        {
            string templateContent = LoadConfiguredTemplate(_settings.ThreadTemplateFile, "ThreadNoteTemplate.md");
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = GetDefaultThreadNoteTemplate();
            }

            Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "title", context.Title ?? string.Empty },
                { "threadId", context.ThreadId ?? string.Empty },
                { "folderPath", context.FolderPath ?? string.Empty }
            };

            return ProcessTemplate(templateContent, replacements);
        }

        public string RenderAppointmentContent(AppointmentTemplateContext context)
        {
            string templateContent = LoadConfiguredTemplate(_settings?.AppointmentTemplateFile, "AppointmentTemplate.md");
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = GetDefaultAppointmentTemplate();
            }

            Dictionary<string, string> replacements = BuildMetadataReplacements(context.Metadata);
            AddReplacement(replacements, "noteTitle", context.NoteTitle);
            AddReplacement(replacements, "subject", context.Subject);
            AddReplacement(replacements, "organizer", context.Organizer);
            AddReplacement(replacements, "organizerEmail", context.OrganizerEmail);
            AddReplacement(replacements, "attendees", context.Attendees);
            AddReplacement(replacements, "optionalAttendees", context.OptionalAttendees);
            AddReplacement(replacements, "resources", context.Resources);
            AddReplacement(replacements, "location", context.Location);
            AddReplacement(replacements, "startDateTime", context.StartDateTime);
            AddReplacement(replacements, "endDateTime", context.EndDateTime);
            AddReplacement(replacements, "recurrence", context.Recurrence);
            AddReplacement(replacements, "date", context.Date);
            AddReplacement(replacements, "body", context.Body);
            AddReplacement(replacements, "taskBlock", context.TaskBlock);
            AddReplacement(replacements, "fileName", context.FileName);
            AddReplacement(replacements, "fileNameNoExt", context.FileNameWithoutExtension);

            return ProcessTemplate(templateContent, replacements);
        }

        public string RenderMeetingNoteContent(MeetingNoteTemplateContext context)
        {
            string meetingNoteTemplateFile = _settings?.MeetingNoteTemplateFile;
            if (string.IsNullOrEmpty(meetingNoteTemplateFile))
            {
                meetingNoteTemplateFile = "MeetingNoteTemplate.md";
            }

            string templateContent = LoadConfiguredTemplate(meetingNoteTemplateFile, "MeetingNoteTemplate.md");
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = GetDefaultMeetingNoteTemplate();
            }

            Dictionary<string, string> replacements = BuildMetadataReplacements(context.Metadata);
            AddReplacement(replacements, "appointmentTitle", context.AppointmentTitle);
            AddReplacement(replacements, "appointmentLink", context.AppointmentLink);
            AddReplacement(replacements, "organizer", context.Organizer);
            AddReplacement(replacements, "attendees", context.Attendees);
            AddReplacement(replacements, "date", context.Date);
            AddReplacement(replacements, "location", context.Location);

            return ProcessTemplate(templateContent, replacements);
        }

        public string GetDefaultAppointmentTemplate()
        {
            return "{{frontmatter}}{{taskBlock}}" + Environment.NewLine +
                "## Attendees" + Environment.NewLine + Environment.NewLine +
                "**Organizer:** {{organizer}}" + Environment.NewLine +
                "**Required:** {{attendees}}" + Environment.NewLine +
                "**Optional:** {{optionalAttendees}}" + Environment.NewLine +
                "**Resources:** {{resources}}" + Environment.NewLine + Environment.NewLine +
                "## Details" + Environment.NewLine + Environment.NewLine +
                "**Location:** {{location}}" + Environment.NewLine +
                "**Start:** {{startDateTime}}" + Environment.NewLine +
                "**End:** {{endDateTime}}" + Environment.NewLine +
                "**Recurrence:** {{recurrence}}" + Environment.NewLine + Environment.NewLine +
                "## Notes" + Environment.NewLine + Environment.NewLine +
                "{{body}}";
        }

        public string GetDefaultMeetingNoteTemplate()
        {
            return "{{frontmatter}}" + Environment.NewLine +
                "## Meeting Notes" + Environment.NewLine + Environment.NewLine +
                "**Appointment:** {{appointmentLink}}" + Environment.NewLine +
                "**Organizer:** {{organizer}}" + Environment.NewLine +
                "**Attendees:** {{attendees}}" + Environment.NewLine +
                "**Date:** {{date}}" + Environment.NewLine +
                "**Location:** {{location}}" + Environment.NewLine + Environment.NewLine +
                "## Agenda" + Environment.NewLine + Environment.NewLine +
                "-" + Environment.NewLine + Environment.NewLine +
                "## Notes" + Environment.NewLine + Environment.NewLine +
                "-" + Environment.NewLine + Environment.NewLine +
                "## Action Items" + Environment.NewLine + Environment.NewLine +
                "- [ ]";
        }

        /// <summary>
        /// Renders a file-system safe filename by applying <c>{Token}</c> replacements and then sanitising the result.
        /// </summary>
        public string RenderFilename(string format, Dictionary<string, string> replacements, string fallbackName)
        {
            string rendered = string.IsNullOrWhiteSpace(format)
                ? fallbackName
                : ApplyFilenameTokens(format, replacements);

            if (string.IsNullOrWhiteSpace(rendered))
            {
                rendered = fallbackName;
            }

            string cleaned = _fileService.CleanFileName(rendered);
            if (string.IsNullOrWhiteSpace(cleaned))
            {
                cleaned = _fileService.CleanFileName(fallbackName);
            }

            if (string.IsNullOrWhiteSpace(cleaned))
            {
                cleaned = "Note";
            }

            return cleaned;
        }

        public string GetDefaultEmailTemplate()
        {
            return "{{frontmatter}}{{taskBlock}}{{body}}";
        }

        public string GetDefaultContactTemplate()
        {
            return @"{{frontmatter}}
# {{contactName}}

## Communication History

```dataviewjs
const current = dv.current();
const contactSources = [current.title, current.file?.name, current.file?.path];

function normalizeSingle(value) {
    if (value == null) return [];

    if (typeof value === ""object"") {
        const candidates = [];
        if (value.path) candidates.push(value.path);
        if (value.display) candidates.push(value.display);
        if (value.file?.path) candidates.push(value.file.path);
        return candidates.flatMap(normalizeSingle);
    }

    const text = String(value).trim();
    if (!text) return [];

    const unwrapped = text
        .replace(/^\[\[/, """")
        .replace(/\]\]$/, """")
        .split(""|"")[0]
        .trim();

    if (!unwrapped) return [];

    const withoutExtension = unwrapped.replace(/\.md$/i, """");
    const pathParts = withoutExtension.split(""/"");
    const fileName = pathParts[pathParts.length - 1];

    return [...new Set([unwrapped, withoutExtension, fileName]
        .map(item => item.trim().toLowerCase())
        .filter(Boolean))];
}

function normalizeValue(value) {
    if (value == null) return [];
    if (Array.isArray(value)) return value.flatMap(normalizeSingle);
    return normalizeSingle(value);
}

const contactKeys = new Set(normalizeValue(contactSources));

function containsContact(field) {
    return normalizeValue(field).some(value => contactKeys.has(value));
}

function isEmailPage(page) {
    const types = normalizeValue(page.type);
    return types.includes('email') || !!page.fromEmail || !!page.internetMessageId || !!page.entryId;
}

const emails = dv.pages()
    .where(page => isEmailPage(page) && (
        containsContact(page.from) ||
        containsContact(page.to) ||
        containsContact(page.cc)
    ))
    .sort(page => page.date, 'desc');

dv.table([""Date"", ""Subject"", ""Type""],
    emails.map(page => {
        const role = containsContact(page.from)
            ? ""From""
            : containsContact(page.to)
                ? ""To""
                : ""CC"";

        return [page.date, page.file.link, role];
    })
);
```

## Notes
";
        }

        public string GetDefaultRichContactTemplate()
        {
            StringBuilder template = new StringBuilder();
            template.AppendLine("{{frontmatter}}");
            template.AppendLine("# {{contactName}}");
            template.AppendLine();
            template.AppendLine("## Contact Details");
            template.AppendLine();
            template.AppendLine("**Phone:** {{phone}}");
            template.AppendLine("**Email:** {{email}}");
            template.AppendLine("**Company:** {{company}}");
            template.AppendLine("**Title:** {{jobTitle}}");
            template.AppendLine("**Address:** {{address}}");
            template.AppendLine("**Birthday:** {{birthday}}");
            template.AppendLine();
            template.AppendLine("## Communication History");
            template.AppendLine();
            template.AppendLine("```dataviewjs");
            template.AppendLine("const current = dv.current();");
            template.AppendLine("const contactSources = [current.title, current.file?.name, current.file?.path];");
            template.AppendLine();
            template.AppendLine("function normalizeSingle(value) {");
            template.AppendLine("    if (!value && value !== 0) return [];");
            template.AppendLine();
            template.AppendLine("    if (typeof value === \"object\") {");
            template.AppendLine("        const candidates = [];");
            template.AppendLine("        if (value.path) candidates.push(value.path);");
            template.AppendLine("        if (value.display) candidates.push(value.display);");
            template.AppendLine("        if (value.file?.path) candidates.push(value.file.path);");
            template.AppendLine("        return candidates.flatMap(normalizeSingle);");
            template.AppendLine("    }");
            template.AppendLine();
            template.AppendLine("    const text = String(value).trim();");
            template.AppendLine("    if (!text) return [];");
            template.AppendLine();
            template.AppendLine("    const unwrapped = text");
            template.AppendLine("        .replace(/^\\[\\[/, \"\")");
            template.AppendLine("        .replace(/\\]\\]$/, \"\")");
            template.AppendLine("        .split(\"|\")[0]");
            template.AppendLine("        .trim();");
            template.AppendLine();
            template.AppendLine("    if (!unwrapped) return [];");
            template.AppendLine();
            template.AppendLine("    const withoutExtension = unwrapped.replace(/\\.md$/i, \"\");");
            template.AppendLine("    const pathParts = withoutExtension.split(\"/\");");
            template.AppendLine("    const fileName = pathParts[pathParts.length - 1];");
            template.AppendLine();
            template.AppendLine("    return [...new Set([unwrapped, withoutExtension, fileName]");
            template.AppendLine("        .map(item => item.trim().toLowerCase())");
            template.AppendLine("        .filter(Boolean))];");
            template.AppendLine("}");
            template.AppendLine();
            template.AppendLine("function normalizeValue(value) {");
            template.AppendLine("    if (!value && value !== 0) return [];");
            template.AppendLine("    if (Array.isArray(value)) return value.flatMap(normalizeSingle);");
            template.AppendLine("    return normalizeSingle(value);");
            template.AppendLine("}");
            template.AppendLine();
            template.AppendLine("const contactKeys = new Set(normalizeValue(contactSources));");
            template.AppendLine();
            template.AppendLine("function containsContact(field) {");
            template.AppendLine("    return normalizeValue(field).some(value => contactKeys.has(value));");
            template.AppendLine("}");
            template.AppendLine();
            template.AppendLine("function isEmailPage(page) {");
            template.AppendLine("    const types = normalizeValue(page.type);");
            template.AppendLine("    return types.includes('email') || !!page.fromEmail || !!page.internetMessageId || !!page.entryId;");
            template.AppendLine("}");
            template.AppendLine();
            template.AppendLine("const emails = dv.pages()");
            template.AppendLine("    .where(page => isEmailPage(page) && (");
            template.AppendLine("        containsContact(page.from) ||");
            template.AppendLine("        containsContact(page.to) ||");
            template.AppendLine("        containsContact(page.cc)");
            template.AppendLine("    ))");
            template.AppendLine("    .sort(page => page.date, 'desc');");
            template.AppendLine();
            template.AppendLine("dv.table([\"Date\", \"Subject\", \"Type\"],");
            template.AppendLine("    emails.map(page => {");
            template.AppendLine("        const role = containsContact(page.from)");
            template.AppendLine("            ? \"From\"");
            template.AppendLine("            : containsContact(page.to)");
            template.AppendLine("                ? \"To\"");
            template.AppendLine("                : \"CC\";");
            template.AppendLine();
            template.AppendLine("        return [page.date, page.file.link, role];");
            template.AppendLine("    })");
            template.AppendLine(");");
            template.AppendLine("```");
            template.AppendLine();
            template.AppendLine("## Notes");
            template.AppendLine();
            template.Append("{{notes}}");
            return template.ToString();
        }

        public string GetDefaultTaskTemplate()
        {
            return "- [ ] {{noteLink}} {{tags}} ➕ {{createdDate}} 🛫 {{reminderDate}} 📅 {{dueDate}}";
        }

        /// <summary>
        /// Fallback template for thread summary notes when the user has not provided their own version.
        /// Contains a DataviewJS script that summarises all emails sharing the same <c>threadId</c>.
        /// </summary>
        public string GetDefaultThreadNoteTemplate()
        {
            return @"---
title: ""{{title}}""
type: email-thread
threadId: ""{{threadId}}""
tags: [email-thread]
---

# {{title}}

```dataviewjs
// Get all emails with matching threadId from current folder
const threadId = ""{{threadId}}"";
const emails = dv.pages('""{{folderPath}}""')
    .where(p => p.threadId === threadId && p.file.name !== dv.current().file.name)
    .sort(p => p.date, 'desc');

// Display thread summary
if (emails.length > 0) {
    const startDate = emails[emails.length-1].date;
    const latestDate = emails[0].date;
    const participants = new Set();
    emails.forEach(e => {
        // Handle from field
        if (e.from) {
            const fromName = String(e.from).match(/\[\[(.*?)\]\]/)?.[1];
            if (fromName) participants.add(fromName);
        }

        // Handle to field
        if (e.to) {
            const toList = Array.isArray(e.to) ? e.to : [e.to];
            toList.forEach(to => {
                const name = String(to).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }

        // Handle cc field
        if (e.cc) {
            const ccList = Array.isArray(e.cc) ? e.cc : [e.cc];
            ccList.forEach(cc => {
                const name = String(cc).match(/\[\[(.*?)\]\]/)?.[1];
                if (name) participants.add(name);
            });
        }
    });

    dv.header(2, 'Thread Summary');
    dv.list([
        `Started: ${startDate}`,
        `Latest: ${latestDate}`,
        `Messages: ${emails.length}`,
        `Participants: ${Array.from(participants).map(p => `[[${p}]]`).join(', ')}`
    ]);
}

// Display email timeline
dv.header(2, 'Email Timeline');
for (const email of emails) {
    dv.header(3, `${email.file.name} - ${email.date}`);
    dv.paragraph(`![[${email.file.name}]]`);
}
```";
        }

        private string LoadConfiguredTemplate(string configuredTemplateFile, string defaultTemplateFile)
        {
            string configuredTemplateContent = LoadTemplate(configuredTemplateFile);
            if (!string.IsNullOrEmpty(configuredTemplateContent))
            {
                return configuredTemplateContent;
            }

            if (!string.Equals(configuredTemplateFile, defaultTemplateFile, StringComparison.OrdinalIgnoreCase))
            {
                return LoadTemplate(defaultTemplateFile);
            }

            return null;
        }

        private List<string> BuildTemplateCandidatePaths(string templateName)
        {
            // Pure path-resolution logic lives in TemplatePathResolver; we just supply the
            // environment-bound base directories.
            List<string> baseDirectories = new List<string>
            {
                AppDomain.CurrentDomain.BaseDirectory,
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                Directory.GetCurrentDirectory(),
                Environment.CurrentDirectory
            };
            return _templatePathResolver.Resolve(templateName, _settings, baseDirectories);
        }

        private static void AddReplacement(Dictionary<string, string> replacements, string key, string value)
        {
            replacements[key] = value ?? string.Empty;
        }

        private Dictionary<string, string> BuildMetadataReplacements(Dictionary<string, object> metadata)
        {
            Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "frontmatter", BuildFrontMatter(metadata ?? new Dictionary<string, object>()) }
            };

            if (metadata == null)
            {
                return replacements;
            }

            foreach (KeyValuePair<string, object> item in metadata)
            {
                if (item.Value == null)
                {
                    continue;
                }

                IEnumerable<string> stringEnumerable = item.Value as IEnumerable<string>;
                if (stringEnumerable != null && !(item.Value is string))
                {
                    List<string> values = stringEnumerable.Where(staticValue => !string.IsNullOrWhiteSpace(staticValue)).ToList();
                    replacements[item.Key] = string.Join(", ", values);
                    replacements[item.Key + "Csv"] = string.Join(", ", values);
                    replacements[item.Key + "Yaml"] = values.Count == 0
                        ? "[]"
                        : string.Join(Environment.NewLine, values.Select(value => $"- \"{value}\""));
                }
                else if (item.Value is DateTime dateTimeValue)
                {
                    // Respect EmailDateFormat setting when rendering DateTime frontmatter values
                    // (the same format used for {{timestamp}} and the email "date" metadata field).
                    string format = _settings?.EmailDateFormat;
                    if (string.IsNullOrWhiteSpace(format))
                    {
                        format = "yyyy-MM-dd HH:mm:ss";
                    }
                    replacements[item.Key] = _dateFormatter.Format(dateTimeValue, format);
                }
                else
                {
                    replacements[item.Key] = Convert.ToString(item.Value) ?? string.Empty;
                }
            }

            return replacements;
        }

        private static string ApplyFilenameTokens(string format, Dictionary<string, string> replacements)
        {
            string result = format ?? string.Empty;
            if (replacements == null)
            {
                return result;
            }

            foreach (KeyValuePair<string, string> replacement in replacements)
            {
                result = result.Replace("{" + replacement.Key + "}", replacement.Value ?? string.Empty);
            }

            return result;
        }

        /// <summary>
        /// Escapes a string value so it is safe to embed inside a double-quoted YAML scalar.
        /// Replaces backslashes, double-quotes, and newline/carriage-return characters with
        /// their YAML escape sequences so the surrounding double-quote delimiters remain valid.
        /// </summary>
        internal static string EscapeYamlDoubleQuotedScalar(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            // Order matters: backslash must be escaped first to avoid double-escaping.
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r\n", "\\n")
                .Replace("\n", "\\n")
                .Replace("\r", "\\n");
        }

        private static void WriteYamlList(StringBuilder frontMatter, string key, IEnumerable<string> values)
        {
            List<string> materializedValues = values == null
                ? new List<string>()
                : values.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();

            if (materializedValues.Count == 0)
            {
                frontMatter.AppendLine($"{key}: []");
                return;
            }

            frontMatter.AppendLine($"{key}: ");
            foreach (string value in materializedValues)
            {
                frontMatter.AppendLine($"  - \"{EscapeYamlDoubleQuotedScalar(value)}\"");
            }
        }
    }
}


