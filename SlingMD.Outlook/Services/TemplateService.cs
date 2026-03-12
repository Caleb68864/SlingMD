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
    }

    /// <summary>
    /// Handles loading, rendering and construction of markdown templates. The service supports both
    /// repo-shipped defaults and user-provided overrides stored in the configured templates folder.
    /// </summary>
    public class TemplateService
    {
        private readonly FileService _fileService;
        private readonly ObsidianSettings _settings;

        public TemplateService(FileService fileService)
        {
            _fileService = fileService;
            _settings = fileService.GetSettings();
        }

        /// <summary>
        /// Attempts to locate <paramref name="templateName"/> in the configured template folder, the vault,
        /// and the application deployment folders. The first hit is returned as raw text.
        /// </summary>
        public string LoadTemplate(string templateName)
        {
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return null;
            }

            if (Path.IsPathRooted(templateName) && File.Exists(templateName))
            {
                return File.ReadAllText(templateName);
            }

            List<string> candidatePaths = BuildTemplateCandidatePaths(templateName);
            foreach (string path in candidatePaths)
            {
                if (File.Exists(path))
                {
                    return File.ReadAllText(path);
                }
            }

            return null;
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
                    frontMatter.AppendLine($"{item.Key}: \"{stringValue}\"");
                }
                else if (item.Value is DateTime dateTimeValue)
                {
                    frontMatter.AppendLine($"{item.Key}: {dateTimeValue:yyyy-MM-dd HH:mm}");
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
            string templateContent = LoadConfiguredTemplate(_settings.ContactTemplateFile, "ContactTemplate.md");
            if (string.IsNullOrEmpty(templateContent))
            {
                templateContent = GetDefaultContactTemplate();
            }

            Dictionary<string, string> replacements = BuildMetadataReplacements(context.Metadata);
            AddReplacement(replacements, "contactName", context.ContactName);
            AddReplacement(replacements, "contactShortName", context.ContactShortName);
            AddReplacement(replacements, "created", context.Created);
            AddReplacement(replacements, "fileName", context.FileName);
            AddReplacement(replacements, "fileNameNoExt", context.FileNameWithoutExtension);

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
                { "threadId", context.ThreadId ?? string.Empty }
            };

            return ProcessTemplate(templateContent, replacements);
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
// Find all emails where this contact appears in from, to, or cc fields
// Use title from frontmatter (original name) rather than file.name (cleaned name)
const contact = dv.current().title || dv.current().file.name;

// Helper to check if a field contains this contact
// Handles both Dataview Link objects and plain strings
function containsContact(field, contactName) {
    if (!field) return false;
    // Handle Dataview Link objects (have .path property)
    if (field.path) return field.path === contactName;
    // Handle string format - check for [[Name]] or exact match
    const str = String(field);
    return str.includes(`[[${contactName}]]`) || str === contactName;
}

// Helper to check arrays (to/cc fields can be arrays)
function checkArray(arr, contactName) {
    if (!arr) return false;
    if (!Array.isArray(arr)) return containsContact(arr, contactName);
    return arr.some(item => containsContact(item, contactName));
}

// Query all pages, then filter to only emails (pages with fromEmail field)
// and where this contact is mentioned in from, to, or cc
const emails = dv.pages()
    .where(p => {
        // Only include pages that are emails (have fromEmail field)
        if (!p.fromEmail) return false;
        // Check if this contact is mentioned in from, to, or cc
        return containsContact(p.from, contact) ||
               checkArray(p.to, contact) ||
               checkArray(p.cc, contact);
    })
    .sort(p => p.date, 'desc');

dv.table([""Date"", ""Subject"", ""Type""],
    emails.map(p => {
        // Determine message type
        const isFrom = containsContact(p.from, contact);
        const isTo = checkArray(p.to, contact);
        const type = isFrom ? ""From"" : isTo ? ""To"" : ""CC"";
        return [p.date, p.file.link, type];
    })
);
```

## Notes
";
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
const emails = dv.pages("")
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
            HashSet<string> directories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            AddDirectoryIfValid(directories, _settings.GetTemplatesPath());
            if (!Path.IsPathRooted(_settings.TemplatesFolder))
            {
                AddDirectoryIfValid(directories, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _settings.TemplatesFolder));
                AddDirectoryIfValid(directories, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), _settings.TemplatesFolder));
                AddDirectoryIfValid(directories, Path.Combine(Directory.GetCurrentDirectory(), _settings.TemplatesFolder));
                AddDirectoryIfValid(directories, Path.Combine(Environment.CurrentDirectory, _settings.TemplatesFolder));
            }

            AddDirectoryIfValid(directories, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates"));
            AddDirectoryIfValid(directories, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Templates"));
            AddDirectoryIfValid(directories, Path.Combine(Directory.GetCurrentDirectory(), "Templates"));
            AddDirectoryIfValid(directories, Path.Combine(Environment.CurrentDirectory, "Templates"));

            List<string> candidatePaths = new List<string>();
            foreach (string directory in directories)
            {
                candidatePaths.Add(Path.Combine(directory, templateName));
            }

            return candidatePaths;
        }

        private static void AddDirectoryIfValid(ISet<string> directories, string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
            {
                directories.Add(path);
            }
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
                    replacements[item.Key] = dateTimeValue.ToString("yyyy-MM-dd HH:mm:ss");
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
                frontMatter.AppendLine($"  - \"{value}\"");
            }
        }
    }
}
