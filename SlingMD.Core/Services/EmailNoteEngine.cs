using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Builds an email note from <see cref="EmailData"/>: title, base file name, front matter,
    /// Markdown body and the rendered note. No COM or UI; the add-in's <c>EmailProcessor</c> reads
    /// the <c>MailItem</c> with <c>OutlookEmailReader</c>, calls this, then does the file, thread,
    /// contact and task side effects. Moved here from <c>EmailProcessor</c> unchanged.
    /// </summary>
    public class EmailNoteEngine
    {
        private static readonly Regex ForwardSubjectRegex =
            new Regex(@"^\s*(?:fw|fwd)\s*:", RegexOptions.IgnoreCase);

        private readonly ObsidianSettings _settings;
        private readonly TemplateService _templateService;
        private readonly Func<string, string, string> _formatContactLink;
        private readonly DateFormatter _dateFormatter = new DateFormatter();
        private readonly BodyCleanerService _bodyCleaner;
        private readonly NoteTitleBuilder _noteTitleBuilder = new NoteTitleBuilder();

        /// <param name="formatContactLink">
        /// (display name, email) to the contact link written in <c>from</c>; the add-in passes
        /// <c>ContactService.FormatContactLink</c>, which looks for existing contact notes.
        /// </param>
        public EmailNoteEngine(ObsidianSettings settings, TemplateService templateService, Func<string, string, string> formatContactLink)
        {
            _settings = settings;
            _templateService = templateService;
            _formatContactLink = formatContactLink ?? ((name, email) => $"[[{name}]]");
            _bodyCleaner = new BodyCleanerService(settings ?? new ObsidianSettings());
        }

        /// <summary>
        /// The note title: <see cref="ObsidianSettings.NoteTitleFormat"/> with {Subject}, {Sender}
        /// and {Date}, trimmed to <see cref="ObsidianSettings.NoteTitleMaxLength"/>.
        /// </summary>
        public string BuildNoteTitle(string subjectClean, string senderClean, DateTime receivedTime)
        {
            string dateStr = receivedTime.ToString("yyyy-MM-dd");
            string titleFormat = _settings.NoteTitleFormat ?? "{Subject} - {Date}";
            bool includeDate = _settings.NoteTitleIncludeDate;
            int maxLength = _settings.NoteTitleMaxLength > 0 ? _settings.NoteTitleMaxLength : 50;

            Dictionary<string, string> titleTokens = new Dictionary<string, string>
            {
                { "Subject", subjectClean },
                { "Sender", senderClean },
                { "Date", includeDate ? dateStr : string.Empty }
            };
            // BuildTrimmed handles trailing dash/space left behind when a token like {Date} renders empty.
            return _noteTitleBuilder.BuildTrimmed(titleFormat, titleTokens, maxLength);
        }

        /// <summary>The front-matter date (<see cref="ObsidianSettings.EmailDateFormat"/>, safe fallback).</summary>
        public string FormatFrontMatterDate(DateTime receivedTime)
        {
            return _dateFormatter.FormatOrDefault(receivedTime, _settings.EmailDateFormat, _dateFormatter.Format(receivedTime, ObsidianSettings.DefaultEmailDateFormat));
        }

        public string BuildBaseFileName(DateTime receivedTime, string subjectClean, string senderClean, string fileDateTime, bool isThreaded)
        {
            bool includeDate = _settings.NoteTitleIncludeDate;
            string legacyBaseName;
            if (isThreaded)
            {
                if (includeDate)
                {
                    legacyBaseName = _settings.MoveDateToFrontInThread
                        ? $"{fileDateTime}-{subjectClean}-{senderClean}"
                        : $"{subjectClean}-{senderClean}-{fileDateTime}";
                }
                else
                {
                    legacyBaseName = $"{subjectClean}-{senderClean}";
                }
            }
            else
            {
                legacyBaseName = includeDate
                    ? $"{subjectClean}-{senderClean}-{fileDateTime}"
                    : $"{subjectClean}-{senderClean}";
            }

            Dictionary<string, string> replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Subject", subjectClean },
                { "Sender", senderClean },
                { "Date", receivedTime.ToString("yyyy-MM-dd") },
                { "Timestamp", fileDateTime },
                { "Threaded", isThreaded ? "true" : "false" }
            };

            return _templateService.RenderFilename(_settings.EmailFilenameFormat, replacements, legacyBaseName);
        }

        public Dictionary<string, object> BuildMetadata(
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
            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", noteTitle },
                { "type", "email" },
                { "from", _formatContactLink(senderName, senderEmail) },
                { "fromEmail", senderEmail },
                { "to", toLinked },
                { "toEmail", toEmails },
                { "threadId", conversationId },
                { "date", FormatFrontMatterDate(receivedTime) },
                { "internetMessageId", realInternetMessageId },
                { "entryId", realEntryId }
            };

            if (_settings.IncludeDailyNoteLink)
            {
                string dailyLinkFormat = _settings.DailyNoteLinkFormat ?? "[[yyyy-MM-dd]]";
                string rawDailyFormat = dailyLinkFormat.Replace("[[", string.Empty).Replace("]]", string.Empty);
                // A malformed DailyNoteLinkFormat (e.g. from hand-edited JSON) would otherwise throw
                // FormatException and fail the whole export; fall back to a safe ISO date.
                string dailyNoteLink = _dateFormatter.FormatOrDefault(receivedTime, rawDailyFormat, _dateFormatter.Format(receivedTime, "yyyy-MM-dd"));
                metadata.Add("dailyNoteLink", "[[" + dailyNoteLink + "]]");
            }

            List<string> noteTags = TagNormalizer.Normalize(_settings.DefaultNoteTags);
            if (noteTags.Count > 0)
            {
                metadata.Add("tags", noteTags);
            }

            if (ccEmails != null && ccEmails.Count > 0)
            {
                metadata.Add("cc", ccLinked);
                metadata.Add("ccEmail", ccEmails);
            }

            if (shouldGroupThread)
            {
                metadata.Add("threadNote", $"[[0-{threadNoteName}]]");
            }

            return metadata;
        }

        /// <summary>
        /// Converts the email body to Markdown, folding the quoted history of a reply into a collapsed
        /// callout when <see cref="ObsidianSettings.FoldQuotedReplies"/> is on. A forward's quoted part is
        /// the message itself, so forwards are always converted in full. The
        /// <see cref="ObsidianSettings.BodyCleanupPatterns"/> run on the reply and on the quoted history
        /// before it is folded, and <c>bodyCleaned: true</c> is recorded in <paramref name="metadata"/>
        /// when they removed anything.
        /// </summary>
        public string ConvertBody(EmailData email, Dictionary<string, object> metadata)
        {
            bool removedContent = false;
            Func<string, string> clean = text =>
            {
                string cleaned = _bodyCleaner.Clean(text, out bool removed);
                removedContent |= removed;
                return cleaned;
            };

            string body;
            if (_settings.FoldQuotedReplies && !ForwardSubjectRegex.IsMatch(email.Subject ?? string.Empty))
            {
                body = QuotedReplyFolder.ConvertFolded(email.HtmlBody, email.Body, clean);
            }
            else
            {
                body = clean(HtmlToMarkdownConverter.Convert(email.HtmlBody, email.Body));
            }

            if (removedContent && metadata != null)
            {
                metadata["bodyCleaned"] = true;
            }

            return body;
        }

        /// <summary>
        /// Applies <see cref="ObsidianSettings.BodyCleanupPatterns"/> to the Markdown body and, when a
        /// pattern removed text, records <c>bodyCleaned: true</c> in <paramref name="metadata"/>.
        /// </summary>
        public string CleanNoteBody(string body, Dictionary<string, object> metadata)
        {
            string cleaned = _bodyCleaner.Clean(body, out bool removedContent);
            if (removedContent && metadata != null)
            {
                metadata["bodyCleaned"] = true;
            }

            return cleaned;
        }

        /// <summary>
        /// Renders the whole note (front matter and body) through the email template.
        /// <paramref name="taskBlock"/> is the Obsidian task text (with its trailing blank line) or empty.
        /// </summary>
        public string RenderNote(EmailData email, Dictionary<string, object> metadata, string noteTitle, string subjectClean, string senderClean, string noteBody, string fileNameNoExt, string conversationId, string threadNoteLink, string taskBlock)
        {
            EmailTemplateContext context = new EmailTemplateContext
            {
                Metadata = metadata,
                NoteTitle = noteTitle,
                Subject = subjectClean,
                SenderName = email.SenderName ?? "Unknown Sender",
                SenderShortName = senderClean,
                SenderEmail = email.SenderEmail,
                Date = email.ReceivedTime.ToString("yyyy-MM-dd"),
                Timestamp = FormatFrontMatterDate(email.ReceivedTime),
                Body = noteBody,
                TaskBlock = taskBlock ?? string.Empty,
                FileName = fileNameNoExt + ".md",
                FileNameWithoutExtension = fileNameNoExt,
                ThreadNote = threadNoteLink,
                ThreadId = conversationId
            };

            return _templateService.RenderEmailContent(context);
        }

        /// <summary>
        /// Returns <paramref name="path"/> relative to the vault root with forward slashes, or
        /// <c>null</c> when it does not lie inside the vault.
        /// </summary>
        public string GetVaultRelativePath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }

            string vaultRoot = _settings.GetFullVaultPath().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (path.Length <= vaultRoot.Length + 1
                || !path.StartsWith(vaultRoot, StringComparison.OrdinalIgnoreCase)
                || (path[vaultRoot.Length] != Path.DirectorySeparatorChar && path[vaultRoot.Length] != Path.AltDirectorySeparatorChar))
            {
                return null;
            }

            return path.Substring(vaultRoot.Length + 1).Replace('\\', '/');
        }
    }
}
