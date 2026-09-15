using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Formats one or more Outlook email snapshots for pasting into terminals and AI sessions.
    /// </summary>
    public class ClipboardEmailService
    {
        // Windows does not publish one universal clipboard text limit, but very large payloads can
        // fail due to available memory or limitations in the receiving application. Warn before a
        // 10 MiB UTF-16 text payload while still allowing the user to attempt the copy.
        public const int ClipboardWarningBytes = 10 * 1024 * 1024;

        public bool ExceedsClipboardWarningSize(string content)
        {
            return Encoding.Unicode.GetByteCount(content ?? string.Empty) > ClipboardWarningBytes;
        }

        public string Format(IEnumerable<ClipboardEmail> emails, bool includeFrontmatter)
        {
            List<ClipboardEmail> ordered = (emails ?? Enumerable.Empty<ClipboardEmail>())
                .Where(email => email != null)
                .OrderBy(GetSortDate)
                .ToList();

            StringBuilder result = new StringBuilder();
            for (int index = 0; index < ordered.Count; index++)
            {
                if (index > 0)
                {
                    result.AppendLine();
                    result.AppendLine($"----- Email {index + 1} of {ordered.Count} -----");
                    result.AppendLine();
                }

                ClipboardEmail email = ordered[index];
                if (includeFrontmatter)
                {
                    AppendFrontmatter(result, email);
                }

                result.Append((email.Body ?? string.Empty).Trim());
            }

            return result.ToString();
        }

        private static DateTime GetSortDate(ClipboardEmail email)
        {
            return email.SentOn == DateTime.MinValue ? email.ReceivedTime : email.SentOn;
        }

        private static void AppendFrontmatter(StringBuilder result, ClipboardEmail email)
        {
            result.AppendLine("---");
            AppendScalar(result, "title", email.Subject);
            AppendScalar(result, "type", "email");
            AppendScalar(result, "from", email.SenderName);
            AppendScalar(result, "fromEmail", email.SenderEmail);
            AppendList(result, "to", email.To);
            AppendList(result, "cc", email.Cc);
            AppendScalar(result, "sent", FormatDate(email.SentOn));
            AppendScalar(result, "received", FormatDate(email.ReceivedTime));
            AppendScalar(result, "conversationId", email.ConversationId);
            AppendScalar(result, "entryId", email.EntryId);
            result.AppendLine("---");
            result.AppendLine();
        }

        private static string FormatDate(DateTime value)
        {
            return value == DateTime.MinValue ? string.Empty : value.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private static void AppendScalar(StringBuilder result, string key, string value)
        {
            result.AppendLine($"{key}: \"{EscapeYaml(value)}\"");
        }

        private static void AppendList(StringBuilder result, string key, IEnumerable<string> values)
        {
            List<string> list = (values ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList();
            result.AppendLine(key + ":");
            foreach (string value in list)
            {
                result.AppendLine($"  - \"{EscapeYaml(value)}\"");
            }
        }

        private static string EscapeYaml(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }
    }
}
