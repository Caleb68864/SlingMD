using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Helpers;

namespace SlingMD.Outlook.Services.Formatting
{
    internal class ContactAliasWriter
    {
        private static readonly ContactNameNormalizer _normalizer = new ContactNameNormalizer();
        private static readonly FrontmatterReader _frontmatterReader = new FrontmatterReader();

        private static readonly Regex AliasesBlockRegex = new Regex(
            @"(aliases:\s*\r?\n)((?:[ \t]+-[ \t]+.+\r?\n?)*)",
            RegexOptions.Compiled);

        private static readonly Regex AliasesInlineRegex = new Regex(
            @"aliases:\s*\[([^\]]*)\]",
            RegexOptions.Compiled);

        // Characters requiring YAML quoting
        private static readonly Regex NeedsQuotingRegex = new Regex(
            @"[:\-\[\]""'#&*|>%@`]|^\s|\s$|[\x00-\x1f]",
            RegexOptions.Compiled);

        public bool TryAppendAlias(string filePath, string aliasToAdd)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException("filePath");
            if (string.IsNullOrEmpty(aliasToAdd))
                throw new ArgumentNullException("aliasToAdd");

            // Snapshot mtime and hash before reading
            DateTime mtimeBefore = File.GetLastWriteTimeUtc(filePath);
            string content = File.ReadAllText(filePath, Encoding.UTF8);
            string hashBefore = ComputeSha256(content);

            // Check if alias already present (case-insensitive normalized)
            string normalizedNew = _normalizer.Normalize(aliasToAdd);
            IReadOnlyList<string> existingAliases = _frontmatterReader.ExtractAliases(content);
            foreach (string existing in existingAliases)
            {
                if (_normalizer.Normalize(existing) == normalizedNew)
                    return true;
            }

            // Build the YAML-safe alias value
            string yamlValue = FormatYamlValue(aliasToAdd);

            // Rewrite content with alias appended
            string newContent = AppendAlias(content, yamlValue);

            // Check for concurrent modification before writing
            DateTime mtimeAfter = File.GetLastWriteTimeUtc(filePath);
            string contentAfterRead = File.ReadAllText(filePath, Encoding.UTF8);
            string hashAfterRead = ComputeSha256(contentAfterRead);

            if (mtimeAfter != mtimeBefore || hashAfterRead != hashBefore)
            {
                Logger.Instance.Warning(
                    $"ContactAliasWriter: concurrent modification detected, skipping write for '{filePath}'");
                return false;
            }

            // Atomic write via temp file + File.Replace
            string tempPath = filePath + ".tmp." + Guid.NewGuid().ToString("N");
            try
            {
                File.WriteAllText(tempPath, newContent, Encoding.UTF8);
                File.Replace(tempPath, filePath, null);
            }
            catch
            {
                try { File.Delete(tempPath); } catch { }
                throw;
            }

            return true;
        }

        private static string AppendAlias(string content, string yamlValue)
        {
            string newEntry = "  - " + yamlValue;

            // Try block-style match first
            Match blockMatch = AliasesBlockRegex.Match(content);
            if (blockMatch.Success)
            {
                // Append after the last block item
                string header = blockMatch.Groups[1].Value;
                string items = blockMatch.Groups[2].Value;
                string newItems = items.TrimEnd('\r', '\n') + "\n" + newEntry + "\n";
                return content.Substring(0, blockMatch.Index)
                    + header
                    + newItems
                    + content.Substring(blockMatch.Index + blockMatch.Length);
            }

            // Try inline-array match
            Match inlineMatch = AliasesInlineRegex.Match(content);
            if (inlineMatch.Success)
            {
                // Convert inline to block-style and append
                string[] parts = inlineMatch.Groups[1].Value.Split(',');
                var sb = new StringBuilder();
                sb.Append("aliases:\n");
                foreach (string part in parts)
                {
                    string val = part.Trim();
                    if (!string.IsNullOrEmpty(val))
                        sb.Append("  - ").Append(val).Append("\n");
                }
                sb.Append(newEntry).Append("\n");

                return content.Substring(0, inlineMatch.Index)
                    + sb.ToString()
                    + content.Substring(inlineMatch.Index + inlineMatch.Length);
            }

            // No aliases key: insert before closing ---
            // Find end of frontmatter
            int fmStart = content.IndexOf("---", StringComparison.Ordinal);
            if (fmStart >= 0)
            {
                int fmEnd = content.IndexOf("---", fmStart + 3, StringComparison.Ordinal);
                if (fmEnd >= 0)
                {
                    string insertion = "aliases:\n" + newEntry + "\n";
                    return content.Substring(0, fmEnd)
                        + insertion
                        + content.Substring(fmEnd);
                }
            }

            // No frontmatter at all: prepend one
            return "---\naliases:\n" + newEntry + "\n---\n" + content;
        }

        internal static string FormatYamlValue(string value)
        {
            if (NeedsQuotingRegex.IsMatch(value))
            {
                // Double-quoted with backslash-escaping for " and \
                string escaped = value
                    .Replace("\\", "\\\\")
                    .Replace("\"", "\\\"");
                return "\"" + escaped + "\"";
            }
            return value;
        }

        private static string ComputeSha256(string content)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(content));
                return BitConverter.ToString(bytes).Replace("-", string.Empty);
            }
        }
    }
}
