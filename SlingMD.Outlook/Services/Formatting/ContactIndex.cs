using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace SlingMD.Outlook.Services.Formatting
{
    internal class ContactIndex
    {
        private const int MaxFiles = 5000;
        private const int MaxFrontmatterLines = 30;

        private readonly Dictionary<string, List<ContactIndexEntry>> _index =
            new Dictionary<string, List<ContactIndexEntry>>(StringComparer.Ordinal);

        private readonly ContactNameNormalizer _normalizer = new ContactNameNormalizer();

        public void BuildContactsFolderTier(string contactsFolderPath)
        {
            if (string.IsNullOrEmpty(contactsFolderPath) || !Directory.Exists(contactsFolderPath))
                return;

            string[] files = Directory.GetFiles(contactsFolderPath, "*.md", SearchOption.TopDirectoryOnly);
            if (files.Length > MaxFiles)
            {
                Trace.TraceWarning(
                    $"[ContactIndex] Contacts folder tier skipped: {files.Length} files exceeds the {MaxFiles}-file cap at '{contactsFolderPath}'.");
                return;
            }

            foreach (string filePath in files)
                IndexFile(filePath);
        }

        public void BuildVaultTier(string vaultPath)
        {
            if (string.IsNullOrEmpty(vaultPath) || !Directory.Exists(vaultPath))
                return;

            string[] files = Directory.GetFiles(vaultPath, "*.md", SearchOption.AllDirectories);
            if (files.Length > MaxFiles)
            {
                Trace.TraceWarning(
                    $"[ContactIndex] Vault tier skipped: {files.Length} files exceeds the {MaxFiles}-file cap at '{vaultPath}'.");
                return;
            }

            foreach (string filePath in files)
                IndexFile(filePath);
        }

        public IReadOnlyList<ContactIndexEntry> Lookup(string normalizedKey)
        {
            if (string.IsNullOrEmpty(normalizedKey))
                return Array.Empty<ContactIndexEntry>();

            if (_index.TryGetValue(normalizedKey, out List<ContactIndexEntry> entries))
                return entries.AsReadOnly();

            return Array.Empty<ContactIndexEntry>();
        }

        private void IndexFile(string filePath)
        {
            string displayName = Path.GetFileNameWithoutExtension(filePath);
            string frontmatter = ReadFrontmatter(filePath);
            IReadOnlyList<string> aliases = ExtractAliases(frontmatter);

            string normalizedDisplayName = _normalizer.Normalize(displayName);
            List<string> normalizedAliases = new List<string>(aliases.Count);
            foreach (string alias in aliases)
                normalizedAliases.Add(_normalizer.Normalize(alias));

            ContactIndexEntry entry = new ContactIndexEntry(
                filePath,
                displayName,
                aliases,
                normalizedDisplayName,
                normalizedAliases.AsReadOnly());

            AddToIndex(normalizedDisplayName, entry);
            foreach (string normAlias in normalizedAliases)
                AddToIndex(normAlias, entry);
        }

        private void AddToIndex(string key, ContactIndexEntry entry)
        {
            if (string.IsNullOrEmpty(key))
                return;

            if (!_index.TryGetValue(key, out List<ContactIndexEntry> list))
            {
                list = new List<ContactIndexEntry>();
                _index[key] = list;
            }

            list.Add(entry);
        }

        private static string ReadFrontmatter(string filePath)
        {
            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string firstLine = reader.ReadLine();
                    if (firstLine == null || firstLine.TrimEnd() != "---")
                        return string.Empty;

                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    sb.AppendLine(firstLine);

                    int lineCount = 1;
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        sb.AppendLine(line);
                        lineCount++;

                        if (line.TrimEnd() == "---")
                            break;

                        if (lineCount >= MaxFrontmatterLines)
                            break;
                    }

                    return sb.ToString();
                }
            }
            catch (System.Exception)
            {
                return string.Empty;
            }
        }

        private static IReadOnlyList<string> ExtractAliases(string frontmatter)
        {
            if (string.IsNullOrEmpty(frontmatter))
                return Array.Empty<string>();

            FrontmatterReader reader = new FrontmatterReader();
            return reader.ExtractAliases(frontmatter);
        }
    }
}
