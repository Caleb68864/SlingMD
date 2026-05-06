using System;
using System.Collections.Generic;
using System.IO;

namespace SlingMD.Outlook.Services.Formatting
{
    internal class ContactMatcher
    {
        private const int MaxFiles = 5000;
        private const int MaxFrontmatterLines = 30;

        // Full normalized name / alias → entries
        private readonly Dictionary<string, List<ContactIndexEntry>> _fullIndex =
            new Dictionary<string, List<ContactIndexEntry>>(StringComparer.Ordinal);

        // "firstKey lastKey" → entries (for HighConfidence / Ambiguous tier)
        private readonly Dictionary<string, List<ContactIndexEntry>> _firstLastIndex =
            new Dictionary<string, List<ContactIndexEntry>>(StringComparer.Ordinal);

        private readonly ContactNameNormalizer _normalizer = new ContactNameNormalizer();

        public ContactMatcher(string contactsFolderPath = null, string vaultPath = null)
        {
            if (!string.IsNullOrEmpty(contactsFolderPath) && Directory.Exists(contactsFolderPath))
                IndexDirectory(contactsFolderPath, SearchOption.TopDirectoryOnly);

            if (!string.IsNullOrEmpty(vaultPath) && Directory.Exists(vaultPath))
                IndexDirectory(vaultPath, SearchOption.AllDirectories);
        }

        public MatchResult Match(string displayName, string email)
        {
            if (string.IsNullOrWhiteSpace(displayName))
                return new MatchResult(MatchTier.None, Array.Empty<ContactIndexEntry>());

            string normalizedFull = _normalizer.Normalize(displayName);

            IReadOnlyList<ContactIndexEntry> fullHits = Lookup(_fullIndex, normalizedFull);
            if (fullHits.Count == 1)
                return new MatchResult(MatchTier.Exact, fullHits);
            if (fullHits.Count > 1)
                return new MatchResult(MatchTier.Ambiguous, fullHits);

            (string firstKey, string lastKey) = _normalizer.NormalizeFirstLast(displayName);
            if (string.IsNullOrEmpty(firstKey) && string.IsNullOrEmpty(lastKey))
                return new MatchResult(MatchTier.None, Array.Empty<ContactIndexEntry>());

            string compositeKey = (firstKey + " " + lastKey).Trim();
            IReadOnlyList<ContactIndexEntry> flHits = Lookup(_firstLastIndex, compositeKey);

            // De-duplicate by FilePath
            List<ContactIndexEntry> unique = Deduplicate(flHits);
            if (unique.Count == 1)
                return new MatchResult(MatchTier.HighConfidence, unique.AsReadOnly());
            if (unique.Count > 1)
                return new MatchResult(MatchTier.Ambiguous, unique.AsReadOnly());

            return new MatchResult(MatchTier.None, Array.Empty<ContactIndexEntry>());
        }

        private void IndexDirectory(string dir, SearchOption searchOption)
        {
            string[] files = Directory.GetFiles(dir, "*.md", searchOption);
            if (files.Length > MaxFiles)
                return;

            foreach (string filePath in files)
                IndexFile(filePath);
        }

        private void IndexFile(string filePath)
        {
            string displayName = Path.GetFileNameWithoutExtension(filePath);
            string frontmatter = ReadFrontmatter(filePath);
            IReadOnlyList<string> aliases = ExtractAliases(frontmatter);

            string normFull = _normalizer.Normalize(displayName);
            List<string> normAliases = new List<string>(aliases.Count);
            foreach (string alias in aliases)
                normAliases.Add(_normalizer.Normalize(alias));

            ContactIndexEntry entry = new ContactIndexEntry(
                filePath,
                displayName,
                aliases,
                normFull,
                normAliases.AsReadOnly());

            AddToFullIndex(normFull, entry);
            foreach (string normAlias in normAliases)
                AddToFullIndex(normAlias, entry);

            AddToFirstLastIndex(displayName, entry);
            foreach (string alias in aliases)
                AddToFirstLastIndex(alias, entry);
        }

        private void AddToFullIndex(string key, ContactIndexEntry entry)
        {
            if (string.IsNullOrEmpty(key))
                return;
            if (!_fullIndex.TryGetValue(key, out List<ContactIndexEntry> list))
            {
                list = new List<ContactIndexEntry>();
                _fullIndex[key] = list;
            }
            list.Add(entry);
        }

        private void AddToFirstLastIndex(string name, ContactIndexEntry entry)
        {
            (string firstKey, string lastKey) = _normalizer.NormalizeFirstLast(name);
            if (string.IsNullOrEmpty(firstKey) && string.IsNullOrEmpty(lastKey))
                return;

            string key = (firstKey + " " + lastKey).Trim();
            if (string.IsNullOrEmpty(key))
                return;

            if (!_firstLastIndex.TryGetValue(key, out List<ContactIndexEntry> list))
            {
                list = new List<ContactIndexEntry>();
                _firstLastIndex[key] = list;
            }
            list.Add(entry);
        }

        private static IReadOnlyList<ContactIndexEntry> Lookup(
            Dictionary<string, List<ContactIndexEntry>> index, string key)
        {
            if (string.IsNullOrEmpty(key))
                return Array.Empty<ContactIndexEntry>();

            if (index.TryGetValue(key, out List<ContactIndexEntry> entries))
                return entries.AsReadOnly();

            return Array.Empty<ContactIndexEntry>();
        }

        private static List<ContactIndexEntry> Deduplicate(IReadOnlyList<ContactIndexEntry> entries)
        {
            Dictionary<string, ContactIndexEntry> seen =
                new Dictionary<string, ContactIndexEntry>(StringComparer.Ordinal);

            foreach (ContactIndexEntry e in entries)
            {
                if (!seen.ContainsKey(e.FilePath))
                    seen[e.FilePath] = e;
            }

            return new List<ContactIndexEntry>(seen.Values);
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

            return new FrontmatterReader().ExtractAliases(frontmatter);
        }
    }
}
