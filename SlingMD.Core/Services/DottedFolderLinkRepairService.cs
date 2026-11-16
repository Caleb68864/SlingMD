using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>What a <see cref="DottedFolderLinkRepairService.RepairVault"/> run did.</summary>
    internal class DottedFolderLinkRepairReport
    {
        public int NotesScanned { get; set; }
        public int NotesChanged { get; set; }
        public int LinksRewritten { get; set; }
        public List<string> ChangedNotes { get; } = new List<string>();
        public List<string> FailedNotes { get; } = new List<string>();
    }

    /// <summary>
    /// One-shot vault fix-up for links SlingMD wrote before <see cref="FolderNameSanitizer"/>: it scans
    /// every Markdown note in the vault and rewrites link targets whose folder segment ends in a dot or
    /// space to the folder Windows actually created (see <see cref="DottedFolderLinkRewriter"/>). A link
    /// is changed only when its target does not exist under the exact name written and does exist under
    /// the repaired one, checked relative to the note and to the vault root as Obsidian resolves them.
    /// Uses only System.IO, so it is tested on any OS.
    /// </summary>
    internal class DottedFolderLinkRepairService
    {
        private readonly string _vaultRoot;

        // Directory listings by full path, so each folder is read once per run.
        private readonly Dictionary<string, HashSet<string>> _entries =
            new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        public DottedFolderLinkRepairService(string vaultRoot)
        {
            if (string.IsNullOrEmpty(vaultRoot))
            {
                throw new ArgumentNullException(nameof(vaultRoot));
            }

            _vaultRoot = Path.GetFullPath(vaultRoot);
        }

        /// <summary>Repairs every note in the vault, skipping dot-folders such as .obsidian and .trash.</summary>
        public DottedFolderLinkRepairReport RepairVault()
        {
            DottedFolderLinkRepairReport report = new DottedFolderLinkRepairReport();
            if (!Directory.Exists(_vaultRoot))
            {
                return report;
            }

            foreach (string note in EnumerateNotes(_vaultRoot))
            {
                report.NotesScanned++;
                try
                {
                    int rewritten = RepairNote(note);
                    if (rewritten > 0)
                    {
                        report.NotesChanged++;
                        report.LinksRewritten += rewritten;
                        report.ChangedNotes.Add(note);
                    }
                }
                catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    Logger.Instance.Warning($"DottedFolderLinkRepairService: could not repair '{note}': {ex.Message}");
                    report.FailedNotes.Add(note);
                }
            }

            Logger.Instance.Info(
                $"DottedFolderLinkRepairService: scanned {report.NotesScanned} notes, rewrote {report.LinksRewritten} links in {report.NotesChanged}.");
            return report;
        }

        /// <summary>Repairs one note in place and returns how many links were rewritten.</summary>
        public int RepairNote(string notePath)
        {
            byte[] bytes = File.ReadAllBytes(notePath);
            bool hasBom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
            string content = new UTF8Encoding(false).GetString(bytes, hasBom ? 3 : 0, bytes.Length - (hasBom ? 3 : 0));

            string noteDir = Path.GetDirectoryName(Path.GetFullPath(notePath));
            int rewritten;
            string updated = DottedFolderLinkRewriter.Rewrite(content, target => Resolves(noteDir, target), out rewritten);
            if (rewritten == 0 || updated == content)
            {
                return 0;
            }

            string tempPath = notePath + ".tmp." + Guid.NewGuid().ToString("N");
            bool replaced = false;
            try
            {
                File.WriteAllText(tempPath, updated, new UTF8Encoding(hasBom));
                File.Replace(tempPath, notePath, null);
                replaced = true;
            }
            finally
            {
                if (!replaced && File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                    {
                        Logger.Instance.Warning($"DottedFolderLinkRepairService: could not delete '{tempPath}': {ex.Message}");
                    }
                }
            }

            return rewritten;
        }

        /// <summary>
        /// Whether a '/'-separated link path names an existing file or folder, relative to the note's
        /// folder or (as Obsidian also tries) the vault root. Names are compared exactly, segment by
        /// segment, against directory listings: Windows would otherwise open "Jr.\file.pdf" as
        /// "Jr\file.pdf" and make every dotted link look fine. A path without an extension may name a
        /// note, so "name.md" is tried too.
        /// </summary>
        internal bool Resolves(string noteDir, string linkPath)
        {
            if (string.IsNullOrEmpty(linkPath))
            {
                return false;
            }

            if (linkPath.StartsWith("/", StringComparison.Ordinal))
            {
                return ExistsExactly(_vaultRoot, linkPath.TrimStart('/'));
            }

            return ExistsExactly(noteDir, linkPath) || ExistsExactly(_vaultRoot, linkPath);
        }

        private bool ExistsExactly(string baseDir, string linkPath)
        {
            string[] segments = linkPath.Split('/');
            string current = baseDir;
            for (int i = 0; i < segments.Length; i++)
            {
                string segment = segments[i];
                if (segment.Length == 0 || segment == ".")
                {
                    continue;
                }

                if (segment == "..")
                {
                    current = Path.GetDirectoryName(current);
                    if (current == null)
                    {
                        return false;
                    }

                    continue;
                }

                HashSet<string> names = ListEntries(current);
                bool last = i == segments.Length - 1;
                if (names.Contains(segment))
                {
                    current = Path.Combine(current, segment);
                }
                else if (last && string.IsNullOrEmpty(Path.GetExtension(segment)) && names.Contains(segment + ".md"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        private HashSet<string> ListEntries(string directory)
        {
            HashSet<string> names;
            if (_entries.TryGetValue(directory, out names))
            {
                return names;
            }

            names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (Directory.Exists(directory))
                {
                    foreach (string entry in Directory.EnumerateFileSystemEntries(directory))
                    {
                        names.Add(Path.GetFileName(entry));
                    }
                }
            }
            catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Logger.Instance.Warning($"DottedFolderLinkRepairService: could not list '{directory}': {ex.Message}");
            }

            _entries[directory] = names;
            return names;
        }

        private static IEnumerable<string> EnumerateNotes(string root)
        {
            Stack<string> pending = new Stack<string>();
            pending.Push(root);
            while (pending.Count > 0)
            {
                string directory = pending.Pop();
                string[] files;
                string[] subdirectories;
                try
                {
                    files = Directory.GetFiles(directory, "*.md");
                    subdirectories = Directory.GetDirectories(directory);
                }
                catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    Logger.Instance.Warning($"DottedFolderLinkRepairService: could not read '{directory}': {ex.Message}");
                    continue;
                }

                // GetFiles("*.md") also matches "*.mdx" style 8.3 quirks on Windows; keep real .md only.
                foreach (string file in files.Where(f => string.Equals(Path.GetExtension(f), ".md", StringComparison.OrdinalIgnoreCase)))
                {
                    yield return file;
                }

                foreach (string subdirectory in subdirectories)
                {
                    if (!Path.GetFileName(subdirectory).StartsWith(".", StringComparison.Ordinal))
                    {
                        pending.Push(subdirectory);
                    }
                }
            }
        }
    }
}
