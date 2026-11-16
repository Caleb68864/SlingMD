using System;
using System.IO;
using SlingMD.Outlook.Helpers;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// The ids (Internet message id and EntryID) of the email notes already in the inbox, read from
    /// their front matter, so a re-sling is caught as a duplicate without scanning every note each
    /// time. Moved from <c>EmailProcessor</c> unchanged; EmailProcessor keeps one static instance, as
    /// it kept static fields before, so all processors share it.
    /// </summary>
    public class ProcessedEmailCache
    {
        // Rebuild the processed-id cache at most every N minutes. Fresh cache on every sling would
        // re-scan the whole inbox; a cache that lives forever would miss external edits.
        private const int CacheTtlMinutes = 5;

        // Cache of processed email IDs and prevent O(n*m) file scanning on every email.
        // The cache is scoped to the inbox path it was built from; when the resolved inbox path
        // changes (e.g. a batch sling redirects output to a subfolder) the cache is rebuilt so a
        // subfolder-scoped id set is never reused for the main inbox (would miss real duplicates).
        private readonly BoundedHashSet _processedEmailIds = new BoundedHashSet();
        private DateTime _cacheLastBuilt = DateTime.MinValue;
        private string _cacheBuiltForPath = null;
        private readonly object _cacheBuildLock = new object();

        /// <summary>
        /// Builds or rebuilds the email ID cache by scanning all markdown files in the inbox.
        /// Cache is rebuilt if it's older than 5 minutes or empty.
        /// This dramatically improves performance by avoiding O(n*m) file scanning on every email.
        /// </summary>
        private void EnsureBuilt(string inboxPath, DateTime now)
        {
            // Fresh enough AND built from this same inbox path -> reuse. A different path (batch
            // redirect) forces a rebuild so the id set always matches the folder being written to.
            bool samePath = string.Equals(_cacheBuiltForPath, inboxPath, StringComparison.OrdinalIgnoreCase);
            if (samePath && (now - _cacheLastBuilt).TotalMinutes < CacheTtlMinutes && _processedEmailIds.Count > 0)
            {
                return;
            }

            lock (_cacheBuildLock)
            {
                // Double-check after acquiring lock.
                samePath = string.Equals(_cacheBuiltForPath, inboxPath, StringComparison.OrdinalIgnoreCase);
                if (samePath && (now - _cacheLastBuilt).TotalMinutes < CacheTtlMinutes && _processedEmailIds.Count > 0)
                {
                    return;
                }

                _processedEmailIds.Clear();
                _cacheBuiltForPath = inboxPath;

                if (!Directory.Exists(inboxPath))
                {
                    // Inbox folder does not yet exist (first run or vault not yet created).
                    // Treat as an empty cache – no files to scan.
                    _cacheLastBuilt = now;
                    return;
                }

                string[] mdFiles;
                try
                {
                    mdFiles = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
                }
                catch (System.Exception ex)
                {
                    // A permission-denied subfolder or a too-long descendant path fails the whole
                    // recursive enumeration. Degrade to an empty cache rather than throwing.
                    Logger.Instance.Warning($"EmailProcessor: could not enumerate inbox for dedup cache: {ex.Message}");
                    _cacheLastBuilt = now;
                    return;
                }

                foreach (var file in mdFiles)
                {
                    bool inFrontMatter = false;
                    System.Collections.Generic.IEnumerable<string> lines;
                    try
                    {
                        // Materialize so a mid-enumeration read error on one file is caught here and
                        // skipped, instead of aborting the entire cache build.
                        lines = File.ReadAllLines(file);
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Warning($"EmailProcessor: skipping unreadable note during dedup scan '{file}': {ex.Message}");
                        continue;
                    }

                    foreach (var line in lines)
                    {
                        if (line.Trim() == "---")
                        {
                            if (!inFrontMatter)
                            {
                                inFrontMatter = true;
                                continue;
                            }
                            else
                            {
                                // End of frontmatter
                                break;
                            }
                        }
                        if (inFrontMatter)
                        {
                            var trimmed = line.Trim();
                            if (trimmed.StartsWith("internetMessageId:", StringComparison.OrdinalIgnoreCase))
                            {
                                var value = trimmed.Substring("internetMessageId:".Length).Trim().Trim('"');
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    _processedEmailIds.Add(value);
                                }
                            }
                            else if (trimmed.StartsWith("entryId:", StringComparison.OrdinalIgnoreCase))
                            {
                                var value = trimmed.Substring("entryId:".Length).Trim().Trim('"');
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    _processedEmailIds.Add(value);
                                }
                            }
                        }
                    }
                }

                _cacheLastBuilt = now;
            }
        }

        /// <summary>
        /// Checks if an email with the given InternetMessageID or EntryID already exists.
        /// Uses an in-memory cache for O(1) lookups instead of scanning all files.
        /// Cache is automatically rebuilt every 5 minutes to stay synchronized with external changes.
        /// <paramref name="now"/> is the caller's clock (EmailProcessor passes its IClock).
        /// </summary>
        public bool IsDuplicate(string inboxPath, string internetMessageId, string entryId, DateTime now)
        {
            // Ensure cache is built and up-to-date
            EnsureBuilt(inboxPath, now);

            // Fast O(1) lookup in cache
            if (!string.IsNullOrWhiteSpace(internetMessageId) && _processedEmailIds.Contains(internetMessageId))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(entryId) && _processedEmailIds.Contains(entryId))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Adds email IDs to the cache after successfully saving an email note.
        /// This keeps the cache synchronized without requiring a full rebuild.
        /// </summary>
        public void Add(string internetMessageId, string entryId)
        {
            if (!string.IsNullOrWhiteSpace(internetMessageId))
            {
                _processedEmailIds.Add(internetMessageId);
            }
            if (!string.IsNullOrWhiteSpace(entryId))
            {
                _processedEmailIds.Add(entryId);
            }
        }
    }
}
