using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Provides helper routines for detecting email conversation threads, generating folder/note names,
    /// manipulating files that belong to a thread and updating the thread summary note.
    /// All publicly exposed members are safe for unit-testing and are free of any Outlook UI dependencies.
    /// </summary>
    public class ThreadService
    {
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ObsidianSettings _settings;
        private readonly ThreadIdHasher _threadIdHasher;
        private readonly FrontmatterReader _frontmatter;
        private readonly LegacyFilenameStripper _legacyFilenameStripper;
        private readonly UniqueFilenameResolver _uniqueFilenameResolver = new UniqueFilenameResolver();

        public ThreadService(FileService fileService, TemplateService templateService, ObsidianSettings settings)
        {
            _fileService = fileService;
            _templateService = templateService;
            _settings = settings;
            // Use a fallback settings instance when null so "store-only" constructor pattern
            // is preserved (exceptions only surface when methods are actually called).
            _threadIdHasher = new ThreadIdHasher(new SubjectCleanerService(settings ?? new ObsidianSettings()));
            _frontmatter = new FrontmatterReader();
            _legacyFilenameStripper = new LegacyFilenameStripper();
        }

        /// <summary>
        /// Derives a stable 20-character hash that uniquely identifies an Outlook conversation.
        /// The method tries several strategies (ConversationTopic, PR_CONVERSATION_INDEX, normalised subject)
        /// and falls back to a random GUID segment if everything else fails. Using 20 characters instead of 16
        /// significantly reduces the probability of hash collisions while remaining compact.
        /// </summary>
        /// <param name="mail">The <see cref="MailItem"/> for which to obtain the conversation id.</param>
        /// <returns>A 20-character hexadecimal string suitable for use as an identifier inside vault front-matter.</returns>
        public string GetConversationId(MailItem mail)
        {
            try
            {
                // Strategy 1: Hash the conversation topic (most reliable for threading).
                if (!string.IsNullOrEmpty(mail.ConversationTopic))
                {
                    return _threadIdHasher.Hash(mail.ConversationTopic);
                }

                // Strategy 2: PR_CONVERSATION_INDEX — Exchange-native thread bytes. Different hash shape
                // than ThreadIdHasher; preserved as a separate strategy because its ID is not derived
                // from the subject line.
                byte[] conversationIndex = (byte[])mail.PropertyAccessor.GetProperty(MapiPropertyTags.PrConversationIndex);

                if (conversationIndex != null && conversationIndex.Length >= 22)
                {
                    return BitConverter.ToString(conversationIndex.Take(22).ToArray())
                        .Replace("-", "").Substring(0, 20);
                }

                // Strategy 3: Fall back to hashing the subject line.
                return _threadIdHasher.Hash(mail.Subject ?? string.Empty);
            }
            catch
            {
                return Guid.NewGuid().ToString("N").Substring(0, 20);
            }
        }

        /// <summary>
        /// Returns a folder-friendly thread name (no timestamp)
        /// </summary>
        /// <param name="mail">The email that belongs to the thread.</param>
        /// <param name="cleanSubject">A sanitised subject line, typically produced by <see cref="FileService.CleanFileName(string)"/>.</param>
        /// <param name="firstSender">Short name of the first sender in the thread.</param>
        /// <param name="firstRecipient">Short name of the first recipient ("To") in the thread.</param>
        /// <returns>The folder-friendly name without any leading "0-" prefix.</returns>
        public string GetThreadFolderName(MailItem mail, string cleanSubject, string firstSender, string firstRecipient)
        {
            string threadSubject = !string.IsNullOrEmpty(mail.ConversationTopic)
                ? mail.ConversationTopic
                : mail.Subject;
            threadSubject = _fileService.CleanFileName(threadSubject);
            if (threadSubject.Length > 50)
            {
                threadSubject = threadSubject.Substring(0, 47) + "...";
            }
            firstSender = _fileService.CleanFileName(firstSender);
            firstRecipient = _fileService.CleanFileName(firstRecipient);
            return $"{threadSubject.Trim()}-{firstSender}-{firstRecipient}".Replace("--", "-");
        }

        /// <summary>
        /// Returns a note-friendly thread name (with timestamp)
        /// </summary>
        /// <param name="mail">The email that belongs to the thread.</param>
        /// <param name="cleanSubject">A sanitised subject line, typically produced by <see cref="FileService.CleanFileName(string)"/>.</param>
        /// <param name="firstSender">Short name of the first sender in the thread.</param>
        /// <param name="firstRecipient">Short name of the first recipient ("To") in the thread.</param>
        /// <returns>The folder-friendly name without any leading "0-" prefix.</returns>
        public string GetThreadNoteName(MailItem mail, string cleanSubject, string firstSender, string firstRecipient)
        {
            string threadSubject = !string.IsNullOrEmpty(mail.ConversationTopic)
                ? mail.ConversationTopic
                : mail.Subject;
            threadSubject = _fileService.CleanFileName(threadSubject);
            if (threadSubject.Length > 50)
            {
                threadSubject = threadSubject.Substring(0, 47) + "...";
            }
            firstSender = _fileService.CleanFileName(firstSender);
            firstRecipient = _fileService.CleanFileName(firstRecipient);
            string timestamp = mail.ReceivedTime.ToString("yyyy-MM-dd HHmm", System.Globalization.CultureInfo.InvariantCulture);
            return $"{timestamp}-{threadSubject.Trim()}-{firstSender}-{firstRecipient}".Replace("--", "-");
        }

        /// <summary>
        /// Generates or refreshes the <c>0-threadnote.md</c> summary file that lives inside a thread folder.
        /// </summary>
        /// <param name="threadFolderPath">Full path to the thread folder inside the vault inbox.</param>
        /// <param name="threadNotePath">Full path to the summary note file.</param>
        /// <param name="conversationId">Identifier created by <see cref="GetConversationId"/>.</param>
        /// <param name="threadNoteName">Base name (without <c>0-</c> prefix) for the summary file.</param>
        /// <param name="mail">The current email being processed – used only for the title.</param>
        public Task UpdateThreadNote(string threadFolderPath, string threadNotePath, string conversationId, string threadNoteName, MailItem mail)
        {
            string threadTitle = mail.ConversationTopic ?? mail.Subject;
            threadTitle = _fileService.CleanFileName(threadTitle);

            string vaultRoot = Path.Combine(_settings.VaultBasePath, _settings.VaultName);
            string folderPath = string.Empty;
            if (threadFolderPath.StartsWith(vaultRoot, StringComparison.OrdinalIgnoreCase))
            {
                folderPath = threadFolderPath.Substring(vaultRoot.Length)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    .Replace(Path.DirectorySeparatorChar, '/');
            }

            ThreadTemplateContext context = new ThreadTemplateContext
            {
                Title = threadTitle,
                ThreadId = conversationId,
                FolderPath = folderPath
            };

            string content = _templateService.RenderThreadContent(context);
            _fileService.WriteUtf8File(threadNotePath, content);

            return Task.CompletedTask;
        }

        /// <summary>
        /// Moves an email note that was originally written to the inbox into the designated thread folder.
        /// Any existing file with the same name will be overwritten.
        /// The method also appends a deterministic suffix based on the <c>internetMessageId</c> or <c>entryId</c>
        /// found in the note front-matter to avoid collisions between messages that share the same timestamp.
        /// Includes path length validation to prevent exceeding Windows' 260-character path limit by truncating
        /// the email ID suffix and/or base filename as necessary.
        /// </summary>
        /// <param name="emailPath">Absolute path of the markdown file to move.</param>
        /// <param name="threadFolderPath">Destination folder for the thread.</param>
        /// <param name="threadNoteName">Thread name (without the <c>0-</c> prefix). When supplied, the moved
        /// note receives the <c>threadNote</c> front-matter field that notes written straight into a thread get.</param>
        /// <returns>The absolute path of the moved file in its new location.</returns>
        public string MoveToThreadFolder(string emailPath, string threadFolderPath, string threadNoteName = null)
        {
            string fileName = Path.GetFileName(emailPath);
            string newFileName = fileName;
            // Add email id as a temporary suffix to avoid overwriting
            string emailId = null;
            bool inFrontMatter = false;
            foreach (var line in File.ReadLines(emailPath))
            {
                if (line.Trim() == "---")
                {
                    if (!inFrontMatter) { inFrontMatter = true; continue; }
                    else break;
                }
                if (inFrontMatter && line.Trim().StartsWith("internetMessageId:", StringComparison.OrdinalIgnoreCase))
                {
                    emailId = line.Trim().Substring("internetMessageId:".Length).Trim().Trim('"');
                    break;
                }
                if (inFrontMatter && line.Trim().StartsWith("entryId:", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(emailId))
                {
                    emailId = line.Trim().Substring("entryId:".Length).Trim().Trim('"');
                }
            }
            if (!string.IsNullOrEmpty(emailId))
            {
                // Sanitize emailId for filename (alphanumeric only)
                string safeId = new string(emailId.Where(char.IsLetterOrDigit).ToArray());
                string nameNoExt = Path.GetFileNameWithoutExtension(fileName);
                string ext = Path.GetExtension(fileName);

                // Windows path limit is 260 characters. Reserve 20 chars for safety margin.
                // Calculate how much space we have for the filename
                int maxPathLength = 240;
                int folderPathLength = threadFolderPath.Length + 1; // +1 for directory separator
                int availableForFilename = maxPathLength - folderPathLength - ext.Length;

                // Truncate safeId if necessary
                string fullFilename = $"{nameNoExt}-eid{safeId}";
                if (fullFilename.Length > availableForFilename)
                {
                    // Truncate the email ID, keeping at least the first 8 characters for uniqueness
                    int maxIdLength = availableForFilename - nameNoExt.Length - 4; // -4 for "-eid"
                    if (maxIdLength < 8)
                    {
                        // If even 8 chars won't fit, truncate the base name too
                        int maxBaseLength = availableForFilename - 12; // -12 for "-eid" + 8 char ID
                        if (maxBaseLength > 0)
                        {
                            nameNoExt = nameNoExt.Substring(0, Math.Min(nameNoExt.Length, maxBaseLength));
                        }
                        maxIdLength = 8;
                    }
                    safeId = safeId.Substring(0, Math.Min(safeId.Length, maxIdLength));
                }

                newFileName = $"{nameNoExt}-eid{safeId}{ext}";
            }

            string threadPath = Path.Combine(threadFolderPath, newFileName);

            // Final validation: ensure path doesn't exceed limit
            if (threadPath.Length > 240)
            {
                // Emergency truncation - this should rarely happen
                string nameNoExt = Path.GetFileNameWithoutExtension(newFileName);
                string ext = Path.GetExtension(newFileName);
                int maxFilenameLength = 240 - threadFolderPath.Length - ext.Length - 1;
                if (maxFilenameLength > 10)
                {
                    // Clamp the substring length to the actual name length so a short name with a
                    // large budget can't throw ArgumentOutOfRangeException.
                    nameNoExt = nameNoExt.Substring(0, Math.Min(nameNoExt.Length, maxFilenameLength));
                    newFileName = $"{nameNoExt}{ext}";
                    threadPath = Path.Combine(threadFolderPath, newFileName);
                }
                else
                {
                    // The folder itself is so deep that no filename fits the budget. Fall back to a
                    // minimal fixed-length name rather than moving a >260-char path (which throws
                    // PathTooLongException and aborts the whole thread export).
                    string ext2 = Path.GetExtension(newFileName);
                    string shortId = new string(new string(newFileName.Where(char.IsLetterOrDigit).ToArray())
                        .Reverse().Take(8).Reverse().ToArray());
                    if (string.IsNullOrEmpty(shortId)) shortId = "note";
                    newFileName = $"{shortId}{ext2}";
                    threadPath = Path.Combine(threadFolderPath, newFileName);
                }
            }

            _fileService.EnsureDirectoryExists(threadFolderPath);
            try
            {
                if (File.Exists(threadPath))
                {
                    // Two distinct notes can truncate to the same -eid name under the path budget.
                    // Never delete the incumbent — move to a "_N"-suffixed free name instead.
                    string resolved = _uniqueFilenameResolver.Resolve(threadFolderPath, newFileName, File.Exists);
                    if (resolved == null)
                    {
                        Helpers.Logger.Instance.Warning(
                            $"ThreadService.MoveToThreadFolder: no free name for '{newFileName}' in '{threadFolderPath}'; leaving note in inbox.");
                        return emailPath;
                    }
                    threadPath = resolved;
                }
                File.Move(emailPath, threadPath);
            }
            catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // A locked source/target (open in Obsidian, transient AV) must not abort the whole
                // thread export. Leave the note in the inbox and return its current path so the
                // caller keeps working with a real file, consistent with ResuffixThreadNotes.
                Helpers.Logger.Instance.Warning($"ThreadService.MoveToThreadFolder: could not move '{emailPath}' to '{threadPath}': {ex.Message}");
                return File.Exists(emailPath) ? emailPath : threadPath;
            }

            // The note was written as a standalone email, so bring along what it owns and give it
            // the field that ties it to the thread summary.
            RelocateNoteCompanions(emailPath, threadPath);
            if (!string.IsNullOrEmpty(threadNoteName))
            {
                try
                {
                    string content = File.ReadAllText(threadPath);
                    string updated = AddThreadNoteField(content, threadNoteName);
                    if (updated != content)
                    {
                        _fileService.WriteUtf8File(threadPath, updated);
                    }
                }
                catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    Helpers.Logger.Instance.Warning($"ThreadService.MoveToThreadFolder: could not add threadNote to '{threadPath}': {ex.Message}");
                }
            }
            return threadPath;
        }

        /// <summary>
        /// Inserts a <c>threadNote</c> field at the end of the note's front-matter unless one is already present.
        /// </summary>
        internal static string AddThreadNoteField(string content, string threadNoteName)
        {
            if (string.IsNullOrEmpty(content) || string.IsNullOrEmpty(threadNoteName) || !content.StartsWith("---"))
            {
                return content;
            }

            Match closing = Regex.Match(content, @"\r?\n---[ \t]*(?:\r?\n|$)");
            if (!closing.Success)
            {
                return content;
            }

            string frontMatter = content.Substring(0, closing.Index);
            if (Regex.IsMatch(frontMatter, @"(?:^|\n)threadNote:", RegexOptions.IgnoreCase))
            {
                return content;
            }

            string newLine = closing.Value.StartsWith("\r") ? "\r\n" : "\n";
            return frontMatter + newLine + $"threadNote: \"[[0-{threadNoteName}]]\"" + content.Substring(closing.Index);
        }

        /// <summary>
        /// Keeps a note's attachments and self-references intact after the note file has been moved
        /// or renamed from <paramref name="oldNotePath"/> to <paramref name="newNotePath"/>:
        /// the per-note attachment folder follows the note under its new name, attachments stored
        /// beside the note follow it into a new directory, and links inside the note are rewritten
        /// to match. Failures are logged and never abort the export.
        /// </summary>
        internal void RelocateNoteCompanions(string oldNotePath, string newNotePath)
        {
            try
            {
                string oldDir = Path.GetDirectoryName(oldNotePath);
                string newDir = Path.GetDirectoryName(newNotePath);
                string oldName = Path.GetFileNameWithoutExtension(oldNotePath);
                string newName = Path.GetFileNameWithoutExtension(newNotePath);
                bool sameDir = string.Equals(oldDir, newDir, StringComparison.OrdinalIgnoreCase);
                if (sameDir && string.Equals(oldName, newName, StringComparison.Ordinal))
                {
                    return;
                }

                string content = File.ReadAllText(newNotePath);
                string updated = content;

                string oldFolder = Path.Combine(oldDir, oldName);
                string newFolder = Path.Combine(newDir, newName);

                // The note was grouped with its attachments (it lived in a folder of its own name).
                // What is left in that folder becomes the note's attachment folder at the new
                // location, and the bare links that worked while they were siblings gain the folder.
                bool wasGrouped = !sameDir
                    && string.Equals(Path.GetFileName(oldDir), oldName, StringComparison.OrdinalIgnoreCase);
                if (wasGrouped)
                {
                    if (!Directory.EnumerateFileSystemEntries(oldDir).Any())
                    {
                        Directory.Delete(oldDir);
                    }
                    else if (!Directory.Exists(newFolder)
                        && !File.Exists(newFolder)
                        && Directory.GetFiles(oldDir, "*.md", SearchOption.TopDirectoryOnly).Length == 0)
                    {
                        Directory.Move(oldDir, newFolder);
                        string newEscaped = Uri.EscapeUriString(newName).Replace("#", "%23");
                        foreach (string attachment in Directory.GetFiles(newFolder).Select(Path.GetFileName))
                        {
                            // Inline images are linked URL-escaped in markdown mode; the attachment
                            // list is not.
                            string escaped = Uri.EscapeUriString(attachment).Replace("#", "%23");
                            updated = updated
                                .Replace("[[" + attachment + "]]", "[[" + newName + "/" + attachment + "]]")
                                .Replace("](" + escaped + ")", "](" + newEscaped + "/" + escaped + ")");
                            if (escaped != attachment)
                            {
                                updated = updated.Replace("](" + attachment + ")", "](" + newName + "/" + attachment + ")");
                            }
                        }
                    }
                }
                // Attachment folder named after the note (AttachmentStorageMode.SubfolderPerNote).
                // A folder that holds markdown is a thread folder that happens to share the name.
                else if (Directory.Exists(oldFolder)
                    && !Directory.Exists(newFolder)
                    && !File.Exists(newFolder)
                    && Directory.GetFiles(oldFolder, "*.md", SearchOption.TopDirectoryOnly).Length == 0)
                {
                    Directory.Move(oldFolder, newFolder);
                    if (!string.Equals(oldName, newName, StringComparison.Ordinal))
                    {
                        string oldEscaped = Uri.EscapeUriString(oldName).Replace("#", "%23");
                        string newEscaped = Uri.EscapeUriString(newName).Replace("#", "%23");
                        updated = updated
                            .Replace("[[" + oldName + "/", "[[" + newName + "/")
                            .Replace("](" + oldName + "/", "](" + newName + "/")
                            .Replace("](" + oldEscaped + "/", "](" + newEscaped + "/");
                    }
                }

                // Attachments stored beside the note (AttachmentStorageMode.SameAsNote).
                if (!sameDir && !wasGrouped)
                {
                    foreach (string target in FindSameFolderLinkTargets(updated))
                    {
                        string source = Path.Combine(oldDir, target);
                        string destination = Path.Combine(newDir, target);
                        if (File.Exists(source) && !File.Exists(destination))
                        {
                            File.Move(source, destination);
                        }
                    }
                }

                // The Obsidian task block links back to the note by file name.
                if (!string.Equals(oldName, newName, StringComparison.Ordinal))
                {
                    updated = updated
                        .Replace("[[" + oldName + "]]", "[[" + newName + "]]")
                        .Replace("[[" + oldName + "|", "[[" + newName + "|");
                }

                if (updated != content)
                {
                    _fileService.WriteUtf8File(newNotePath, updated);
                }
            }
            catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
            {
                Helpers.Logger.Instance.Warning(
                    $"ThreadService: could not carry attachments from '{oldNotePath}' to '{newNotePath}': {ex.Message}");
            }
        }

        private static readonly Regex WikilinkTargetRegex = new Regex(@"\[\[([^\]\|#]+)(?:[\|#][^\]]*)?\]\]", RegexOptions.Compiled);
        private static readonly Regex MarkdownLinkTargetRegex = new Regex(@"\]\(([^)\s]+)\)", RegexOptions.Compiled);

        /// <summary>
        /// Returns the bare file names (no folder segment, not markdown) that the note links to.
        /// </summary>
        private static IEnumerable<string> FindSameFolderLinkTargets(string content)
        {
            HashSet<string> targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in WikilinkTargetRegex.Matches(content))
            {
                targets.Add(match.Groups[1].Value.Trim());
            }
            foreach (Match match in MarkdownLinkTargetRegex.Matches(content))
            {
                targets.Add(Uri.UnescapeDataString(match.Groups[1].Value));
            }

            char[] invalidChars = Path.GetInvalidFileNameChars();
            return targets
                .Where(target => target.Length > 0 && target.IndexOfAny(invalidChars) < 0)
                .Where(target => !string.IsNullOrEmpty(Path.GetExtension(target))
                    && !string.Equals(Path.GetExtension(target), ".md", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// Scans the inbox (and any existing thread folders) for notes that belong to the supplied conversation id.  
        /// Returns information that helps the caller decide whether the current email should be grouped into 
        /// an existing thread.
        /// </summary>
        /// <param name="conversationId">Identifier created by <see cref="GetConversationId"/>.</param>
        /// <param name="inboxPath">Absolute path to the vault inbox folder.</param>
        /// <returns>Tuple containing: <c>hasExistingThread</c>, the name of the thread folder the conversation
        /// already lives in (<c>null</c> when it has none yet), the earliest email date in the thread and a
        /// count of matching messages.</returns>
        public (bool hasExistingThread, string earliestEmailThreadName, DateTime? earliestEmailDate, int emailCount) 
            FindExistingThread(string conversationId, string inboxPath)
        {
            bool hasExistingThread = false;
            DateTime? earliestEmailDate = null;
            string earliestEmailThreadName = null;
            int emailCount = 0;
            DateTime? threadFolderDate = null;
            List<string> matchingFiles = new List<string>(); // Track matching files for debugging

            try
            {
                // Guard against a missing inbox folder (e.g. first run before any emails have been exported).
                if (!Directory.Exists(inboxPath))
                {
                    return (false, null, null, 0);
                }

                // Get all markdown files from the inbox and subfolders
                var files = Directory.GetFiles(inboxPath, "*.md", SearchOption.AllDirectories);
                
                // Search through each file for the thread ID
                foreach (var file in files)
                {
                    try
                    {
                        string emailContent = File.ReadAllText(file);
                        string fileThreadId = _frontmatter.ExtractThreadId(emailContent);

                        // If this file belongs to the conversation thread
                        if (fileThreadId == conversationId)
                        {
                            // Compare full, separator-normalized paths case-insensitively so a casing
                            // or trailing-slash difference doesn't misclassify an inbox note.
                            string directory = Path.GetDirectoryName(file);
                            bool inInboxRoot = string.Equals(
                                Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar),
                                Path.GetFullPath(inboxPath).TrimEnd(Path.DirectorySeparatorChar),
                                StringComparison.OrdinalIgnoreCase);

                            // Notes filed in any other subfolder (e.g. a batch-sling folder) are not
                            // part of a thread: adopting that folder as the thread would rename every
                            // note in it.
                            bool inNoteFolder = !inInboxRoot && IsNoteFolder(file, inboxPath);
                            if (!inInboxRoot && !inNoteFolder && !IsThreadFolder(directory, inboxPath))
                            {
                                continue;
                            }

                            // A note grouped with its attachments is still a standalone inbox note.
                            inInboxRoot = inInboxRoot || inNoteFolder;

                            hasExistingThread = true;
                            emailCount++; // Increment email count for each matching email
                            matchingFiles.Add(file); // Add to our debugging list

                            // Parse the date to find the earliest email.
                            DateTime emailDate = default(DateTime);
                            string rawDate = _frontmatter.ExtractRawDate(emailContent);
                            bool hasDate = rawDate != null && TryParseThreadDate(rawDate, out emailDate);
                            if (hasDate && (!earliestEmailDate.HasValue || emailDate < earliestEmailDate.Value))
                            {
                                earliestEmailDate = emailDate;
                            }

                            // An existing thread folder always wins, even when an older email of the
                            // conversation is still sitting in the inbox root — otherwise the thread
                            // would be split across two folders. With several folders, the one holding
                            // the oldest note is used. When no folder exists yet the name stays null
                            // and the caller derives it from the email being processed.
                            if (!inInboxRoot
                                && (earliestEmailThreadName == null
                                    || (hasDate && (!threadFolderDate.HasValue || emailDate < threadFolderDate.Value))))
                            {
                                earliestEmailThreadName = Path.GetFileName(directory);
                                if (hasDate)
                                {
                                    threadFolderDate = emailDate;
                                }
                            }
                        }
                    }
                    catch (System.Exception)
                    {
                        // Skip files that can't be read
                        continue;
                    }
                }
                
                // If we found any matches, show more details for debugging
                if (emailCount > 0 && _settings.ShowThreadDebug)
                {
                    string filesList = string.Join("\n", matchingFiles);
                    System.Windows.Forms.MessageBox.Show(
                        $"Found {emailCount} existing emails with thread ID: {conversationId}\n\nFiles:\n{filesList}",
                        "Thread Match Details",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error searching for thread: {ex.Message}", "Thread Search Error", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }

            return (hasExistingThread, earliestEmailThreadName, earliestEmailDate, emailCount);
        }

        /// <summary>
        /// A note folder sits directly inside the inbox, shares the note's name and holds the note
        /// together with its attachments (<see cref="ObsidianSettings.GroupNoteWithAttachments"/>).
        /// </summary>
        private static bool IsNoteFolder(string notePath, string inboxPath)
        {
            string directory = Path.GetFullPath(Path.GetDirectoryName(notePath)).TrimEnd(Path.DirectorySeparatorChar);
            return string.Equals(Path.GetFileName(directory), Path.GetFileNameWithoutExtension(notePath), StringComparison.OrdinalIgnoreCase)
                && string.Equals(
                    Path.GetDirectoryName(directory),
                    Path.GetFullPath(inboxPath).TrimEnd(Path.DirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// A thread folder sits directly inside the inbox and contains a <c>0-</c> summary note.
        /// </summary>
        private static bool IsThreadFolder(string directory, string inboxPath)
        {
            string parent = Path.GetDirectoryName(Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar));
            if (!string.Equals(
                    parent,
                    Path.GetFullPath(inboxPath).TrimEnd(Path.DirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return Directory.GetFiles(directory, "0-*.md", SearchOption.TopDirectoryOnly).Length > 0;
        }

        /// <summary>
        /// Renames every email note in a thread folder so that they receive a chronological <c>yyyy-MM-dd_HHmmss</c>
        /// prefix.  This guarantees a predictable sort order inside Obsidian and prevents filename collisions.
        /// Any temporary <c>-eid*</c> suffixes or outdated numeric suffixes are stripped before each file's
        /// actual base name is preserved. If multiple emails have identical timestamps, a counter suffix (_001, _002, etc.) is added.
        /// </summary>
        /// <param name="threadFolderPath">Full path to the thread folder.</param>
        /// <param name="baseName">Deprecated parameter (no longer used - kept for backward compatibility).</param>
        /// <param name="currentEmailPath">If supplied, returns the new path for the file that corresponds to the current email.</param>
        /// <returns>The new path for <c>currentEmailPath</c> if it was provided; otherwise <c>null</c>.</returns>
        /// <summary>
        /// Parses a date string from a thread note's frontmatter.
        /// Accepts second-precision dates written by the current exporter ("yyyy-MM-dd HH:mm:ss")
        /// and legacy minute-precision dates ("yyyy-MM-dd HH:mm") for backward compatibility.
        /// </summary>
        /// <param name="value">The raw date string extracted from the frontmatter.</param>
        /// <param name="result">The parsed date, if successful.</param>
        /// <returns><c>true</c> if the date could be parsed; otherwise <c>false</c>.</returns>
        public static bool TryParseThreadDate(string value, out DateTime result)
        {
            if (DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out result))
            {
                return true;
            }

            if (DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out result))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Parses a note's <c>date</c> front-matter value the way <see cref="ResuffixThreadNotes"/> does.
        /// </summary>
        internal static bool TryParseNoteDate(string value, out DateTime date)
        {
            return DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out date)
                || DateTime.TryParse(value, out date);
        }

        /// <summary>
        /// Returns the file name (without extension) a note carries inside a thread folder:
        /// a <c>yyyy-MM-dd_HHmmss</c> prefix followed by the base name stripped of its own date and
        /// temporary decorations. Standalone notes are written under this name too, so that a note
        /// keeps its name — and links to it keep resolving — when it later moves into a thread.
        /// </summary>
        public string GetThreadedNoteName(string nameNoExt, DateTime date)
        {
            string datePrefix = date.ToString("yyyy-MM-dd_HHmmss", System.Globalization.CultureInfo.InvariantCulture);
            return $"{datePrefix}_{_legacyFilenameStripper.Strip(nameNoExt)}";
        }

        public string ResuffixThreadNotes(string threadFolderPath, string baseName, string currentEmailPath = null)
        {
            if (!Directory.Exists(threadFolderPath)) return null;

            // Get all markdown files in the thread folder except the thread summary (starts with "0-")
            var files = Directory.GetFiles(threadFolderPath, "*.md", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("0-"))
                .ToList();
            var fileDates = new List<(string file, DateTime date)>();
            foreach (var file in files)
            {
                DateTime? date = null;
                bool inFrontMatter = false;
                foreach (var line in File.ReadLines(file))
                {
                    if (line.Trim() == "---")
                    {
                        if (!inFrontMatter) { inFrontMatter = true; continue; }
                        else break;
                    }
                    if (inFrontMatter && line.Trim().StartsWith("date:", StringComparison.OrdinalIgnoreCase))
                    {
                        var value = line.Trim().Substring("date:".Length).Trim().Trim('"');
                        if (TryParseNoteDate(value, out DateTime parsed))
                            date = parsed;
                        break;
                    }
                }
                if (date.HasValue)
                {
                    fileDates.Add((file, date.Value));
                }
            }
            // Order oldest first so lexicographical date prefix sorts correctly
            fileDates = fileDates.OrderBy(fd => fd.date).ToList();

            string newCurrentPath = null;
            int counter = 0;
            string lastGeneratedName = null;

            foreach (var fd in fileDates)
            {
                // Build new filename with date prefix (yyyy-MM-dd_HHmmss) followed by actual base name from file
                string nameNoExt = Path.GetFileNameWithoutExtension(fd.file);
                string ext = Path.GetExtension(fd.file);

                // Strip SlingMD's legacy filename decorations (eid suffix, -NNN suffix, leading
                // date prefix) — unit-tested helper.
                nameNoExt = _legacyFilenameStripper.Strip(nameNoExt);

                // Include seconds in date prefix to reduce collision probability
                string datePrefix = fd.date.ToString("yyyy-MM-dd_HHmmss", System.Globalization.CultureInfo.InvariantCulture);
                string newName = $"{datePrefix}_{nameNoExt}{ext}";
                string newPath = Path.Combine(threadFolderPath, newName);

                // If we have a collision with the last generated name, add a counter suffix
                if (newName == lastGeneratedName)
                {
                    counter++;
                    newName = $"{datePrefix}_{nameNoExt}_{counter:D3}{ext}";
                    newPath = Path.Combine(threadFolderPath, newName);
                }
                else
                {
                    counter = 0;
                    lastGeneratedName = newName;
                }

                bool moved = true;
                if (!fd.file.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                {
                    // Ensure no conflict before moving — keep bumping the counter until a free name
                    // is found (a single bump could still collide with an existing _NNN from a prior
                    // run), since File.Move refuses to overwrite and would throw.
                    while (File.Exists(newPath))
                    {
                        counter++;
                        newName = $"{datePrefix}_{nameNoExt}_{counter:D3}{ext}";
                        newPath = Path.Combine(threadFolderPath, newName);
                    }

                    try
                    {
                        File.Move(fd.file, newPath);
                        RelocateNoteCompanions(fd.file, newPath);
                    }
                    catch (System.Exception ex)
                    {
                        // Skip this one file rather than aborting the whole resuffix pass (which
                        // would leave the thread folder half-renamed).
                        SlingMD.Outlook.Helpers.Logger.Instance.Warning(
                            $"ThreadService: could not rename '{fd.file}' to '{newPath}': {ex.Message}");
                        moved = false;
                    }
                }

                if (currentEmailPath != null && fd.file.Equals(currentEmailPath, StringComparison.OrdinalIgnoreCase))
                {
                    newCurrentPath = moved ? newPath : fd.file;
                }
            }
            return newCurrentPath;
        }
    }
} 
