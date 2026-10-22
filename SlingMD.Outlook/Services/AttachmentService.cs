using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Handles extraction and storage of email attachments, including inline images.
    /// Properly releases COM objects to prevent memory leaks.
    /// </summary>
    public class AttachmentService
    {
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly IClock _clock;
        private readonly UniqueFilenameResolver _uniqueFilenameResolver = new UniqueFilenameResolver();

        public AttachmentService(ObsidianSettings settings, FileService fileService, IClock clock = null)
        {
            _settings = settings;
            _fileService = fileService;
            _clock = clock ?? new SystemClock();
        }

        /// <summary>
        /// Processes all attachments from an email and saves them according to settings.
        /// </summary>
        /// <param name="mail">The email containing attachments.</param>
        /// <param name="noteFilePath">Full path to the markdown note file.</param>
        /// <returns>Information about processed attachments.</returns>
        public AttachmentInfo ProcessAttachments(MailItem mail, string noteFilePath)
        {
            AttachmentInfo info = new AttachmentInfo();

            Attachments attachments = null;
            string targetFolder = null;
            try
            {
                attachments = mail.Attachments;
                info.TotalCount = attachments.Count;

                if (info.TotalCount == 0)
                {
                    return info;
                }

                // Determine target folder based on storage mode
                targetFolder = GetAttachmentFolder(noteFilePath);
                _fileService.EnsureDirectoryExists(targetFolder);

                // Process each attachment
                for (int i = 1; i <= attachments.Count; i++)
                {
                    Attachment attachment = null;
                    try
                    {
                        attachment = attachments[i];
                        AttachmentSignals signals = ReadSignals(attachment);
                        bool isInline = ClassifyFromSignals(signals);
                        bool shouldSave = ShouldSave(isInline, _settings);

                        SavedAttachment saved = null;
                        if (shouldSave)
                        {
                            saved = SaveAttachment(attachment, targetFolder, isInline);
                            if (saved != null)
                            {
                                saved.ContentId = signals.ContentId;
                                info.SavedAttachments.Add(saved);
                            }
                        }

                        Logger.Instance.Debug($"Attachment '{attachment.FileName}' classified as {(isInline ? "inline" : "real")}; {(saved != null ? "saved" : "skipped")}.");
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Instance.Error($"Failed to process attachment {i}: {ex.Message}", ex);
                    }
                    finally
                    {
                        // Release individual Attachment COM object
                        if (attachment != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(attachment);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"Failed to process attachments: {ex.Message}", ex);
            }
            finally
            {
                // Release Attachments collection COM object
                if (attachments != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(attachments);
                }

                RemoveEmptyNoteFolder(targetFolder);
            }

            return info;
        }

        /// <summary>
        /// Deletes the per-note attachment folder when nothing ended up in it (every attachment
        /// was skipped or reused from elsewhere), so emails don't leave empty folders in the inbox.
        /// </summary>
        internal void RemoveEmptyNoteFolder(string targetFolder)
        {
            if (_settings.AttachmentStorageMode != AttachmentStorageMode.SubfolderPerNote || string.IsNullOrEmpty(targetFolder))
            {
                return;
            }

            try
            {
                if (Directory.Exists(targetFolder) && !Directory.EnumerateFileSystemEntries(targetFolder).Any())
                {
                    Directory.Delete(targetFolder);
                }
            }
            catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Logger.Instance.Warning($"AttachmentService: could not remove empty attachment folder '{targetFolder}': {ex.Message}");
            }
        }

        /// <summary>
        /// When <see cref="ObsidianSettings.GroupNoteWithAttachments"/> is on and the note's attachment
        /// folder exists, moves the note into that folder and returns its new path. Call after
        /// <see cref="ProcessAttachments"/> and before generating links, so links are built relative
        /// to the note's final location. Returns <paramref name="noteFilePath"/> when nothing moved.
        /// </summary>
        internal string GroupNoteWithAttachments(string noteFilePath)
        {
            if (!_settings.GroupNoteWithAttachments
                || _settings.AttachmentStorageMode != AttachmentStorageMode.SubfolderPerNote
                || string.IsNullOrEmpty(noteFilePath))
            {
                return noteFilePath;
            }

            try
            {
                string folder = GetAttachmentFolder(noteFilePath);
                string groupedPath = Path.Combine(folder, Path.GetFileName(noteFilePath));
                if (!Directory.Exists(folder) || File.Exists(groupedPath))
                {
                    return noteFilePath;
                }

                File.Move(noteFilePath, groupedPath);
                return groupedPath;
            }
            catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Logger.Instance.Warning($"AttachmentService: could not move '{noteFilePath}' into its attachment folder: {ex.Message}");
                return noteFilePath;
            }
        }

        /// <summary>
        /// When <see cref="ObsidianSettings.ReuseDuplicateAttachments"/> is on and a file with identical
        /// content has been saved before, replaces the just-saved file with a single copy in the shared
        /// attachments folder and returns that copy's path. Otherwise returns <paramref name="savedPath"/>.
        /// The earlier note keeps its own copy, so no existing note is rewritten and the shared copy
        /// never moves when a note is later filed into a thread.
        /// </summary>
        internal string ReuseDuplicateIfAny(string savedPath)
        {
            if (!_settings.ReuseDuplicateAttachments || string.IsNullOrEmpty(savedPath) || !File.Exists(savedPath))
            {
                return savedPath;
            }

            try
            {
                string sharedFolder = _settings.GetSharedAttachmentsPath();
                long size = new FileInfo(savedPath).Length;
                string hash = null;

                string shared = FindIdenticalFile(sharedFolder, SearchOption.TopDirectoryOnly, savedPath, size, ref hash);
                if (shared == null)
                {
                    string earlier = FindIdenticalFile(_settings.GetInboxPath(), SearchOption.AllDirectories, savedPath, size, ref hash);
                    if (earlier == null)
                    {
                        return savedPath;
                    }

                    _fileService.EnsureDirectoryExists(sharedFolder);
                    shared = _uniqueFilenameResolver.Resolve(sharedFolder, Path.GetFileName(savedPath), File.Exists);
                    if (shared == null)
                    {
                        return savedPath;
                    }
                    File.Move(savedPath, shared);
                    return shared;
                }

                File.Delete(savedPath);
                return shared;
            }
            catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                Logger.Instance.Warning($"AttachmentService: duplicate check failed for '{savedPath}', keeping its own copy: {ex.Message}");
                return savedPath;
            }
        }

        /// <summary>
        /// Finds a file under <paramref name="folder"/> (other than <paramref name="savedPath"/>) with the
        /// same length and SHA-256 hash. Only same-length candidates are ever hashed.
        /// </summary>
        private static string FindIdenticalFile(string folder, SearchOption searchOption, string savedPath, long size, ref string hash)
        {
            if (!Directory.Exists(folder))
            {
                return null;
            }

            foreach (string candidate in Directory.EnumerateFiles(folder, "*", searchOption))
            {
                if (string.Equals(candidate, savedPath, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(Path.GetExtension(candidate), ".md", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    if (new FileInfo(candidate).Length != size)
                    {
                        continue;
                    }

                    hash = hash ?? ComputeFileHash(savedPath);
                    if (ComputeFileHash(candidate) == hash)
                    {
                        return candidate;
                    }
                }
                catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    // A locked or vanished candidate just isn't a reuse target.
                }
            }

            return null;
        }

        private static string ComputeFileHash(string path)
        {
            using (System.Security.Cryptography.SHA256 sha = System.Security.Cryptography.SHA256.Create())
            using (FileStream stream = File.OpenRead(path))
            {
                return BitConverter.ToString(sha.ComputeHash(stream));
            }
        }

        /// <summary>
        /// Reads MAPI signals from an attachment's COM object. Each read is isolated in
        /// its own try/catch so a single unsupported property leaves that field at its default
        /// instead of failing the whole read.
        /// </summary>
        internal AttachmentSignals ReadSignals(Attachment attachment)
        {
            AttachmentSignals signals = new AttachmentSignals();

            try
            {
                signals.IsEmbeddedItem = attachment.Type == OlAttachmentType.olEmbeddeditem;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"AttachmentService.ReadSignals: could not read attachment Type: {ex.Message}");
            }

            // attachment.PropertyAccessor mints a COM object; obtain it once, reuse it for all three
            // MAPI reads, and release it — the previous code accessed it three times (leaking two)
            // and swallowed every failure with a bare catch (hiding Convert/COM diagnostics).
            PropertyAccessor pa = null;
            try
            {
                pa = attachment.PropertyAccessor;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Debug($"AttachmentService.ReadSignals: PropertyAccessor unavailable: {ex.Message}");
            }

            if (pa != null)
            {
                try
                {
                    object hidden = pa.GetProperty(MapiPropertyTags.PrAttachHidden);
                    signals.IsHidden = hidden != null && Convert.ToBoolean(hidden);
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"AttachmentService.ReadSignals: PrAttachHidden unavailable: {ex.Message}");
                }

                try
                {
                    object flags = pa.GetProperty(MapiPropertyTags.PrAttachFlags);
                    if (flags != null)
                    {
                        int flagsValue = Convert.ToInt32(flags);
                        signals.HasMhtmlRef = (flagsValue & 0x4) != 0;
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"AttachmentService.ReadSignals: PrAttachFlags unavailable: {ex.Message}");
                }

                try
                {
                    object contentId = pa.GetProperty(MapiPropertyTags.PrAttachContentId);
                    signals.ContentId = contentId?.ToString();
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Debug($"AttachmentService.ReadSignals: PrAttachContentId unavailable: {ex.Message}");
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(pa);
            }

            return signals;
        }

        /// <summary>
        /// Pure classification: determines whether an attachment is inline based on its
        /// previously-read MAPI signals.
        /// </summary>
        internal bool ClassifyFromSignals(AttachmentSignals signals)
        {
            return signals.IsEmbeddedItem
                || signals.IsHidden
                || signals.HasMhtmlRef
                || !string.IsNullOrEmpty(signals.ContentId);
        }

        /// <summary>
        /// Pure save decision based on the inline/real classification and current settings.
        /// </summary>
        internal bool ShouldSave(bool isInline, ObsidianSettings settings)
        {
            return (isInline && settings.SaveInlineImages)
                || (!isInline && settings.SaveRealAttachments)
                || settings.SaveAllAttachments;
        }

        /// <summary>Points HTML body cid images at the files saved in the vault.</summary>
        internal string ResolveInlineImageLinks(string content, AttachmentInfo info, string noteFilePath)
        {
            if (string.IsNullOrEmpty(content) || info == null)
            {
                return content;
            }

            foreach (SavedAttachment saved in info.SavedAttachments)
            {
                if (!saved.IsInline || string.IsNullOrWhiteSpace(saved.ContentId))
                {
                    continue;
                }

                string contentId = saved.ContentId.Trim().Trim('<', '>');
                string target = BuildAttachmentLinkTarget(saved.FullPath, noteFilePath);
                string escapedTarget = EscapeLinkTarget(target);
                string pattern = @"!\[([^\]]*)\]\(cid:" + Regex.Escape(contentId) + @"\)";
                content = Regex.Replace(content, pattern, match =>
                    _settings.UseObsidianWikilinks
                        ? $"![[{target}]]"
                        : $"![{match.Groups[1].Value}]({escapedTarget})",
                    RegexOptions.IgnoreCase);
            }

            return content;
        }

        /// <summary>
        /// Saves an attachment to disk and returns information about the saved file.
        /// </summary>
        private SavedAttachment SaveAttachment(Attachment attachment, string targetFolder, bool isInline)
        {
            try
            {
                string originalFilename = attachment.FileName ?? "attachment";
                string safeFilename = _fileService.CleanFileName(originalFilename);

                // If CleanFileName returns empty, use default name
                if (string.IsNullOrWhiteSpace(safeFilename))
                {
                    string originalExt = Path.GetExtension(originalFilename);
                    safeFilename = string.IsNullOrEmpty(originalExt) ? "attachment.dat" : $"attachment{originalExt}";
                }

                // Validate path length (Windows has 260 char limit, use 240 for safety)
                const int maxPathLength = 240;
                string nameWithoutExt = Path.GetFileNameWithoutExtension(safeFilename);
                string extension = Path.GetExtension(safeFilename);

                // Calculate available space for filename
                int folderPathLength = targetFolder.Length;
                int availableForFilename = maxPathLength - folderPathLength - extension.Length - 10; // 10 chars buffer for counter "_999"

                if (availableForFilename < 10)
                {
                    Logger.Instance.Warning($"Path too long for attachment: {targetFolder}. Attachment may fail to save.");
                    availableForFilename = 10; // Minimum filename length
                }

                // Truncate filename if necessary
                if (nameWithoutExt.Length > availableForFilename)
                {
                    nameWithoutExt = nameWithoutExt.Substring(0, availableForFilename);
                    safeFilename = nameWithoutExt + extension;
                    Logger.Instance.Warning($"Truncated attachment filename to {safeFilename} due to path length constraints.");
                }

                // Handle filename conflicts by appending numbers (unit-tested helper).
                string fullPath = _uniqueFilenameResolver.Resolve(targetFolder, safeFilename, File.Exists);
                if (fullPath == null)
                {
                    Logger.Instance.Error($"Failed to find unique filename for attachment after {UniqueFilenameResolver.MaxAttempts} attempts: {originalFilename}");
                    return null;
                }
                safeFilename = Path.GetFileName(fullPath);
                const int maxRetries = 10;

                // Save the attachment with retry logic to handle race conditions
                bool saved = false;
                int saveAttempt = 0;
                while (!saved && saveAttempt < maxRetries)
                {
                    try
                    {
                        attachment.SaveAsFile(fullPath);
                        saved = true;
                    }
                    catch (System.IO.IOException ex) when (ex.Message.Contains("already exists") || ex.HResult == -2147024816)
                    {
                        // File was created between check and save (race condition).
                        // Re-resolve from the current candidate so the resolver can find
                        // the next free name without us hand-tracking a counter.
                        saveAttempt++;
                        if (saveAttempt >= maxRetries)
                        {
                            Logger.Instance.Error($"Failed to save attachment after {maxRetries} attempts due to race condition: {originalFilename}");
                            return null;
                        }
                        string raceResolved = _uniqueFilenameResolver.Resolve(targetFolder, safeFilename, File.Exists);
                        if (raceResolved == null)
                        {
                            Logger.Instance.Error($"Failed to find unique filename during race-recovery for attachment: {originalFilename}");
                            return null;
                        }
                        fullPath = raceResolved;
                        safeFilename = Path.GetFileName(fullPath);
                    }
                }

                fullPath = ReuseDuplicateIfAny(fullPath);
                safeFilename = Path.GetFileName(fullPath);

                // Get file size
                FileInfo fileInfo = new FileInfo(fullPath);

                return new SavedAttachment
                {
                    Filename = safeFilename,
                    FullPath = fullPath,
                    IsInline = isInline,
                    SizeBytes = fileInfo.Length
                };
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"Failed to save attachment {attachment.FileName}: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// Determines the target folder for attachments based on settings and note path.
        /// </summary>
        private string GetAttachmentFolder(string noteFilePath)
        {
            string noteDir = Path.GetDirectoryName(noteFilePath);
            string noteNameWithoutExt = Path.GetFileNameWithoutExtension(noteFilePath);

            switch (_settings.AttachmentStorageMode)
            {
                case AttachmentStorageMode.SubfolderPerNote:
                    // Create subfolder with same name as note
                    return Path.Combine(noteDir, noteNameWithoutExt);

                case AttachmentStorageMode.Centralized:
                    // Use centralized attachments folder with year-month subfolders
                    string vaultPath = _settings.GetFullVaultPath();
                    string attachmentsFolder = _settings.AttachmentsFolder ?? "Attachments";
                    string yearMonth = _clock.Now.ToString("yyyy-MM");
                    return Path.Combine(vaultPath, attachmentsFolder, yearMonth);

                case AttachmentStorageMode.SameAsNote:
                default:
                    // Same directory as the note
                    return noteDir;
            }
        }

        /// <summary>
        /// Generates an Obsidian wikilink for an attachment using only its filename.
        /// Use this overload only when the attachment is stored in the same directory as the note
        /// (i.e. <see cref="AttachmentStorageMode.SameAsNote"/>).
        /// For other storage modes use <see cref="GenerateWikilink(string, string, bool)"/> so that
        /// the correct relative path is included in the link.
        /// </summary>
        /// <param name="filename">The filename of the attachment.</param>
        /// <param name="isImage">Whether this is an image (affects formatting).</param>
        /// <returns>Wikilink string in Obsidian format.</returns>
        public string GenerateWikilink(string filename, bool isImage)
        {
            if (_settings.UseObsidianWikilinks)
            {
                // Obsidian wikilink format
                return isImage ? $"![[{filename}]]" : $"[[{filename}]]";
            }
            else
            {
                // Standard markdown format; spaces and # in the target would break the link.
                string target = EscapeLinkTarget(filename);
                return isImage ? $"![{filename}]({target})" : $"[{filename}]({target})";
            }
        }

        /// <summary>
        /// Generates an Obsidian wikilink for an attachment using the saved file path relative to the note.
        /// For <see cref="AttachmentStorageMode.SameAsNote"/> the link contains only the bare filename,
        /// preserving the existing behavior.  For <see cref="AttachmentStorageMode.SubfolderPerNote"/>
        /// and <see cref="AttachmentStorageMode.Centralized"/> modes the link includes the relative path
        /// from the note's directory so that it resolves correctly inside the vault.
        /// </summary>
        /// <param name="attachmentFullPath">Absolute path to the saved attachment file.</param>
        /// <param name="noteFilePath">Absolute path to the markdown note that will reference the attachment.</param>
        /// <param name="isImage">Whether this is an image (affects formatting).</param>
        /// <returns>Wikilink string in Obsidian format.</returns>
        public string GenerateWikilink(string attachmentFullPath, string noteFilePath, bool isImage)
        {
            string linkTarget = BuildAttachmentLinkTarget(attachmentFullPath, noteFilePath);

            if (_settings.UseObsidianWikilinks)
            {
                return isImage ? $"![[{linkTarget}]]" : $"[[{linkTarget}]]";
            }
            else
            {
                string displayName = Path.GetFileName(attachmentFullPath) ?? linkTarget;
                string escapedTarget = EscapeLinkTarget(linkTarget);
                return isImage
                    ? $"![{displayName}]({escapedTarget})"
                    : $"[{displayName}]({escapedTarget})";
            }
        }

        /// <summary>URL-escapes a Markdown link target the way inline image links are written.</summary>
        internal static string EscapeLinkTarget(string target)
        {
            return Uri.EscapeUriString(target ?? string.Empty).Replace("#", "%23");
        }

        /// <summary>
        /// Builds the "## Attachments" list appended to an email note: one bullet per saved file,
        /// leaving out inline images the body already shows (signature logos, pasted pictures).
        /// Returns an empty string when nothing is left to list.
        /// </summary>
        /// <param name="info">What was saved for the email.</param>
        /// <param name="noteFilePath">The note the links are relative to.</param>
        /// <param name="noteContent">The note's content after <see cref="ResolveInlineImageLinks"/>.</param>
        internal string BuildAttachmentSection(AttachmentInfo info, string noteFilePath, string noteContent)
        {
            if (info == null)
            {
                return string.Empty;
            }

            StringBuilder section = new StringBuilder();
            foreach (SavedAttachment saved in info.SavedAttachments)
            {
                if (saved.IsInline && IsShownInBody(saved, noteFilePath, noteContent))
                {
                    continue;
                }

                section.Append("- ").Append(GenerateWikilink(saved.FullPath, noteFilePath, false)).Append('\n');
            }

            return section.Length == 0 ? string.Empty : "\n\n## Attachments\n\n" + section;
        }

        private bool IsShownInBody(SavedAttachment saved, string noteFilePath, string noteContent)
        {
            if (string.IsNullOrEmpty(noteContent))
            {
                return false;
            }

            string target = BuildAttachmentLinkTarget(saved.FullPath, noteFilePath);
            return noteContent.Contains("![[" + target + "]]")
                || noteContent.Contains("](" + EscapeLinkTarget(target) + ")");
        }

        /// <summary>
        /// Computes the link target for an attachment relative to the note file.
        /// For same-folder storage the bare filename is returned so existing behavior is preserved.
        /// For other modes a forward-slash delimited relative path is returned.
        /// </summary>
        internal string BuildAttachmentLinkTarget(string attachmentFullPath, string noteFilePath)
        {
            if (string.IsNullOrWhiteSpace(attachmentFullPath) || string.IsNullOrWhiteSpace(noteFilePath))
            {
                return Path.GetFileName(attachmentFullPath) ?? string.Empty;
            }

            string noteDir = Path.GetDirectoryName(noteFilePath);
            string attachmentDir = Path.GetDirectoryName(attachmentFullPath);
            string attachmentFilename = Path.GetFileName(attachmentFullPath);

            if (string.IsNullOrEmpty(noteDir) || string.IsNullOrEmpty(attachmentDir))
            {
                return attachmentFilename ?? string.Empty;
            }

            // Shared (reused) attachments are linked by their path from the vault root, which
            // Obsidian resolves from any note — so the link survives the note moving into a thread.
            string vaultRoot = _settings.GetFullVaultPath().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string sharedDir = _settings.GetSharedAttachmentsPath();
            if (string.Equals(attachmentDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), sharedDir, StringComparison.OrdinalIgnoreCase))
            {
                return attachmentFullPath.Substring(vaultRoot.Length)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    .Replace('\\', '/');
            }

            // Normalise to the same case and separator for comparison
            string normNoteDir = noteDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string normAttachDir = attachmentDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            if (string.Equals(normNoteDir, normAttachDir, StringComparison.OrdinalIgnoreCase))
            {
                // Same directory – use bare filename (preserves existing SameAsNote behavior)
                return attachmentFilename;
            }

            // Build a URI-style relative path from the note directory to the attachment file
            // using forward slashes as required by Obsidian wikilinks and standard markdown links.
            // Prefix with "file:///" so that Uri can handle Windows drive-letter paths correctly.
            string noteDirUri = "file:///" + normNoteDir.Replace('\\', '/') + "/";
            string attachFileUri = "file:///" + attachmentFullPath.Replace('\\', '/');

            Uri noteUri = new Uri(noteDirUri, UriKind.Absolute);
            Uri attachUri = new Uri(attachFileUri, UriKind.Absolute);
            string relativePath = Uri.UnescapeDataString(noteUri.MakeRelativeUri(attachUri).ToString());

            return relativePath;
        }
    }
}
