using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Utility wrapper around the .NET <see cref="System.IO"/> APIs that applies the user selected
    /// <see cref="ObsidianSettings"/> when interacting with the file-system.  All routines are kept
    /// <c>virtual</c> to facilitate mocking in unit tests.
    /// </summary>
    public class FileService
    {
        private readonly ObsidianSettings _settings;
        private readonly FileNameSanitizer _fileNameSanitizer;
        private readonly PathLengthBudget _pathBudget;

        public FileService(ObsidianSettings settings)
        {
            _settings = settings;
            _fileNameSanitizer = new FileNameSanitizer();
            _pathBudget = new PathLengthBudget();
        }

        public virtual ObsidianSettings GetSettings()
        {
            return _settings;
        }

        /// <summary>
        /// Writes the supplied string to <paramref name="filePath"/> using UTF-8 *without* emitting a BOM.
        /// The directory hierarchy is created on-the-fly when it does not yet exist.
        /// </summary>
        /// <param name="filePath">Absolute path of the file that should be created or overwritten.</param>
        /// <param name="content">String content that will become the file body.</param>
        public virtual void WriteUtf8File(string filePath, string content)
        {
            // Ensure directory exists
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Write atomically: fully materialize the content in a sibling temp file, then move it
            // into place. A direct overwrite truncates the existing note the instant the stream
            // opens, so an interrupted write (disk full, vault drive drops, lock) would leave a
            // previously-good note empty/truncated. Temp-then-move preserves the old file until the
            // new content is complete.
            string tempPath = filePath + ".slingtmp";
            try
            {
                using (var writer = new StreamWriter(tempPath, false, new UTF8Encoding(false)))
                {
                    writer.Write(content);
                }

                if (File.Exists(filePath))
                {
                    // File.Replace preserves the destination on failure and is atomic where supported.
                    File.Replace(tempPath, filePath, null);
                }
                else
                {
                    File.Move(tempPath, filePath);
                }
            }
            finally
            {
                // Clean up the temp file if the move/replace never happened.
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }

        /// <summary>
        /// Appends the supplied string to an existing file at <paramref name="filePath"/> using UTF-8 *without* emitting a BOM.
        /// </summary>
        /// <param name="filePath">Absolute path of the file to append to.</param>
        /// <param name="content">String content to append.</param>
        public virtual void AppendToFile(string filePath, string content)
        {
            if (!File.Exists(filePath))
            {
                // If file doesn't exist, create it instead
                WriteUtf8File(filePath, content);
                return;
            }

            // Append to file with UTF-8 encoding without BOM
            using (var writer = new StreamWriter(filePath, true, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        /// <summary>
        /// Sanitises a string so that it can safely be used as part of a Windows filename.  
        /// Besides removing invalid characters, the method also applies the user configured
        /// <see cref="ObsidianSettings.SubjectCleanupPatterns"/>.
        /// </summary>
        /// <param name="input">Any raw subject or name string.</param>
        /// <returns>A cleaned filename segment with problematic characters replaced by underscores.</returns>
        public virtual string CleanFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            string cleaned = input;

            // First pass - apply all settings cleanup patterns. (Settings-driven, kept here so
            // CleanFileName remains a one-stop shop for callers that pass raw subjects.)
            // Each pattern is user-supplied and may be invalid (loadable via hand-edited JSON that
            // bypasses Validate) or catastrophically backtracking, so guard both: a timeout bounds
            // ReDoS and ArgumentException skips a bad pattern — either way the sling still completes.
            // This mirrors SubjectCleanerService / FilenameSubjectNormalizer.
            foreach (string pattern in _settings.SubjectCleanupPatterns ?? new List<string>())
            {
                try
                {
                    cleaned = Regex.Replace(cleaned, pattern, "", RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(200));
                }
                catch (ArgumentException ex)
                {
                    Logger.Instance.Warning($"FileService.CleanFileName: skipping invalid cleanup pattern '{pattern}': {ex.Message}");
                }
                catch (RegexMatchTimeoutException ex)
                {
                    Logger.Instance.Warning($"FileService.CleanFileName: cleanup pattern '{pattern}' timed out, skipping: {ex.Message}");
                }
            }

            // Second pass - delegate the pure filesystem-safety logic (invalid chars, prefix
            // strip, separator collapse) to the unit-tested FileNameSanitizer.
            return _fileNameSanitizer.Sanitize(cleaned);
        }

        /// <summary>
        /// Returns the longest note name (without <paramref name="suffix"/>) that keeps every path the note
        /// can produce inside the Windows path limits: the note itself and, with
        /// <see cref="AttachmentStorageMode.SubfolderPerNote"/>, its attachment folder, the attachments in it
        /// and — when <paramref name="mayGroupWithAttachments"/> — the note moved into that folder.
        /// </summary>
        /// <param name="directory">Folder the note is written to.</param>
        /// <param name="suffix">Everything that follows the name, including the extension (e.g. <c>".md"</c>).</param>
        /// <param name="mayGroupWithAttachments">Whether the note can be moved into its attachment folder
        /// (<see cref="ObsidianSettings.GroupNoteWithAttachments"/>; never inside a thread folder).</param>
        public virtual int MaxNoteNameLength(string directory, string suffix, bool mayGroupWithAttachments)
        {
            int directoryLength = (directory ?? string.Empty).Length;
            int max = _pathBudget.MaxNameLength(directory, suffix);
            if (_settings != null && _settings.AttachmentStorageMode == AttachmentStorageMode.SubfolderPerNote)
            {
                max = Math.Min(max, PathLengthBudget.MaxDirectoryPath - directoryLength - 1);
                max = Math.Min(max, PathLengthBudget.MaxFilePath - directoryLength - 2 - PathLengthBudget.AttachmentNameReserve);
                if (mayGroupWithAttachments && _settings.GroupNoteWithAttachments)
                {
                    // directory\name\name.md repeats the name.
                    max = Math.Min(max, (PathLengthBudget.MaxFilePath - directoryLength - 2 - (suffix ?? string.Empty).Length) / 2);
                }
            }
            return max;
        }

        /// <summary>
        /// Shortens <paramref name="name"/> so the paths described by <see cref="MaxNoteNameLength"/> fit.
        /// Names that already fit are returned unchanged.
        /// </summary>
        public virtual string FitNoteName(string directory, string name, string suffix, bool mayGroupWithAttachments)
        {
            return _pathBudget.Fit(name, MaxNoteNameLength(directory, suffix, mayGroupWithAttachments));
        }

        /// <summary>
        /// Opens File Explorer with <paramref name="filePath"/> selected, for people who keep the notes
        /// as plain Markdown files without Obsidian. Failures are logged, never thrown.
        /// </summary>
        public virtual void ShowInExplorer(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + filePath + "\"");
            }
            catch (System.Exception ex) when (ex is System.ComponentModel.Win32Exception || ex is InvalidOperationException)
            {
                Logger.Instance.Warning($"FileService.ShowInExplorer: could not open Explorer for '{filePath}': {ex.Message}");
            }
        }

        /// <summary>
        /// Launches (or brings to front) Obsidian using its custom URI protocol so that the specified
        /// markdown file is opened.  Path components are URI-escaped and the <c>.md</c> extension is
        /// omitted per Obsidian requirements.
        /// </summary>
        /// <param name="vaultName">Target vault name as defined inside Obsidian.</param>
        /// <param name="filePath">Absolute path to the markdown file inside the vault.</param>
        public void LaunchObsidian(string vaultName, string filePath)
        {
            try
            {
                // Replace backslashes with forward slashes for Obsidian URLs
                string normalizedPath = filePath.Replace('\\', '/');
                
                // Remove file extension if present
                if (normalizedPath.EndsWith(".md"))
                {
                    normalizedPath = normalizedPath.Substring(0, normalizedPath.Length - 3);
                }
                
                // Create and launch the Obsidian URL
                string obsidianUrl = $"obsidian://open?vault={Uri.EscapeDataString(vaultName)}&file={Uri.EscapeDataString(normalizedPath)}";
                System.Diagnostics.Process.Start(obsidianUrl);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to launch Obsidian: {ex.Message}", "SlingMD", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Ensures that <paramref name="path"/> exists on disk.  When necessary, all missing intermediate
        /// directories are created.
        /// </summary>
        /// <param name="path">Directory path to check.</param>
        /// <returns><c>true</c> when the directory either already existed or could be created.</returns>
        public virtual bool EnsureDirectoryExists(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public virtual string GetInboxPath() => _settings.GetInboxPath();
    }
} 