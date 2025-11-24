using System;
using System.IO;
using System.Text;
using SlingMD.Outlook.Models;
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

        public FileService(ObsidianSettings settings)
        {
            _settings = settings;
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

            // Write file with UTF-8 encoding without BOM
            using (var writer = new StreamWriter(filePath, false, new UTF8Encoding(false)))
            {
                writer.Write(content);
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

            // First pass - apply all cleanup patterns from settings
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                cleaned = Regex.Replace(cleaned, pattern, "", RegexOptions.IgnoreCase);
            }

            // Replace invalid characters with underscore
            char[] invalidChars = Path.GetInvalidFileNameChars();
            cleaned = string.Join("_", cleaned.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));

            // Replace additional problematic characters
            cleaned = cleaned.Replace("\"", "")
                           .Replace("'", "")
                           .Replace("`", "")
                           .Replace(":", "_")  // Replace colon with underscore to handle "Re: " -> "RE_"
                           .Replace(";", "")
                           .Trim();

            // Second pass - clean up any remaining email prefixes that might have been converted to underscore format
            cleaned = Regex.Replace(cleaned, @"^(?:RE_|FWD_|FW_|Re_|Fwd_)", "", RegexOptions.IgnoreCase);
            
            // Clean up multiple underscores/hyphens
            cleaned = Regex.Replace(cleaned, @"[-_]{2,}", "-");
            
            // Final trim of any remaining leading/trailing separators
            cleaned = cleaned.Trim('-', '_');

            return cleaned;
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