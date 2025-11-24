using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Helpers;

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

        public AttachmentService(ObsidianSettings settings, FileService fileService)
        {
            _settings = settings;
            _fileService = fileService;
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
            try
            {
                attachments = mail.Attachments;
                info.TotalCount = attachments.Count;

                if (info.TotalCount == 0)
                {
                    return info;
                }

                // Determine target folder based on storage mode
                string targetFolder = GetAttachmentFolder(noteFilePath);
                _fileService.EnsureDirectoryExists(targetFolder);

                // Process each attachment
                for (int i = 1; i <= attachments.Count; i++)
                {
                    Attachment attachment = null;
                    try
                    {
                        attachment = attachments[i];
                        bool isInline = IsInlineImage(attachment);

                        // Check if we should save this attachment
                        if ((_settings.SaveInlineImages && isInline) || (_settings.SaveAllAttachments))
                        {
                            SavedAttachment saved = SaveAttachment(attachment, targetFolder, isInline);
                            if (saved != null)
                            {
                                info.SavedAttachments.Add(saved);
                            }
                        }
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
            }

            return info;
        }

        /// <summary>
        /// Determines if an attachment is an inline image based on its properties.
        /// </summary>
        private bool IsInlineImage(Attachment attachment)
        {
            try
            {
                // Check for ContentID (PR_ATTACH_CONTENT_ID)
                const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
                object contentId = attachment.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID);

                if (contentId != null && !string.IsNullOrEmpty(contentId.ToString()))
                {
                    return true;
                }

                // Also check if it's an embedded type (olEmbeddeditem = 4)
                if (attachment.Type == OlAttachmentType.olEmbeddeditem)
                {
                    return true;
                }

                // Check if filename suggests it's an image
                string filename = attachment.FileName ?? "";
                string ext = Path.GetExtension(filename).ToLowerInvariant();
                if (ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".gif" || ext == ".bmp" || ext == ".svg")
                {
                    return true;
                }
            }
            catch
            {
                // If we can't determine, assume it's not inline
            }

            return false;
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

                // Handle filename conflicts by appending numbers
                string fullPath = Path.Combine(targetFolder, safeFilename);
                int counter = 1;
                const int maxRetries = 10;

                while (File.Exists(fullPath) && counter <= maxRetries)
                {
                    safeFilename = $"{nameWithoutExt}_{counter}{extension}";
                    fullPath = Path.Combine(targetFolder, safeFilename);
                    counter++;
                }

                if (counter > maxRetries)
                {
                    Logger.Instance.Error($"Failed to find unique filename for attachment after {maxRetries} attempts: {originalFilename}");
                    return null;
                }

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
                        // File was created between check and save (race condition)
                        // Try with incremented counter
                        saveAttempt++;
                        if (saveAttempt >= maxRetries)
                        {
                            Logger.Instance.Error($"Failed to save attachment after {maxRetries} attempts due to race condition: {originalFilename}");
                            return null;
                        }
                        safeFilename = $"{nameWithoutExt}_{counter}{extension}";
                        fullPath = Path.Combine(targetFolder, safeFilename);
                        counter++;
                    }
                }

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
                    string yearMonth = DateTime.Now.ToString("yyyy-MM");
                    return Path.Combine(vaultPath, attachmentsFolder, yearMonth);

                case AttachmentStorageMode.SameAsNote:
                default:
                    // Same directory as the note
                    return noteDir;
            }
        }

        /// <summary>
        /// Generates an Obsidian wikilink for an attachment.
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
                // Standard markdown format
                return isImage ? $"![{filename}]({filename})" : $"[{filename}]({filename})";
            }
        }
    }
}
