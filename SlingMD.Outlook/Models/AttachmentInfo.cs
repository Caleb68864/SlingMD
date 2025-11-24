using System.Collections.Generic;

namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Contains information about attachments processed from an email.
    /// </summary>
    public class AttachmentInfo
    {
        public int TotalCount { get; set; }
        public bool HasAttachments => TotalCount > 0;
        public List<SavedAttachment> SavedAttachments { get; set; }

        public AttachmentInfo()
        {
            SavedAttachments = new List<SavedAttachment>();
        }
    }

    /// <summary>
    /// Represents a single saved attachment file.
    /// </summary>
    public class SavedAttachment
    {
        public string Filename { get; set; }
        public string FullPath { get; set; }
        public bool IsInline { get; set; }
        public long SizeBytes { get; set; }
    }

    /// <summary>
    /// Defines where attachments should be stored in the vault.
    /// </summary>
    public enum AttachmentStorageMode
    {
        /// <summary>
        /// Store attachments in same folder as the email note.
        /// Example: Inbox/Email.md + Inbox/image.png
        /// </summary>
        SameAsNote = 0,

        /// <summary>
        /// Create a subfolder per note to store its attachments.
        /// Example: Inbox/Email.md → Inbox/Email/image.png
        /// </summary>
        SubfolderPerNote = 1,

        /// <summary>
        /// Store all attachments in a centralized folder organized by date.
        /// Example: Attachments/2025-01/image.png
        /// </summary>
        Centralized = 2
    }
}
