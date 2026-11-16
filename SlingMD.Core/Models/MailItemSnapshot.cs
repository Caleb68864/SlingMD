namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Snapshot of mail item properties for decision logic.
    /// Pure DTO with no Outlook Interop dependencies.
    /// </summary>
    public class MailItemSnapshot
    {
        /// <summary>
        /// Gets or sets the sender's email address.
        /// </summary>
        public string SenderEmail { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the sender's domain (extracted from email).
        /// </summary>
        public string SenderDomain { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the email subject line.
        /// </summary>
        public string Subject { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the categories assigned to the mail item.
        /// </summary>
        public string Categories { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets whether the mail item is flagged.
        /// </summary>
        public bool IsFlagged { get; set; }

        /// <summary>
        /// Gets or sets the name of the folder containing the mail item.
        /// </summary>
        public string FolderName { get; set; } = string.Empty;
    }
}
