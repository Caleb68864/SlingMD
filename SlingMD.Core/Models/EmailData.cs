using System;
using System.Collections.Generic;

namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// One recipient of an email as plain values: display name, SMTP address (may be empty) and
    /// whether it was on the To or the Cc line.
    /// </summary>
    public class EmailRecipientData
    {
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public bool IsCc { get; set; }
    }

    /// <summary>
    /// Everything the note-building engine reads from an Outlook <c>MailItem</c>, copied out as plain
    /// values so <see cref="SlingMD.Outlook.Services.EmailNoteEngine"/> runs without COM.
    /// The add-in fills it with <c>OutlookEmailReader</c>; tests build it directly.
    /// Values are as Outlook returned them (nulls kept), so the engine applies the same
    /// fallbacks ("No Subject", "Unknown Sender") the processor always did.
    /// </summary>
    public class EmailData
    {
        public string Subject { get; set; }
        public string SenderName { get; set; }

        /// <summary>SMTP address of the sender, as <c>SmtpAddressResolver.ForSender</c> resolved it.</summary>
        public string SenderEmail { get; set; }

        public DateTime ReceivedTime { get; set; }
        public string HtmlBody { get; set; }
        public string Body { get; set; }
        public string EntryId { get; set; }

        /// <summary>PR_INTERNET_MESSAGE_ID, or <c>null</c> when it could not be read.</summary>
        public string InternetMessageId { get; set; }

        /// <summary>Name of the first To recipient, or "Unknown" (thread folder naming).</summary>
        public string FirstToRecipientName { get; set; } = "Unknown";

        /// <summary>Every recipient (To and Cc) in Outlook's order.</summary>
        public List<EmailRecipientData> Recipients { get; set; } = new List<EmailRecipientData>();
    }
}
