using System;
using System.Collections.Generic;

namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// A COM-free snapshot used to format email text for the clipboard.
    /// </summary>
    public class ClipboardEmail
    {
        public string EntryId { get; set; } = string.Empty;
        public string ConversationId { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string SenderName { get; set; } = string.Empty;
        public string SenderEmail { get; set; } = string.Empty;
        public List<string> To { get; set; } = new List<string>();
        public List<string> Cc { get; set; } = new List<string>();
        public DateTime SentOn { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; } = string.Empty;
    }
}
