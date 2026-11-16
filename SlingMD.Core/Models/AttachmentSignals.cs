namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// MAPI-derived signals used to classify an attachment as inline (part of the
    /// rendered HTML body) versus a real, user-facing attachment.
    /// </summary>
    internal struct AttachmentSignals
    {
        public bool IsEmbeddedItem;
        public bool IsHidden;
        public bool HasMhtmlRef;
        public string ContentId;
    }
}
