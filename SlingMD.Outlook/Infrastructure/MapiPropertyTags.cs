namespace SlingMD.Outlook.Infrastructure
{
    /// <summary>
    /// Centralized MAPI property tag URIs used with Outlook's PropertyAccessor.
    /// Avoids scattered duplicate string literals that can drift (e.g. a single-char
    /// typo in the proptag hex would silently fail at runtime).
    /// Reference: https://learn.microsoft.com/en-us/office/vba/outlook/concepts/forms/referencing-properties-by-namespace
    /// </summary>
    public static class MapiPropertyTags
    {
        /// <summary>
        /// PR_SMTP_ADDRESS — canonical SMTP address of a recipient or contact.
        /// Used to resolve the true email address when <c>Recipient.Address</c> returns
        /// an Exchange distinguished name instead of SMTP.
        /// </summary>
        public const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        /// <summary>
        /// PR_CONVERSATION_INDEX — Exchange-native byte blob identifying the conversation.
        /// First 22 bytes form the thread identifier; subsequent 5-byte tuples represent replies.
        /// </summary>
        public const string PrConversationIndex = "http://schemas.microsoft.com/mapi/proptag/0x0071001F";

        /// <summary>
        /// PR_ATTACH_CONTENT_ID — Content-ID for an attachment. Non-empty value indicates
        /// the attachment is referenced inline from the HTML body.
        /// </summary>
        public const string PrAttachContentId = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";

        /// <summary>
        /// PR_INTERNET_MESSAGE_ID — RFC-2822 Message-ID header. Stable across Outlook stores.
        /// </summary>
        public const string PrInternetMessageId = "http://schemas.microsoft.com/mapi/proptag/0x1035001E";

        /// <summary>
        /// PR_ATTACH_HIDDEN — boolean flag indicating an attachment is hidden from the user
        /// (e.g. an inline-image attachment that should not also be listed as a regular attachment).
        /// </summary>
        public const string PrAttachHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";

        /// <summary>
        /// PR_ATTACH_FLAGS — bitmask of attachment flags (e.g. ATT_INVISIBLE_IN_HTML, ATT_INVISIBLE_IN_RTF).
        /// </summary>
        public const string PrAttachFlags = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
    }
}
