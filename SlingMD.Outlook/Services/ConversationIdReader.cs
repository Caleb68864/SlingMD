using System;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Reads an Outlook <see cref="MailItem"/>'s conversation id: the COM half of the thread code.
    /// The file-system half is <see cref="ThreadService"/> (SlingMD.Core). Moved here from
    /// <c>ThreadService</c> unchanged.
    /// </summary>
    public class ConversationIdReader
    {
        private readonly ThreadIdHasher _threadIdHasher;

        public ConversationIdReader(ObsidianSettings settings)
        {
            // Use a fallback settings instance when null so "store-only" constructor pattern
            // is preserved (exceptions only surface when methods are actually called).
            _threadIdHasher = new ThreadIdHasher(new SubjectCleanerService(settings ?? new ObsidianSettings()));
        }

        /// <summary>
        /// Derives a stable 20-character hash that uniquely identifies an Outlook conversation.
        /// The method tries several strategies (ConversationTopic, PR_CONVERSATION_INDEX, normalised subject)
        /// and falls back to a random GUID segment if everything else fails. Using 20 characters instead of 16
        /// significantly reduces the probability of hash collisions while remaining compact.
        /// </summary>
        /// <param name="mail">The <see cref="MailItem"/> for which to obtain the conversation id.</param>
        /// <returns>A 20-character hexadecimal string suitable for use as an identifier inside vault front-matter.</returns>
        public string GetConversationId(MailItem mail)
        {
            try
            {
                // Strategy 1: Hash the conversation topic (most reliable for threading).
                if (!string.IsNullOrEmpty(mail.ConversationTopic))
                {
                    return _threadIdHasher.Hash(mail.ConversationTopic);
                }

                // Strategy 2: PR_CONVERSATION_INDEX — Exchange-native thread bytes. Different hash shape
                // than ThreadIdHasher; preserved as a separate strategy because its ID is not derived
                // from the subject line.
                // A failure here (property missing, store offline, item type without it) must fall
                // through to the subject hash, not to the random id below, which would put every such
                // email in a thread folder of its own.
                byte[] conversationIndex = null;
                try
                {
                    conversationIndex = ReadConversationIndex(mail);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Helpers.Logger.Instance.Warning($"ThreadService.GetConversationId: PR_CONVERSATION_INDEX unavailable, using the subject: {ex.Message}");
                }

                if (conversationIndex != null && conversationIndex.Length >= 22)
                {
                    return BitConverter.ToString(conversationIndex.Take(22).ToArray())
                        .Replace("-", "").Substring(0, 20);
                }

                // Strategy 3: Fall back to hashing the subject line.
                return _threadIdHasher.Hash(mail.Subject ?? string.Empty);
            }
            catch (System.Exception ex)
            {
                // Last resort, kept deliberately broad: a sling must not fail over a thread id. The id is
                // then taken from the email's own identity, so slinging it again gives the same id and
                // finds the note written the first time. Logged because the email will not be grouped
                // with the rest of its conversation.
                string fallbackId = FallbackThreadId.FromIdentity(
                    () => ReadInternetMessageId(mail),
                    () => mail.EntryID,
                    message => Helpers.Logger.Instance.Warning($"ThreadService.GetConversationId: {message}"));
                if (fallbackId != null)
                {
                    Helpers.Logger.Instance.Warning($"ThreadService.GetConversationId: no conversation id ({ex.GetType().Name}: {ex.Message}); using a hash of the email's message id or EntryID");
                    return fallbackId;
                }

                // Nothing about the item can be read at all, so there is nothing stable to hash.
                Helpers.Logger.Instance.Warning($"ThreadService.GetConversationId: falling back to a random thread id: {ex.GetType().Name}: {ex.Message}");
                return Guid.NewGuid().ToString("N").Substring(0, FallbackThreadId.Length).ToUpperInvariant();
            }
        }

        /// <summary>
        /// Reads PR_CONVERSATION_INDEX (PT_BINARY) through the item's PropertyAccessor, releasing the
        /// accessor. Returns null when the property is not a byte array. Throws whatever the COM calls
        /// throw; the caller falls back to the subject on a <see cref="System.Runtime.InteropServices.COMException"/>.
        /// </summary>
        private static byte[] ReadConversationIndex(MailItem mail)
        {
            PropertyAccessor accessor = null;
            try
            {
                accessor = mail.PropertyAccessor;
                return accessor.GetProperty(MapiPropertyTags.PrConversationIndex) as byte[];
            }
            finally
            {
                ReleaseAccessor(accessor);
            }
        }

        /// <summary>
        /// Releases a PropertyAccessor obtained from an Outlook item.
        /// </summary>
        private static void ReleaseAccessor(PropertyAccessor accessor)
        {
            // IsComObject: only a real COM accessor has a reference to release (a test double is a
            // plain object, and ReleaseComObject would throw ArgumentException on it).
            if (accessor != null && System.Runtime.InteropServices.Marshal.IsComObject(accessor))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(accessor);
            }
        }

        /// <summary>
        /// Reads PR_INTERNET_MESSAGE_ID through the item's PropertyAccessor, releasing the accessor.
        /// Throws whatever the COM calls throw; <see cref="FallbackThreadId"/> catches it.
        /// </summary>
        private static string ReadInternetMessageId(MailItem mail)
        {
            PropertyAccessor accessor = null;
            try
            {
                accessor = mail.PropertyAccessor;
                return accessor.GetProperty(MapiPropertyTags.PrInternetMessageId) as string;
            }
            finally
            {
                ReleaseAccessor(accessor);
            }
        }

    }
}
