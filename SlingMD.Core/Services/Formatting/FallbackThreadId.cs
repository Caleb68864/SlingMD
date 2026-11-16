using System;
using System.Security.Cryptography;
using System.Text;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// The thread id used when none of the conversation-based strategies (topic, PR_CONVERSATION_INDEX,
    /// subject) could be read. It is derived from the email's own identity, Internet message id first
    /// (it survives moves between stores) and EntryID second, so slinging the same email again gives
    /// the same id and finds the note written the first time, instead of a new random id that put
    /// every retry in a thread folder of its own.
    /// Pure helper with no Outlook dependencies: the caller passes the reads as delegates.
    /// </summary>
    internal static class FallbackThreadId
    {
        /// <summary>Length of every thread id, matching <c>ThreadIdHasher</c>.</summary>
        public const int Length = 20;

        /// <summary>
        /// Returns a 20-character uppercase hex id hashed from the first identity value that can be
        /// read and is not blank, or <c>null</c> when neither can be read.
        /// </summary>
        /// <param name="readInternetMessageId">Reads the RFC 2822 Message-ID. May throw.</param>
        /// <param name="readEntryId">Reads the Outlook EntryID. May throw.</param>
        /// <param name="logFailure">Optional; receives a message for each read that threw.</param>
        public static string FromIdentity(Func<string> readInternetMessageId, Func<string> readEntryId, Action<string> logFailure = null)
        {
            string value = TryRead(readInternetMessageId, "Internet message id", logFailure);
            if (string.IsNullOrWhiteSpace(value))
            {
                value = TryRead(readEntryId, "EntryID", logFailure);
            }

            return string.IsNullOrWhiteSpace(value) ? null : Hash(value.Trim());
        }

        /// <summary>Returns the first 20 hex characters (uppercase) of the SHA-1 of <paramref name="value"/>.</summary>
        public static string Hash(string value)
        {
            using (SHA1 sha1 = SHA1.Create())
            {
                byte[] hashBytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
                return BitConverter.ToString(hashBytes).Replace("-", string.Empty).Substring(0, Length);
            }
        }

        private static string TryRead(Func<string> read, string what, Action<string> logFailure)
        {
            if (read == null)
            {
                return null;
            }

            try
            {
                return read();
            }
            catch (System.Exception ex)
            {
                // Deliberately broad: this runs only after the item already failed to give a thread id,
                // typically because the COM object is gone (COMException, InvalidComObjectException,
                // InvalidCastException...). A failed read means "try the next value", never a failed sling.
                logFailure?.Invoke($"could not read the {what}: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }
    }
}
