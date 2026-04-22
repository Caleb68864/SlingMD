using System;
using System.Security.Cryptography;
using System.Text;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Generates stable, 20-character hexadecimal thread identifiers from conversation topics.
    /// Uses MD5 hashing of the normalized subject to match the historical hash shape used by
    /// ThreadService, ensuring existing thread IDs remain unchanged.
    /// </summary>
    public class ThreadIdHasher
    {
        private readonly SubjectCleanerService _cleaner;

        /// <summary>
        /// Initializes a new instance of the <see cref="ThreadIdHasher"/> class.
        /// </summary>
        /// <param name="cleaner">The subject cleaner service used for prefix normalization.</param>
        public ThreadIdHasher(SubjectCleanerService cleaner)
        {
            _cleaner = cleaner ?? throw new ArgumentNullException(nameof(cleaner));
        }

        /// <summary>
        /// Generates a 20-character hexadecimal hash from the conversation topic.
        /// The topic is first normalized using <see cref="SubjectCleanerService.NormalizeForGrouping"/>
        /// to strip Re:/Fwd: prefixes, ensuring prefix-insensitive thread grouping.
        /// </summary>
        /// <param name="conversationTopic">The conversation topic to hash.</param>
        /// <returns>A 20-character uppercase hexadecimal string suitable for use as a thread identifier.</returns>
        public string Hash(string conversationTopic)
        {
            if (string.IsNullOrEmpty(conversationTopic))
            {
                return string.Empty;
            }

            string normalizedSubject = _cleaner.NormalizeForGrouping(conversationTopic);

            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(normalizedSubject);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                // Convert to uppercase hex string and take first 20 characters
                // This matches the existing ThreadService.GetConversationId behavior
                return BitConverter.ToString(hashBytes).Replace("-", string.Empty).Substring(0, 20);
            }
        }
    }
}
