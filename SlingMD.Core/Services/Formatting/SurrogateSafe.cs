namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Cuts strings without splitting a UTF-16 surrogate pair (emoji and other characters outside
    /// the Basic Multilingual Plane). Half a pair in a file name is written to the note's UTF-8 text
    /// as U+FFFD, so the note's links and title no longer match the file.
    /// </summary>
    internal static class SurrogateSafe
    {
        /// <summary>
        /// The first <paramref name="length"/> UTF-16 units of <paramref name="value"/>, one fewer when
        /// the cut would fall between the two halves of a pair. The whole string when it is shorter.
        /// </summary>
        public static string Truncate(string value, int length)
        {
            if (value == null || length >= value.Length)
            {
                return value;
            }

            if (length <= 0)
            {
                return string.Empty;
            }

            if (char.IsHighSurrogate(value[length - 1]) && char.IsLowSurrogate(value[length]))
            {
                length--;
            }

            return value.Substring(0, length);
        }
    }
}
