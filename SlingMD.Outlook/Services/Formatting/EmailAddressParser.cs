namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper for splitting an email address into its local-part (before the '@') and
    /// domain (after the '@'). Handles null, empty, missing-'@', and trailing-'@' edge
    /// cases by returning empty strings rather than throwing.
    /// </summary>
    public class EmailAddressParser
    {
        /// <summary>
        /// Returns the substring after the first '@' in <paramref name="email"/>.
        /// Returns empty if the input is null/empty, has no '@', or ends with '@'.
        /// </summary>
        public string Domain(string email)
        {
            if (string.IsNullOrEmpty(email))
            {
                return string.Empty;
            }

            int atIndex = email.IndexOf('@');
            if (atIndex < 0 || atIndex == email.Length - 1)
            {
                return string.Empty;
            }

            return email.Substring(atIndex + 1);
        }

        /// <summary>
        /// Returns the substring before the first '@' in <paramref name="email"/>.
        /// Returns empty if the input is null or whitespace. If no '@' is present, the
        /// whole input (trimmed) is returned — callers commonly use this as a display-name
        /// fallback when an email is present but no display name.
        /// </summary>
        public string LocalPart(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return string.Empty;
            }

            int atIndex = email.IndexOf('@');
            if (atIndex > 0)
            {
                return email.Substring(0, atIndex);
            }

            return email;
        }
    }
}
