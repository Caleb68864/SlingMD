using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that strips SlingMD's legacy filename decorations from a name-without-extension:
    /// <list type="bullet">
    ///   <item>Trailing <c>-eid{ID}</c> email-id markers (e.g. <c>"foo-eidABC123"</c> → <c>"foo"</c>).</item>
    ///   <item>Trailing <c>-NNN</c> 3-digit numeric collision suffixes (e.g. <c>"foo-001"</c> → <c>"foo"</c>).</item>
    ///   <item>Leading <c>yyyy-MM-dd[_-]HHmm[ss]_?</c> date prefixes used by older thread folder layouts.</item>
    /// </list>
    /// Order matters — trailing markers are removed first so the leading-date sweep sees a clean tail.
    /// </summary>
    public class LegacyFilenameStripper
    {
        private static readonly Regex EmailIdSuffixRegex = new Regex(@"-eid[0-9A-Za-z]+$", RegexOptions.Compiled);
        private static readonly Regex NumericSuffixRegex = new Regex(@"-\d{3}$", RegexOptions.Compiled);
        private static readonly Regex LegacyDatePrefixRegex = new Regex(@"^\d{4}-\d{2}-\d{2}[_-]\d{4,6}_?", RegexOptions.Compiled);

        /// <summary>
        /// Returns <paramref name="nameNoExt"/> with the three legacy decorations stripped.
        /// </summary>
        public string Strip(string nameNoExt)
        {
            if (string.IsNullOrEmpty(nameNoExt))
            {
                return nameNoExt ?? string.Empty;
            }

            string cleaned = nameNoExt;
            cleaned = EmailIdSuffixRegex.Replace(cleaned, string.Empty);
            cleaned = NumericSuffixRegex.Replace(cleaned, string.Empty);
            cleaned = LegacyDatePrefixRegex.Replace(cleaned, string.Empty);
            return cleaned;
        }
    }
}
