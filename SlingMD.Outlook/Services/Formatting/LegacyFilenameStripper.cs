using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that strips SlingMD's legacy filename decorations from a name-without-extension:
    /// <list type="bullet">
    ///   <item>Trailing <c>-eid{ID}</c> email-id markers (e.g. <c>"foo-eidABC123"</c> → <c>"foo"</c>).</item>
    ///   <item>Trailing <c>-NNN</c> 3-digit numeric collision suffixes (e.g. <c>"foo-001"</c> → <c>"foo"</c>).</item>
    ///   <item>Trailing <c>-yyyy-MM-dd-HHmm</c> received-time stamps carried by notes first written outside a thread.</item>
    ///   <item>Leading <c>yyyy-MM-dd[_-]HHmm[ss]</c> date prefixes, plus the separators that follow them.</item>
    /// </list>
    /// Order matters — trailing markers are removed first so the leading-date sweep sees a clean tail.
    /// </summary>
    internal class LegacyFilenameStripper
    {
        private static readonly Regex EmailIdSuffixRegex = new Regex(@"-eid[0-9A-Za-z]+$", RegexOptions.Compiled);
        private static readonly Regex NumericSuffixRegex = new Regex(@"-\d{3}$", RegexOptions.Compiled);
        private static readonly Regex TrailingDateRegex = new Regex(@"-\d{4}-\d{2}-\d{2}-\d{4}$", RegexOptions.Compiled);
        // Consumes every separator after the prefix: a threaded base name is "{date}-{subject}", so
        // leaving the dash behind produced "2026-09-23_181852_-Subject" once the prefix was re-added.
        private static readonly Regex LegacyDatePrefixRegex = new Regex(@"^\d{4}-\d{2}-\d{2}[_-]\d{4,6}[-_\s]*", RegexOptions.Compiled);

        /// <summary>
        /// Returns <paramref name="nameNoExt"/> with the legacy decorations stripped.
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
            string withoutTrailingDate = TrailingDateRegex.Replace(cleaned, string.Empty);
            cleaned = LegacyDatePrefixRegex.Replace(withoutTrailingDate, string.Empty);

            // A name that is nothing but decorations keeps its trailing date rather than going empty.
            if (cleaned.Length == 0)
            {
                cleaned = LegacyDatePrefixRegex.Replace(nameNoExt, string.Empty);
            }
            return cleaned;
        }
    }
}
