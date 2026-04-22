using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Normalizes a subject line into a filename-safe stem by replacing colons with
    /// underscores and collapsing repeated Re_/Fw_ prefixes (in either case) into a
    /// single canonical "Re_"/"Fw_". Pure helper — no Outlook or filesystem deps.
    /// </summary>
    public class FilenameSubjectNormalizer
    {
        private static readonly Regex ColonSpaceRegex = new Regex(@":\s*", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex ReplyPrefixRegex1 = new Regex(@"(?:Re_\s*)+(?:RE_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex2 = new Regex(@"(?:RE_\s*)+(?:Re_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex3 = new Regex(@"(?:Re_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReplyPrefixRegex4 = new Regex(@"(?:RE_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex ForwardPrefixRegex1 = new Regex(@"(?:Fw_\s*)+(?:FW_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex2 = new Regex(@"(?:FW_\s*)+(?:Fw_\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex3 = new Regex(@"(?:Fw_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardPrefixRegex4 = new Regex(@"(?:FW_\s*){2,}", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex ReplySpaceRegex = new Regex(@"Re_\s+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ForwardSpaceRegex = new Regex(@"Fw_\s+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Normalizes <paramref name="subject"/> for filename use:
        /// 1. Convert "colon + optional space" to a single underscore.
        /// 2. Collapse repeated Re_/RE_ runs into a single "Re_".
        /// 3. Collapse repeated Fw_/FW_ runs into a single "Fw_".
        /// 4. Drop the trailing space after a "Re_" or "Fw_" stub.
        /// </summary>
        public string Normalize(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            string cleaned = ColonSpaceRegex.Replace(subject, "_");

            cleaned = ReplyPrefixRegex1.Replace(cleaned, "Re_");
            cleaned = ReplyPrefixRegex2.Replace(cleaned, "Re_");
            cleaned = ReplyPrefixRegex3.Replace(cleaned, "Re_");
            cleaned = ReplyPrefixRegex4.Replace(cleaned, "Re_");

            cleaned = ForwardPrefixRegex1.Replace(cleaned, "Fw_");
            cleaned = ForwardPrefixRegex2.Replace(cleaned, "Fw_");
            cleaned = ForwardPrefixRegex3.Replace(cleaned, "Fw_");
            cleaned = ForwardPrefixRegex4.Replace(cleaned, "Fw_");

            cleaned = ReplySpaceRegex.Replace(cleaned, "Re_");
            cleaned = ForwardSpaceRegex.Replace(cleaned, "Fw_");

            return cleaned;
        }
    }
}
