namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// A single find/replace regex rule applied by <see cref="Services.Formatting.FilenameSubjectNormalizer"/>
    /// to canonicalize subject strings for filename use (e.g. collapsing repeated "Re_" prefixes).
    /// Pure DTO — no Outlook Interop dependencies.
    /// </summary>
    public class FilenameSubjectRule
    {
        /// <summary>Gets or sets the regex pattern to search for (case-insensitive).</summary>
        public string Pattern { get; set; } = string.Empty;

        /// <summary>Gets or sets the literal replacement string.</summary>
        public string Replacement { get; set; } = string.Empty;
    }
}
