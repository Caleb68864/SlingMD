namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Result of the auto-sling decision evaluation.
    /// Pure DTO with no Outlook Interop dependencies.
    /// </summary>
    public class SlingDecision
    {
        /// <summary>
        /// Gets or sets whether the item should be slung.
        /// </summary>
        public bool ShouldSling { get; set; }

        /// <summary>
        /// Gets or sets the rule that matched (null if no match).
        /// </summary>
        public AutoSlingRule MatchedRule { get; set; }
    }
}
