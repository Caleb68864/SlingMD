namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// What happened to one email passed to <see cref="Services.EmailProcessor.ProcessEmailWithResult"/>.
    /// </summary>
    public enum EmailSlingOutcome
    {
        /// <summary>The note was written.</summary>
        Saved,

        /// <summary>The email is already in the vault; nothing was written.</summary>
        Duplicate,

        /// <summary>The export failed; see <see cref="EmailSlingResult.ErrorMessage"/>.</summary>
        Failed
    }

    /// <summary>
    /// Result of slinging one email. Pure DTO with no Outlook Interop dependencies.
    /// </summary>
    public class EmailSlingResult
    {
        public EmailSlingOutcome Outcome { get; set; }

        /// <summary>Full path of the saved note when <see cref="Outcome"/> is <see cref="EmailSlingOutcome.Saved"/>.</summary>
        public string NotePath { get; set; }

        /// <summary>Why the export failed when <see cref="Outcome"/> is <see cref="EmailSlingOutcome.Failed"/>.</summary>
        public string ErrorMessage { get; set; }
    }
}
