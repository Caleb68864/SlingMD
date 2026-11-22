using System.Collections.Generic;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// The add-in's <see cref="IContactMatchPrompt"/>: shows <see cref="ContactMatchPromptForm"/>.
    /// Moved here from <c>ContactService.TryHandleFuzzyMatch</c> unchanged.
    /// </summary>
    internal sealed class DialogContactMatchPrompt : IContactMatchPrompt
    {
        public bool ChooseExisting(IReadOnlyList<ContactIndexEntry> candidates, ContactInteractionMode mode, bool includeDetails)
        {
            // The form's constructor throws InvalidOperationException in Automated mode; ContactService catches it.
            using (ContactMatchPromptForm form = new ContactMatchPromptForm(candidates, mode, includeDetails))
            {
                return form.ShowDialog() == System.Windows.Forms.DialogResult.OK
                    && form.Result != null
                    && form.Result.Decision == MatchPromptDecision.Match
                    && form.Result.ChosenEntry != null;
            }
        }
    }
}
