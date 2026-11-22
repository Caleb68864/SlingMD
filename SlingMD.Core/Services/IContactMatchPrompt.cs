using System.Collections.Generic;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Asks the user whether an ambiguous contact name is one of the existing contact notes. SlingMD.Core
    /// has no UI; the add-in implements it with <c>ContactMatchPromptForm</c> (<c>DialogContactMatchPrompt</c>).
    /// Internal because the candidates are <see cref="ContactIndexEntry"/>s.
    /// </summary>
    internal interface IContactMatchPrompt
    {
        /// <summary>
        /// Shows <paramref name="candidates"/>; true when the user picked an existing contact, false
        /// when they chose to create a new one (or closed the prompt). May throw
        /// <see cref="System.InvalidOperationException"/> when it cannot prompt in <paramref name="mode"/>.
        /// </summary>
        bool ChooseExisting(IReadOnlyList<ContactIndexEntry> candidates, ContactInteractionMode mode, bool includeDetails);
    }
}
