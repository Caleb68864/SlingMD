namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// The modal prompts <c>EmailProcessor</c> shows on the interactive (non-bulk) path. SlingMD.Core
    /// has no UI, so the add-in implements this over its WinForms dialogs (<c>DialogEmailPrompts</c>);
    /// tests pass a recording double. Methods block until the user dismisses the prompt.
    /// </summary>
    public interface IEmailPrompts
    {
        /// <summary>Shows a "SlingMD Error" message with an OK button.</summary>
        void ShowError(string message);

        /// <summary>Shows the "opening Obsidian in N seconds" countdown and returns when it ends or is dismissed.</summary>
        void ShowLaunchCountdown(int seconds);
    }
}
