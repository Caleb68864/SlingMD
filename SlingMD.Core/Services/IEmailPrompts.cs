using SlingMD.Outlook.Models;

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

        /// <summary>
        /// Asks for the task due/reminder dates, starting from the given defaults. Returns the choice,
        /// or <c>null</c> when the user cancels (the caller then creates no task).
        /// </summary>
        TaskDateChoice PromptTaskOptions(int defaultDueDays, int defaultReminderDays, int defaultReminderHour, bool useRelativeReminder);
    }
}
