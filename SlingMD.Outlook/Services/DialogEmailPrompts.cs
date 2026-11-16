using System.Windows.Forms;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// <see cref="IEmailPrompts"/> over the add-in's WinForms dialogs: a <see cref="MessageBox"/> for
    /// errors, the <see cref="CountdownForm"/> before Obsidian launches and the
    /// <see cref="TaskOptionsForm"/> task-dates prompt.
    /// </summary>
    public sealed class DialogEmailPrompts : IEmailPrompts
    {
        public void ShowError(string message)
        {
            MessageBox.Show(message, "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void ShowLaunchCountdown(int seconds)
        {
            using (var countdown = new CountdownForm(seconds))
            {
                countdown.ShowDialog();
            }
        }

        public TaskDateChoice PromptTaskOptions(int defaultDueDays, int defaultReminderDays, int defaultReminderHour, bool useRelativeReminder)
        {
            using (var form = new TaskOptionsForm(defaultDueDays, defaultReminderDays, defaultReminderHour, useRelativeReminder))
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    return null;
                }

                return new TaskDateChoice(form.DueDays, form.ReminderDays, form.ReminderHour, form.UseRelativeReminder);
            }
        }
    }
}
