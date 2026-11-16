using System.Windows.Forms;
using SlingMD.Outlook.Forms;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// <see cref="IEmailPrompts"/> over the add-in's WinForms dialogs: a <see cref="MessageBox"/> for
    /// errors and the <see cref="CountdownForm"/> before Obsidian launches.
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
    }
}
