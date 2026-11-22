using System.Windows.Forms;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// The add-in's <see cref="IUserNotifier"/>: shows each message in a WinForms message box.
    /// </summary>
    public class MessageBoxUserNotifier : IUserNotifier
    {
        public void ShowWarning(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public void ShowInformation(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
