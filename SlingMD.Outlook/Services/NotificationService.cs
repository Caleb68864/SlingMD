using System.Windows.Forms;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class NotificationService
    {
        private readonly ObsidianSettings _settings;

        public NotificationService(ObsidianSettings settings)
        {
            _settings = settings;
        }

        /// <summary>
        /// Displays an informational notification. In Toast mode, shows a MessageBox.
        /// In Silent mode, writes to the log only.
        /// </summary>
        public virtual void Notify(string message)
        {
            Logger.Instance.Info(message);

            if (_settings.AutoSlingNotificationMode == "Toast")
            {
                try
                {
                    ToastForm.ShowToast(message, isError: false);
                }
                catch (System.Exception ex)
                {
                    // Deliberately broad: a toast is best effort and must never break the caller.
                    // The message itself is already in the log above.
                    Logger.Instance.Warning($"NotificationService.Notify: toast failed: {ex.GetType().Name}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Displays an error notification. In Toast mode, shows a MessageBox with a warning icon.
        /// In Silent mode, writes to the log only.
        /// </summary>
        public virtual void NotifyError(string message, string error)
        {
            string fullMessage = string.IsNullOrEmpty(error) ? message : $"{message}\n\nDetails: {error}";

            Logger.Instance.Error(fullMessage);

            if (_settings.AutoSlingNotificationMode == "Toast")
            {
                try
                {
                    ToastForm.ShowToast(fullMessage, isError: true);
                }
                catch (System.Exception ex)
                {
                    // Deliberately broad: a toast is best effort and must never break the caller.
                    Logger.Instance.Warning($"NotificationService.NotifyError: toast failed: {ex.GetType().Name}: {ex.Message}");
                }
            }
        }
    }
}
