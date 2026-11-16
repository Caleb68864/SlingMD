using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Logs a message and, when <see cref="ObsidianSettings.AutoSlingNotificationMode"/> is "Toast",
    /// shows it through an <see cref="IToastNotifier"/> (the add-in passes <c>ToastFormNotifier</c>).
    /// In any other mode it only logs.
    /// </summary>
    public class NotificationService
    {
        private readonly ObsidianSettings _settings;
        private readonly IToastNotifier _toasts;

        public NotificationService(ObsidianSettings settings, IToastNotifier toastNotifier)
        {
            _settings = settings;
            _toasts = toastNotifier ?? throw new System.ArgumentNullException(nameof(toastNotifier));
        }

        /// <summary>
        /// Displays an informational notification. In Toast mode, shows a toast.
        /// In Silent mode, writes to the log only.
        /// </summary>
        public virtual void Notify(string message)
        {
            Logger.Instance.Info(message);

            if (_settings.AutoSlingNotificationMode == "Toast")
            {
                try
                {
                    _toasts.ShowToast(message, isError: false);
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
        /// Displays an error notification. In Toast mode, shows an error toast.
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
                    _toasts.ShowToast(fullMessage, isError: true);
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
