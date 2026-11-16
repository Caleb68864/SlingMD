namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// What the engine may ask the front end to show the user. SlingMD.Core has no UI, so the
    /// Outlook add-in implements this with message boxes (<c>MessageBoxUserNotifier</c>); tests pass
    /// a recording double or nothing.
    /// </summary>
    public interface IUserNotifier
    {
        /// <summary>Shows a warning the user has to dismiss (the add-in uses a MessageBox).</summary>
        void ShowWarning(string message, string title);
    }
}
