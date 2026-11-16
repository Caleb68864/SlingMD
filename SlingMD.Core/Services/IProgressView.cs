using System;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// The progress window <see cref="StatusService"/> drives. SlingMD.Core has no UI, so the add-in
    /// implements this over its <c>ProgressForm</c> (<c>ProgressFormView</c>); tests pass a recording
    /// double. The view is already shown when it is handed over.
    /// </summary>
    public interface IProgressView : IDisposable
    {
        void UpdateProgress(string message, int percentage);

        void ShowSuccess(string message, bool autoClose);

        void ShowError(string message, bool autoClose);

        void Close();
    }
}
