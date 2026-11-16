using SlingMD.Outlook.Forms;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// <see cref="IProgressView"/> over the add-in's <see cref="ProgressForm"/>, so
    /// <see cref="StatusService"/> (SlingMD.Core) never touches WinForms.
    /// </summary>
    public sealed class ProgressFormView : IProgressView
    {
        private readonly ProgressForm _form;

        private ProgressFormView(ProgressForm form)
        {
            _form = form;
        }

        /// <summary>Creates and shows a progress window, as <c>new StatusService()</c> used to.</summary>
        public static IProgressView Open()
        {
            ProgressForm form = new ProgressForm();
            form.Show();
            return new ProgressFormView(form);
        }

        public void UpdateProgress(string message, int percentage) => _form.UpdateProgress(message, percentage);

        public void ShowSuccess(string message, bool autoClose) => _form.ShowSuccess(message, autoClose);

        public void ShowError(string message, bool autoClose) => _form.ShowError(message, autoClose);

        public void Close() => _form.Close();

        public void Dispose() => _form.Dispose();
    }
}
