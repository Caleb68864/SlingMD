using SlingMD.Outlook.Forms;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// <see cref="IToastNotifier"/> over <see cref="ToastForm.ShowToast"/>, which already logs and
    /// drops a toast it cannot show rather than throwing.
    /// </summary>
    public sealed class ToastFormNotifier : IToastNotifier
    {
        public void ShowToast(string message, bool isError) => ToastForm.ShowToast(message, isError);
    }
}
