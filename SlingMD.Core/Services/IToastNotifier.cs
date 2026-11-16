namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// A short non-modal message the engine asks the front end to show (the add-in's
    /// <c>ToastFormNotifier</c> uses <c>ToastForm</c>). Showing it is best effort: an implementation
    /// must not throw because a toast could not be shown.
    /// </summary>
    public interface IToastNotifier
    {
        void ShowToast(string message, bool isError);
    }
}
