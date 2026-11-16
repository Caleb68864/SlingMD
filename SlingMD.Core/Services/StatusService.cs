using System;
using System.Threading.Tasks;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Reports progress of one sling to a progress window (<see cref="IProgressView"/>), or to
    /// nothing in silent mode. Moved from the add-in; only the window became an interface.
    /// </summary>
    public class StatusService : IDisposable
    {
        private IProgressView _progressForm;
        private bool _isDisposed;

        /// <summary>
        /// Creates the status service over an already shown progress window. When
        /// <paramref name="progressView"/> is null no progress window is shown and all updates are
        /// no-ops, used for batch/bulk processing where one visible progress window per item would be
        /// a focus-stealing window storm (batch sling). The add-in passes
        /// <c>ProgressFormView.Open()</c> where it used to call <c>new StatusService()</c>.
        /// </summary>
        public StatusService(IProgressView progressView)
        {
            _progressForm = progressView;
        }

        public void UpdateProgress(string message, int percentage)
        {
            EnsureNotDisposed();
            _progressForm?.UpdateProgress(message, percentage);
        }

        public void ShowSuccess(string message, bool autoClose = true)
        {
            EnsureNotDisposed();
            _progressForm?.ShowSuccess(message, autoClose);
        }

        public void ShowError(string message, bool autoClose = false)
        {
            EnsureNotDisposed();
            _progressForm?.ShowError(message, autoClose);
        }

        public async Task ShowTemporaryStatusAsync(string message, int durationMs = 3000)
        {
            EnsureNotDisposed();
            UpdateProgress(message, 100);
            await Task.Delay(durationMs);
            _progressForm?.Close();
        }

        private void EnsureNotDisposed()
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException(nameof(StatusService));
            }
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                _progressForm?.Dispose();
                _progressForm = null;
                _isDisposed = true;
            }
        }
    }
} 