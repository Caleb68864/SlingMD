using System;
using System.Threading.Tasks;
using SlingMD.Outlook.Forms;

namespace SlingMD.Outlook.Services
{
    public class StatusService : IDisposable
    {
        private ProgressForm _progressForm;
        private bool _isDisposed;
        private readonly bool _silent;

        public StatusService() : this(false)
        {
        }

        /// <summary>
        /// Creates the status service. When <paramref name="silent"/> is true no progress window is
        /// shown and all updates are no-ops — used for batch/bulk processing where one visible
        /// progress window per item would be a focus-stealing window storm (batch sling).
        /// </summary>
        public StatusService(bool silent)
        {
            _silent = silent;
            if (!_silent)
            {
                _progressForm = new ProgressForm();
                _progressForm.Show();
            }
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