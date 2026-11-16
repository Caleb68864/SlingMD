using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Pins what StatusService did with its ProgressForm before it moved to SlingMD.Core behind
    /// <see cref="IProgressView"/>: calls pass straight through with the same defaults, silent mode
    /// does nothing, the temporary status shows 100% then closes, and use after Dispose throws.
    /// </summary>
    public class StatusServiceTests
    {
        private sealed class RecordingProgressView : IProgressView
        {
            public readonly List<string> Calls = new List<string>();

            public void UpdateProgress(string message, int percentage) => Calls.Add($"Update:{message}:{percentage}");

            public void ShowSuccess(string message, bool autoClose) => Calls.Add($"Success:{message}:{autoClose}");

            public void ShowError(string message, bool autoClose) => Calls.Add($"Error:{message}:{autoClose}");

            public void Close() => Calls.Add("Close");

            public void Dispose() => Calls.Add("Dispose");
        }

        [Fact]
        public void Calls_PassThroughWithOriginalDefaults()
        {
            RecordingProgressView view = new RecordingProgressView();
            using (StatusService status = new StatusService(view))
            {
                status.UpdateProgress("Processing email...", 0);
                status.ShowSuccess("done");
                status.ShowError("bad");
                status.ShowSuccess("kept", autoClose: false);
                status.ShowError("closing", autoClose: true);
            }

            Assert.Equal(new[]
            {
                "Update:Processing email...:0",
                "Success:done:True",
                "Error:bad:False",
                "Success:kept:False",
                "Error:closing:True",
                "Dispose",
            }, view.Calls);
        }

        [Fact]
        public async Task Silent_NullView_EverythingIsANoOp()
        {
            using (StatusService status = new StatusService(null))
            {
                status.UpdateProgress("x", 10);
                status.ShowSuccess("x");
                status.ShowError("x");
                await status.ShowTemporaryStatusAsync("x", 0);
            }
        }

        [Fact]
        public async Task ShowTemporaryStatusAsync_ShowsFullProgressThenCloses()
        {
            RecordingProgressView view = new RecordingProgressView();
            StatusService status = new StatusService(view);

            await status.ShowTemporaryStatusAsync("Saved", 0);

            Assert.Equal(new[] { "Update:Saved:100", "Close" }, view.Calls);
        }

        [Fact]
        public void Dispose_Twice_DisposesViewOnce()
        {
            RecordingProgressView view = new RecordingProgressView();
            StatusService status = new StatusService(view);

            status.Dispose();
            status.Dispose();

            Assert.Equal(new[] { "Dispose" }, view.Calls);
        }

        [Fact]
        public void AfterDispose_CallsThrowObjectDisposed()
        {
            StatusService status = new StatusService(new RecordingProgressView());
            status.Dispose();

            Assert.Throws<ObjectDisposedException>(() => status.UpdateProgress("x", 1));
            Assert.Throws<ObjectDisposedException>(() => status.ShowSuccess("x"));
            Assert.Throws<ObjectDisposedException>(() => status.ShowError("x"));
            Assert.ThrowsAsync<ObjectDisposedException>(() => status.ShowTemporaryStatusAsync("x", 0)).GetAwaiter().GetResult();
        }
    }
}
