using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Xunit;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Pins the wait EmailProcessor did after resuffixing thread notes (WaitForFileAvailability):
    /// up to 5 checks, 4 waits of 50, 100, 200 and 400 ms between them, and no exception when the
    /// file never shows up.
    /// </summary>
    public class FileAvailabilityWaiterTests : IDisposable
    {
        private readonly string _folder;

        public FileAvailabilityWaiterTests()
        {
            _folder = Path.Combine(Path.GetTempPath(), "SlingMDTests", "FileAvailabilityWaiter", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_folder);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_folder, true);
            }
            catch (IOException)
            {
                // Best effort cleanup of a temp folder.
            }
        }

        [Fact]
        public async Task WaitAsync_FileExists_ReturnsWithoutWaiting()
        {
            string path = Path.Combine(_folder, "note.md");
            File.WriteAllText(path, "x");
            List<int> delays = new List<int>();

            await FileAvailabilityWaiter.WaitAsync(path, ms => { delays.Add(ms); return Task.CompletedTask; });

            Assert.Empty(delays);
        }

        [Fact]
        public async Task WaitAsync_FileNeverAppears_WaitsFourTimesDoublingThenReturns()
        {
            string path = Path.Combine(_folder, "missing.md");
            List<int> delays = new List<int>();

            await FileAvailabilityWaiter.WaitAsync(path, ms => { delays.Add(ms); return Task.CompletedTask; });

            Assert.Equal(new[] { 50, 100, 200, 400 }, delays);
        }

        [Fact]
        public async Task WaitAsync_FileAppearsDuringWait_StopsAtNextCheck()
        {
            string path = Path.Combine(_folder, "late.md");
            List<int> delays = new List<int>();

            await FileAvailabilityWaiter.WaitAsync(path, ms =>
            {
                delays.Add(ms);
                if (delays.Count == 2)
                {
                    File.WriteAllText(path, "x");
                }
                return Task.CompletedTask;
            });

            Assert.Equal(new[] { 50, 100 }, delays);
        }

        [Fact]
        public async Task WaitAsync_NullDelay_UsesRealDelayAndReturnsForMissingFile()
        {
            string path = Path.Combine(_folder, "missing.md");

            // 50 + 100 + 200 + 400 ms of real waiting; it must finish and not throw.
            await FileAvailabilityWaiter.WaitAsync(path);

            Assert.False(File.Exists(path));
        }

        [Fact]
        public void BackoffConstants_AreUnchanged()
        {
            Assert.Equal(5, FileAvailabilityWaiter.MaxFileWaitAttempts);
            Assert.Equal(50, FileAvailabilityWaiter.InitialFileWaitDelayMs);
        }
    }
}
