using System;
using System.IO;
using System.Threading.Tasks;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Waits for a file to exist and be openable, with exponential backoff. Moved unchanged from
    /// <c>EmailProcessor.WaitForFileAvailability</c>, which uses it after the thread notes are
    /// resuffixed. It never throws for a missing or locked file: after the last attempt it returns
    /// and the caller carries on, as before.
    /// </summary>
    internal static class FileAvailabilityWaiter
    {
        // Exponential-backoff parameters.
        // 5 attempts × starting at 50 ms doubles to ~1550 ms total wait worst-case.
        internal const int MaxFileWaitAttempts = 5;
        internal const int InitialFileWaitDelayMs = 50;

        /// <summary>
        /// Waits for <paramref name="filePath"/> to become available by checking its existence and
        /// attempting to open it. <paramref name="delay"/> replaces <see cref="Task.Delay(int)"/> in
        /// tests; <c>null</c> means <see cref="Task.Delay(int)"/>.
        /// </summary>
        internal static async Task WaitAsync(string filePath, Func<int, Task> delay = null)
        {
            Func<int, Task> wait = delay ?? (ms => Task.Delay(ms));
            int attempt = 0;
            int delayMs = InitialFileWaitDelayMs;

            while (attempt < MaxFileWaitAttempts)
            {
                try
                {
                    // Check if file exists and can be opened
                    if (File.Exists(filePath))
                    {
                        // Try to open the file briefly to ensure it's not locked
                        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            // File is accessible
                            return;
                        }
                    }
                }
                catch (IOException)
                {
                    // File is locked or not yet available
                }
                catch (UnauthorizedAccessException)
                {
                    // Permission issue, but file exists
                    return;
                }

                // Exponential backoff: InitialFileWaitDelayMs doubled on each retry.
                attempt++;
                if (attempt < MaxFileWaitAttempts)
                {
                    await wait(delayMs);
                    delayMs *= 2;
                }
            }

            // If we get here, file still isn't available but we've tried our best
            // Continue anyway - the worst case is the same as before
        }
    }
}
