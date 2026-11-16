using System;
using System.IO;

namespace SlingMD.Outlook.Helpers
{
    /// <summary>
    /// Simple file-based logger for SlingMD. Writes to AppData/SlingMD.Outlook/logs/.
    /// Thread-safe singleton implementation.
    /// </summary>
    public sealed class Logger
    {
        private static readonly Lazy<Logger> _instance = new Lazy<Logger>(() => new Logger());
        private readonly string _logFilePath;
        private readonly object _lockObject = new object();
        private const int MAX_LOG_FILE_SIZE_MB = 10;

        private Logger()
        {
            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "SlingMD.Outlook",
                "logs"
            );

            Directory.CreateDirectory(logDir);

            string logFileName = $"SlingMD_{DateTime.Now:yyyyMMdd}.log";
            _logFilePath = Path.Combine(logDir, logFileName);
        }

        public static Logger Instance => _instance.Value;

        /// <summary>
        /// Logs an informational message.
        /// </summary>
        public void Info(string message)
        {
            Log("INFO", message);
        }

        /// <summary>
        /// Logs a warning message.
        /// </summary>
        public void Warning(string message)
        {
            Log("WARN", message);
        }

        /// <summary>
        /// Logs an error message.
        /// </summary>
        public void Error(string message, System.Exception ex = null)
        {
            string fullMessage = ex != null
                ? $"{message}\nException: {ex.GetType().Name}: {ex.Message}\nStack Trace: {ex.StackTrace}"
                : message;

            Log("ERROR", fullMessage);
        }

        /// <summary>
        /// Logs a debug message (only in DEBUG builds).
        /// </summary>
        public void Debug(string message)
        {
#if DEBUG
            Log("DEBUG", message);
#endif
        }

        private void Log(string level, string message)
        {
            try
            {
                lock (_lockObject)
                {
                    // Check file size and rotate if needed
                    if (File.Exists(_logFilePath))
                    {
                        FileInfo fileInfo = new FileInfo(_logFilePath);
                        if (fileInfo.Length > MAX_LOG_FILE_SIZE_MB * 1024 * 1024)
                        {
                            // Archive old log
                            string archivePath = _logFilePath.Replace(".log", $"_{DateTime.Now:HHmmss}.log");
                            File.Move(_logFilePath, archivePath);
                        }
                    }

                    string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}";

                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                }
            }
            catch
            {
                // Silently fail - don't crash the add-in if logging fails
            }
        }

        /// <summary>
        /// Cleans up old log files (keeps last 30 days).
        /// </summary>
        public void CleanOldLogs()
        {
            try
            {
                string logDir = Path.GetDirectoryName(_logFilePath);
                string[] logFiles = Directory.GetFiles(logDir, "SlingMD_*.log");

                foreach (string logFile in logFiles)
                {
                    FileInfo fileInfo = new FileInfo(logFile);
                    if (fileInfo.LastWriteTime < DateTime.Now.AddDays(-30))
                    {
                        File.Delete(logFile);
                    }
                }
            }
            catch
            {
                // Silently fail
            }
        }
    }
}
