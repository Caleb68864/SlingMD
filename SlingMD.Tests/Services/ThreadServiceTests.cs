using System;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Unit tests for <see cref="ThreadService"/> focusing on thread date parsing compatibility
    /// and the handling of a missing inbox folder.
    /// </summary>
    public class ThreadServiceTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;
        private readonly ThreadService _threadService;

        public ThreadServiceTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ThreadService_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                InboxFolder = "Inbox",
                ShowThreadDebug = false
            };
            _fileService = new FileService(_settings);
            _templateService = new TemplateService(_fileService);
            _threadService = new ThreadService(_fileService, _templateService, _settings);
        }

        // -----------------------------------------------------------------------------------------
        // TryParseThreadDate helper tests
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Verifies that the second-precision date format written by the current exporter
        /// ("yyyy-MM-dd HH:mm:ss") is parsed correctly.
        /// </summary>
        [Fact]
        public void TryParseThreadDate_ParsesSecondPrecisionDate()
        {
            DateTime result;
            bool success = ThreadService.TryParseThreadDate("2026-03-12 09:30:45", out result);

            Assert.True(success);
            Assert.Equal(new DateTime(2026, 3, 12, 9, 30, 45), result);
        }

        /// <summary>
        /// Verifies that the legacy minute-precision date format ("yyyy-MM-dd HH:mm") used in
        /// notes generated before the seconds fix is still accepted for backward compatibility.
        /// </summary>
        [Fact]
        public void TryParseThreadDate_AcceptsLegacyMinutePrecisionDate()
        {
            DateTime result;
            bool success = ThreadService.TryParseThreadDate("2026-03-12 09:30", out result);

            Assert.True(success);
            Assert.Equal(new DateTime(2026, 3, 12, 9, 30, 0), result);
        }

        [Fact]
        public void TryParseThreadDate_ReturnsFalseForInvalidInput()
        {
            DateTime result;
            bool success = ThreadService.TryParseThreadDate("not-a-date", out result);

            Assert.False(success);
        }

        // -----------------------------------------------------------------------------------------
        // FindExistingThread tests
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Verifies that <see cref="ThreadService.FindExistingThread"/> treats a missing inbox
        /// folder as an empty state (no thread found) without throwing.
        /// </summary>
        [Fact]
        public void FindExistingThread_MissingInboxFolder_ReturnsNoThread()
        {
            string missingInbox = Path.Combine(_testDir, "DoesNotExist");

            var result = _threadService.FindExistingThread("any-conversation-id", missingInbox);

            Assert.False(result.hasExistingThread);
            Assert.Null(result.earliestEmailThreadName);
            Assert.Null(result.earliestEmailDate);
            Assert.Equal(0, result.emailCount);
        }

        /// <summary>
        /// Verifies that a note with a quoted second-precision date field is matched correctly
        /// by <see cref="ThreadService.FindExistingThread"/>.
        /// </summary>
        [Fact]
        public void FindExistingThread_ParsesQuotedSecondPrecisionDate()
        {
            string inboxPath = Path.Combine(_testDir, "Inbox");
            Directory.CreateDirectory(inboxPath);

            string noteContent =
                "---\n" +
                "title: \"Test Email\"\n" +
                "type: email\n" +
                "threadId: \"abc123def456ghi789jk\"\n" +
                "date: \"2026-03-12 09:30:45\"\n" +
                "from: \"[[Sender Person]]\"\n" +
                "to:\n" +
                "  - \"[[Recipient Person]]\"\n" +
                "---\n\n" +
                "Body text";

            File.WriteAllText(Path.Combine(inboxPath, "test-email.md"), noteContent);

            var result = _threadService.FindExistingThread("abc123def456ghi789jk", inboxPath);

            Assert.True(result.hasExistingThread);
            Assert.Equal(1, result.emailCount);
            Assert.True(result.earliestEmailDate.HasValue);
            Assert.Equal(new DateTime(2026, 3, 12, 9, 30, 45), result.earliestEmailDate.Value);
        }

        /// <summary>
        /// Verifies that a note with an unquoted legacy minute-precision date field is matched
        /// and its date parsed correctly for backward compatibility.
        /// </summary>
        [Fact]
        public void FindExistingThread_AcceptsLegacyMinutePrecisionDate()
        {
            string inboxPath = Path.Combine(_testDir, "InboxLegacy");
            Directory.CreateDirectory(inboxPath);

            string noteContent =
                "---\n" +
                "title: \"Legacy Email\"\n" +
                "type: email\n" +
                "threadId: \"legacy111222333444abc\"\n" +
                "date: 2025-01-15 14:20\n" +
                "from: \"[[Old Sender]]\"\n" +
                "to:\n" +
                "  - \"[[Old Recipient]]\"\n" +
                "---\n\n" +
                "Body text";

            File.WriteAllText(Path.Combine(inboxPath, "legacy-email.md"), noteContent);

            var result = _threadService.FindExistingThread("legacy111222333444abc", inboxPath);

            Assert.True(result.hasExistingThread);
            Assert.Equal(1, result.emailCount);
            Assert.True(result.earliestEmailDate.HasValue);
            Assert.Equal(new DateTime(2025, 1, 15, 14, 20, 0), result.earliestEmailDate.Value);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                try
                {
                    Directory.Delete(_testDir, true);
                }
                catch (System.Exception)
                {
                    // Best-effort cleanup
                }
            }
        }
    }
}
