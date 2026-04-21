using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Unit tests for <see cref="ThreadCompletionService"/> focusing on constructor validation
    /// and vault-scanning logic that does not require COM objects.
    /// </summary>
    public class ThreadCompletionServiceTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly ThreadCompletionService _completionService;

        public ThreadCompletionServiceTests()
        {
            _testDir = Path.Combine(
                Path.GetTempPath(),
                "SlingMDTests",
                "ThreadCompletion_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                InboxFolder = "Inbox",
                ShowThreadDebug = false
            };
            _fileService = new FileService(_settings);
            _completionService = new ThreadCompletionService(_fileService, _settings);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testDir))
                {
                    Directory.Delete(_testDir, true);
                }
            }
            catch (System.Exception)
            {
                // Best-effort cleanup
            }
        }

        // -----------------------------------------------------------------------------------------
        // Constructor tests
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void ThreadCompletionService_Constructor_WithValidSettings_DoesNotThrow()
        {
            // Arrange & Act
            ThreadCompletionService service = new ThreadCompletionService(_fileService, _settings);

            // Assert: no exception thrown; instance is not null
            Assert.NotNull(service);
        }

        // -----------------------------------------------------------------------------------------
        // GetExistingEntryIds tests
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void GetExistingEntryIds_EmptyVault_ReturnsEmptySet()
        {
            // Arrange: inbox path does not exist
            string conversationId = "abc123";

            // Act
            HashSet<string> result = _completionService.GetExistingEntryIds(conversationId);

            // Assert
            Assert.NotNull(result);
            Assert.Empty(result);
        }

        [Fact]
        public void GetExistingEntryIds_WithMatchingNote_ReturnsEntryId()
        {
            // Arrange: create inbox directory and a note with matching threadId + entryId
            string inboxPath = _settings.GetInboxPath();
            Directory.CreateDirectory(inboxPath);

            string conversationId = "deadbeef1234567890ab";
            string entryId = "ABCDEF0123456789";

            string noteContent = string.Join(Environment.NewLine,
                "---",
                $"threadId: \"{conversationId}\"",
                $"entryId: \"{entryId}\"",
                "title: \"Test Email\"",
                "---",
                "Body text here.");

            File.WriteAllText(Path.Combine(inboxPath, "test-email.md"), noteContent);

            // Act
            HashSet<string> result = _completionService.GetExistingEntryIds(conversationId);

            // Assert
            Assert.Contains(entryId, result);
        }

        [Fact]
        public void GetExistingEntryIds_NonMatchingThreadId_ReturnsEmptySet()
        {
            // Arrange: create a note with a DIFFERENT threadId
            string inboxPath = _settings.GetInboxPath();
            Directory.CreateDirectory(inboxPath);

            string noteContent = string.Join(Environment.NewLine,
                "---",
                "threadId: \"differentThreadId\"",
                "entryId: \"ABCDEF0123456789\"",
                "---",
                "Body text here.");

            File.WriteAllText(Path.Combine(inboxPath, "other-email.md"), noteContent);

            // Act: query for a DIFFERENT conversation ID
            HashSet<string> result = _completionService.GetExistingEntryIds("targetConversationId");

            // Assert
            Assert.Empty(result);
        }

        [Fact]
        public void GetExistingEntryIds_MultipleMatchingNotes_ReturnsAllEntryIds()
        {
            // Arrange: create two notes with same threadId but different entryIds
            string inboxPath = _settings.GetInboxPath();
            Directory.CreateDirectory(inboxPath);

            string conversationId = "multitest1234567890ab";
            string entryId1 = "ENTRY1111";
            string entryId2 = "ENTRY2222";

            string note1 = string.Join(Environment.NewLine,
                "---",
                $"threadId: \"{conversationId}\"",
                $"entryId: \"{entryId1}\"",
                "---",
                "First email.");

            string note2 = string.Join(Environment.NewLine,
                "---",
                $"threadId: \"{conversationId}\"",
                $"entryId: \"{entryId2}\"",
                "---",
                "Second email.");

            File.WriteAllText(Path.Combine(inboxPath, "email-1.md"), note1);
            File.WriteAllText(Path.Combine(inboxPath, "email-2.md"), note2);

            // Act
            HashSet<string> result = _completionService.GetExistingEntryIds(conversationId);

            // Assert
            Assert.Contains(entryId1, result);
            Assert.Contains(entryId2, result);
            Assert.Equal(2, result.Count);
        }
    }
}
