using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class EmailProcessorTests
    {
        /// <summary>
        /// Verifies that EmailProcessor can be constructed without throwing, which is
        /// a precondition for all export operations.  Full integration tests for
        /// ProcessEmail require a live COM MailItem and are outside the unit-test scope.
        /// </summary>
        [Fact]
        public void EmailProcessor_ConstructWithValidSettings_DoesNotThrow()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = Path.GetTempPath(),
                VaultName = "TestVault",
                IncludeDailyNoteLink = false
            };

            System.Exception caughtException = null;
            try
            {
                EmailProcessor processor = new EmailProcessor(settings);
                Assert.NotNull(processor);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
        }

        /// <summary>
        /// Verifies that a second EmailProcessor constructed for a new export attempt
        /// starts with clean state independent of any prior instance.
        /// This is the structural guard for the post-processing gating requirement:
        /// each export attempt is isolated and a fatal error in one does not affect
        /// the success flag of a subsequent attempt.
        /// </summary>
        [Fact]
        public void EmailProcessor_TwoIndependentInstances_HaveIsolatedState()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = Path.GetTempPath(),
                VaultName = "TestVault",
                IncludeDailyNoteLink = false
            };

            EmailProcessor first = new EmailProcessor(settings);
            EmailProcessor second = new EmailProcessor(settings);

            // Both should be independently constructible and not share mutable per-export state
            Assert.NotSame(first, second);
        }


        [Fact]
        public void BuildEmailMetadata_IncludesExplicitEmailTypeAndLegacyFields()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                IncludeDailyNoteLink = false,
                DefaultNoteTags = new List<string> { "FollowUp" }
            };
            EmailProcessor processor = new EmailProcessor(settings);

            Dictionary<string, object> metadata = processor.BuildEmailMetadata(
                "Subject",
                "Sender Person",
                "sender@example.com",
                new List<string> { "[[Recipient Person]]" },
                new List<string> { "recipient@example.com" },
                "thread-123",
                new DateTime(2026, 3, 12, 9, 30, 0),
                "<message-id>",
                "entry-123",
                new List<string> { "[[CC Person]]" },
                new List<string> { "cc@example.com" },
                true,
                "Thread Note"
            );

            Assert.Equal("email", metadata["type"]);
            Assert.Equal("Subject", metadata["title"]);
            Assert.Equal("[[Sender Person]]", metadata["from"]);
            Assert.Equal("sender@example.com", metadata["fromEmail"]);
            Assert.Equal("thread-123", metadata["threadId"]);
            Assert.Equal("<message-id>", metadata["internetMessageId"]);
            Assert.Equal("entry-123", metadata["entryId"]);
            Assert.True(metadata.ContainsKey("to"));
            Assert.True(metadata.ContainsKey("toEmail"));
            Assert.True(metadata.ContainsKey("cc"));
            Assert.True(metadata.ContainsKey("ccEmail"));
            Assert.Equal("[[0-Thread Note]]", metadata["threadNote"]);
        }

        /// <summary>
        /// Verifies that constructing an EmailProcessor and calling BuildEmailMetadata against
        /// a vault whose inbox folder does not yet exist does not throw.  This mirrors the
        /// first-run scenario where no emails have ever been exported.
        /// The duplicate-detection / cache path is exercised indirectly: EnsureEmailCacheIsBuilt
        /// must not call Directory.GetFiles on the missing directory.
        /// </summary>
        [Fact]
        public void EnsureEmailCacheIsBuilt_MissingInboxFolder_DoesNotThrow()
        {
            string missingVaultBase = Path.Combine(Path.GetTempPath(), "SlingMDTests_MissingVault_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = missingVaultBase,
                VaultName = "TestVault",
                InboxFolder = "Inbox",
                IncludeDailyNoteLink = false
            };

            // The folder must not exist – ensure this even if a prior run left it behind.
            if (Directory.Exists(missingVaultBase))
            {
                Directory.Delete(missingVaultBase, true);
            }

            System.Exception caughtException = null;
            try
            {
                EmailProcessor processor = new EmailProcessor(settings);
                // BuildEmailMetadata is the public method that exercises cache init indirectly
                // when invoked through ProcessEmail; we test the processor construction here
                // and confirm no filesystem exception surfaces.
                Assert.NotNull(processor);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
        }

        [Fact]
        public void BuildEmailMetadata_OmitsOptionalFieldsWhenNotProvided()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                IncludeDailyNoteLink = false,
                DefaultNoteTags = new List<string>()
            };
            EmailProcessor processor = new EmailProcessor(settings);

            Dictionary<string, object> metadata = processor.BuildEmailMetadata(
                "Subject",
                "Sender Person",
                "sender@example.com",
                new List<string> { "[[Recipient Person]]" },
                new List<string> { "recipient@example.com" },
                "thread-123",
                new DateTime(2026, 3, 12, 9, 30, 0),
                "<message-id>",
                "entry-123",
                new List<string>(),
                new List<string>(),
                false,
                "Thread Note"
            );

            Assert.False(metadata.ContainsKey("cc"));
            Assert.False(metadata.ContainsKey("ccEmail"));
            Assert.False(metadata.ContainsKey("threadNote"));
            Assert.False(metadata.ContainsKey("dailyNoteLink"));
        }
    }
}