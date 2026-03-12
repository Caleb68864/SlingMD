using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class EmailProcessorTests
    {
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