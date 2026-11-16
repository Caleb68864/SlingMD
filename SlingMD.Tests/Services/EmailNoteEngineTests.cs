using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Characterization tests for <see cref="EmailNoteEngine"/>, the COM-free half of
    /// EmailProcessor: given an <see cref="EmailData"/>, the exact title, file name, front matter
    /// and note text. The expected values are what EmailProcessor produced before the split (same
    /// code, moved). Linked into SlingMD.Core.Tests, so they run on Linux.
    /// </summary>
    public class EmailNoteEngineTests : IDisposable
    {
        private readonly string _testDir;
        private readonly string _templatesDir;

        public EmailNoteEngineTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "EmailNoteEngine", Guid.NewGuid().ToString("N"));
            _templatesDir = Path.Combine(_testDir, "Vault", "Templates");
            Directory.CreateDirectory(_templatesDir);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_testDir, true);
            }
            catch (IOException)
            {
            }
        }

        private ObsidianSettings NewSettings()
        {
            return new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                TemplatesFolder = "Templates",
                IncludeDailyNoteLink = false,
                DefaultNoteTags = new List<string> { "FollowUp" }
            };
        }

        private static EmailNoteEngine NewEngine(ObsidianSettings settings)
        {
            TemplateService templates = new TemplateService(new FileService(settings));
            return new EmailNoteEngine(settings, templates, (name, email) => $"[[{name}]]");
        }

        private static EmailData SampleEmail()
        {
            return new EmailData
            {
                Subject = "Re: Budget review",
                SenderName = "Sara Adams",
                SenderEmail = "sara@example.com",
                ReceivedTime = new DateTime(2026, 9, 23, 18, 18, 52),
                HtmlBody = null,
                Body = "Numbers attached.",
                EntryId = "ENTRY1",
                InternetMessageId = "<msg-1@example.com>",
                FirstToRecipientName = "Caleb Bennett"
            };
        }

        [Fact]
        public void BuildNoteTitle_DefaultFormat_SubjectAndDate()
        {
            EmailNoteEngine engine = NewEngine(NewSettings());

            Assert.Equal("Budget review - 2026-09-23", engine.BuildNoteTitle("Budget review", "SaraA", new DateTime(2026, 9, 23, 18, 18, 52)));
        }

        [Fact]
        public void BuildNoteTitle_DateExcluded_DropsTrailingSeparator()
        {
            ObsidianSettings settings = NewSettings();
            settings.NoteTitleIncludeDate = false;
            EmailNoteEngine engine = NewEngine(settings);

            Assert.Equal("Budget review", engine.BuildNoteTitle("Budget review", "SaraA", new DateTime(2026, 9, 23)));
        }

        [Fact]
        public void BuildBaseFileName_NotThreaded_SubjectSenderTimestamp()
        {
            EmailNoteEngine engine = NewEngine(NewSettings());

            Assert.Equal("Budget review-SaraA-2026-09-23-1818",
                engine.BuildBaseFileName(new DateTime(2026, 9, 23, 18, 18, 52), "Budget review", "SaraA", "2026-09-23-1818", false));
        }

        [Fact]
        public void BuildBaseFileName_ThreadedWithDateMovedToFront_TimestampFirst()
        {
            ObsidianSettings settings = NewSettings();
            settings.MoveDateToFrontInThread = true;
            EmailNoteEngine engine = NewEngine(settings);

            Assert.Equal("2026-09-23-1818-Budget review-SaraA",
                engine.BuildBaseFileName(new DateTime(2026, 9, 23, 18, 18, 52), "Budget review", "SaraA", "2026-09-23-1818", true));
        }

        [Fact]
        public void BuildMetadata_FullEmail_KeysAndValuesInOrder()
        {
            EmailNoteEngine engine = NewEngine(NewSettings());

            Dictionary<string, object> metadata = engine.BuildMetadata(
                "Budget review - 2026-09-23", "Sara Adams", "sara@example.com",
                new List<string> { "[[Caleb Bennett]]" }, new List<string> { "caleb@example.com" },
                "thread-1", new DateTime(2026, 9, 23, 18, 18, 52), "<msg-1@example.com>", "ENTRY1",
                new List<string> { "[[Dana Lee]]" }, new List<string> { "dana@example.com" },
                true, "Budget review-SaraA-CalebB");

            Assert.Equal(
                new[] { "title", "type", "from", "fromEmail", "to", "toEmail", "threadId", "date", "internetMessageId", "entryId", "tags", "cc", "ccEmail", "threadNote" },
                new List<string>(metadata.Keys).ToArray());
            Assert.Equal("[[Sara Adams]]", metadata["from"]);
            Assert.Equal("email", metadata["type"]);
            Assert.Equal("[[0-Budget review-SaraA-CalebB]]", metadata["threadNote"]);
            Assert.Equal(new List<string> { "FollowUp" }, metadata["tags"]);
        }

        [Fact]
        public void BuildMetadata_NoCcNotGrouped_OmitsOptionalKeys()
        {
            ObsidianSettings settings = NewSettings();
            settings.DefaultNoteTags = new List<string>();
            EmailNoteEngine engine = NewEngine(settings);

            Dictionary<string, object> metadata = engine.BuildMetadata(
                "T", "S", "s@example.com", new List<string>(), new List<string>(), "thread-1",
                new DateTime(2026, 9, 23), null, "ENTRY1", new List<string>(), new List<string>(), false, null);

            Assert.False(metadata.ContainsKey("tags"));
            Assert.False(metadata.ContainsKey("cc"));
            Assert.False(metadata.ContainsKey("threadNote"));
            Assert.False(metadata.ContainsKey("dailyNoteLink"));
        }

        [Fact]
        public void BuildMetadata_DailyNoteLink_UsesConfiguredFormat()
        {
            ObsidianSettings settings = NewSettings();
            settings.IncludeDailyNoteLink = true;
            settings.DailyNoteLinkFormat = "[[yyyy-MM-dd]]";
            EmailNoteEngine engine = NewEngine(settings);

            Dictionary<string, object> metadata = engine.BuildMetadata(
                "T", "S", "s@example.com", new List<string>(), new List<string>(), "thread-1",
                new DateTime(2026, 9, 23), null, "ENTRY1", null, null, false, null);

            Assert.Equal("[[2026-09-23]]", metadata["dailyNoteLink"]);
        }

        [Fact]
        public void ConvertBody_PlainText_PassedThrough()
        {
            EmailNoteEngine engine = NewEngine(NewSettings());
            Dictionary<string, object> metadata = new Dictionary<string, object>();

            string body = engine.ConvertBody(SampleEmail(), metadata);

            Assert.Equal("Numbers attached.", body);
            Assert.False(metadata.ContainsKey("bodyCleaned"));
        }

        [Fact]
        public void ConvertBody_CleanupPatternMatches_RemovesTextAndSetsFlag()
        {
            ObsidianSettings settings = NewSettings();
            settings.BodyCleanupPatterns = new List<string> { "^CAUTION:.*$" };
            EmailNoteEngine engine = NewEngine(settings);
            EmailData email = SampleEmail();
            email.Body = "CAUTION: External sender.\n\nNumbers attached.";
            Dictionary<string, object> metadata = new Dictionary<string, object>();

            string body = engine.ConvertBody(email, metadata);

            Assert.DoesNotContain("CAUTION", body);
            Assert.Contains("Numbers attached.", body);
            Assert.Equal(true, metadata["bodyCleaned"]);
        }

        [Fact]
        public void RenderNote_VaultTemplate_FillsEveryTokenFromEmailData()
        {
            File.WriteAllText(Path.Combine(_templatesDir, "EmailTemplate.md"),
                "{{frontmatter}}|{{noteTitle}}|{{subject}}|{{senderName}}|{{senderShortName}}|{{senderEmail}}|{{date}}|{{timestamp}}|{{taskBlock}}{{body}}|{{fileName}}|{{fileNameNoExt}}|{{threadNote}}|{{threadId}}");
            EmailNoteEngine engine = NewEngine(NewSettings());
            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", "Budget review - 2026-09-23" },
                { "type", "email" }
            };

            string note = engine.RenderNote(SampleEmail(), metadata, "Budget review - 2026-09-23", "Budget review", "SaraA",
                "Numbers attached.", "Budget review-SaraA-2026-09-23-1818", "thread-1", "[[0-Budget review]]", "- [ ] task\n\n");

            string expected =
                "---\n"
                + "title: \"Budget review - 2026-09-23\"\n"
                + "type: \"email\"\n"
                + "---\n"
                + "\n"
                + "|Budget review - 2026-09-23|Budget review|Sara Adams|SaraA|sara@example.com|2026-09-23|2026-09-23T18:18:52|- [ ] task\n"
                + "\n"
                + "Numbers attached.|Budget review-SaraA-2026-09-23-1818.md|Budget review-SaraA-2026-09-23-1818|[[0-Budget review]]|thread-1";

            Assert.Equal(expected, note.Replace("\r\n", "\n"));
        }

        [Fact]
        public void RenderNote_DefaultTemplate_ExactNote()
        {
            ObsidianSettings settings = NewSettings();
            // The built-in default template, written into the vault so that no EmailTemplate.md beside
            // the test assembly can take its place.
            File.WriteAllText(Path.Combine(_templatesDir, "EmailTemplate.md"),
                new TemplateService(new FileService(settings)).GetDefaultEmailTemplate());
            EmailNoteEngine engine = NewEngine(settings);
            EmailData email = SampleEmail();
            Dictionary<string, object> metadata = engine.BuildMetadata(
                "Budget review - 2026-09-23", email.SenderName, email.SenderEmail,
                new List<string> { "[[Caleb Bennett]]" }, new List<string> { "caleb@example.com" },
                "thread-1", email.ReceivedTime, email.InternetMessageId, email.EntryId,
                new List<string>(), new List<string>(), false, null);
            string body = engine.ConvertBody(email, metadata);

            string note = engine.RenderNote(email, metadata, "Budget review - 2026-09-23", "Budget review", "SaraA",
                body, "Budget review-SaraA-2026-09-23-1818", "thread-1", string.Empty, string.Empty);

            string expected =
                "---\n"
                + "title: \"Budget review - 2026-09-23\"\n"
                + "type: \"email\"\n"
                + "from: \"[[Sara Adams]]\"\n"
                + "fromEmail: \"sara@example.com\"\n"
                + "to: \n"
                + "  - \"[[Caleb Bennett]]\"\n"
                + "toEmail: \n"
                + "  - \"caleb@example.com\"\n"
                + "threadId: \"thread-1\"\n"
                + "date: \"2026-09-23T18:18:52\"\n"
                + "internetMessageId: \"<msg-1@example.com>\"\n"
                + "entryId: \"ENTRY1\"\n"
                + "tags: \n"
                + "  - \"FollowUp\"\n"
                + "---\n"
                + "\n"
                + "Numbers attached.";

            Assert.Equal(expected, note.Replace("\r\n", "\n"));
        }
    }
}
