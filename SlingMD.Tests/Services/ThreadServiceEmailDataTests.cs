using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// <see cref="ThreadService"/> on <see cref="EmailData"/> (SlingMD.Core): thread folder and note
    /// names, the summary note title and folder path, and the thread-debug message through
    /// <see cref="IUserNotifier"/>. The expected names are what the add-in's ThreadService produced
    /// from a <c>MailItem</c> with the same topic, subject and received time before the move
    /// (checked under Mono with Moq doubles). Linked into SlingMD.Core.Tests.
    /// </summary>
    public class ThreadServiceEmailDataTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;

        public ThreadServiceEmailDataTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ThreadEmailData_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(_testDir);
            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                InboxFolder = "Inbox",
                ShowThreadDebug = false
            };
            _fileService = new FileService(_settings);
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

        private ThreadService NewService(IUserNotifier notifier = null)
        {
            return new ThreadService(_fileService, new TemplateService(_fileService), _settings, notifier);
        }

        private static EmailData Email(string topic, string subject)
        {
            return new EmailData
            {
                ConversationTopic = topic,
                Subject = subject,
                ReceivedTime = new DateTime(2026, 9, 23, 18, 18, 52)
            };
        }

        private sealed class RecordingNotifier : IUserNotifier
        {
            public readonly List<string> Shown = new List<string>();

            public void ShowWarning(string message, string title)
            {
                Shown.Add("warning|" + title + "|" + message);
            }

            public void ShowInformation(string message, string title)
            {
                Shown.Add("info|" + title + "|" + message);
            }
        }

        [Theory]
        [InlineData("Budget review", "RE: Other", "Budget review-SaraA-CalebB")]
        [InlineData("", "RE: Budget: Q3 / review?", "Budget_ Q3 _ review-SaraA-CalebB")]
        [InlineData(null, "Fwd: Plan", "Plan-SaraA-CalebB")]
        [InlineData("A very long conversation topic that goes on and on beyond fifty characters for sure", "x", "A very long conversation topic that goes on and...-SaraA-CalebB")]
        [InlineData("Résumé — naïve café 日本語", "x", "Résumé — naïve café 日本語-SaraA-CalebB")]
        public void GetThreadFolderName_TopicElseSubject(string topic, string subject, string expected)
        {
            Assert.Equal(expected, NewService().GetThreadFolderName(Email(topic, subject), "ignored", "SaraA", "CalebB"));
        }

        [Theory]
        [InlineData("Budget review", "RE: Other", "2026-09-23 1818-Budget review-SaraA-CalebB")]
        [InlineData("", "RE: Budget: Q3 / review?", "2026-09-23 1818-Budget_ Q3 _ review-SaraA-CalebB")]
        [InlineData(null, "Fwd: Plan", "2026-09-23 1818-Plan-SaraA-CalebB")]
        public void GetThreadNoteName_TimestampThenFolderName(string topic, string subject, string expected)
        {
            Assert.Equal(expected, NewService().GetThreadNoteName(Email(topic, subject), "ignored", "SaraA", "CalebB"));
        }

        [Fact]
        public void GetThreadFolderName_TrailingDotInNames_IsDropped()
        {
            Assert.Equal("Budget-Robert Smith Jr.-Jr", NewService().GetThreadFolderName(Email("Budget", "x"), "", "Robert Smith Jr.", "Jr."));
        }

        [Fact]
        public void UpdateThreadNote_NoTopic_TitleFromSubjectAndVaultRelativeFolder()
        {
            string folder = Path.Combine(_testDir, "Vault", "Inbox", "T1");
            Directory.CreateDirectory(folder);
            string summary = Path.Combine(folder, "0-T1.md");

            NewService().UpdateThreadNote(folder, summary, "ID1", "T1", Email(null, "RE: Budget")).Wait();

            // The built-in thread template is a verbatim string, so its line endings are whatever the
            // source file was checked out with (CRLF on Windows with text=auto); compare with \n only.
            string content = File.ReadAllText(summary).Replace("\r\n", "\n");
            Assert.StartsWith("---\ntitle: \"Budget\"\ntype: email-thread\nthreadId: \"ID1\"\ntags: [email-thread]\n---\n\n# Budget\n\n", content);
            Assert.Contains("dv.pages('\"Inbox/T1\"')", content);
        }

        [Fact]
        public void UpdateThreadNote_EmptyTopic_TitleFromSubject()
        {
            // The folder name already treats an empty topic like a missing one; the title used ??
            // and came out empty ("title: \"\"", "# ").
            string folder = Path.Combine(_testDir, "Vault", "Inbox", "T2");
            Directory.CreateDirectory(folder);
            string summary = Path.Combine(folder, "0-T2.md");

            NewService().UpdateThreadNote(folder, summary, "ID2", "T2", Email("", "RE: Budget")).Wait();

            string content = File.ReadAllText(summary).Replace("\r\n", "\n");
            Assert.StartsWith("---\ntitle: \"Budget\"\n", content);
            Assert.Contains("\n# Budget\n", content);
        }

        [Fact]
        public void FindExistingThread_ShowThreadDebug_ShowsTheMatchesThroughTheNotifier()
        {
            _settings.ShowThreadDebug = true;
            string inbox = Path.Combine(_testDir, "Vault", "Inbox");
            Directory.CreateDirectory(inbox);
            File.WriteAllText(Path.Combine(inbox, "a.md"), "---\nthreadId: \"ID1\"\ndate: \"2026-09-23 18:18:52\"\n---\n\nBody");
            RecordingNotifier notifier = new RecordingNotifier();

            var result = NewService(notifier).FindExistingThread("ID1", inbox);

            Assert.True(result.hasExistingThread);
            Assert.Single(notifier.Shown);
            Assert.StartsWith("info|Thread Match Details|Found 1 existing emails with thread ID: ID1", notifier.Shown[0]);
        }

        [Fact]
        public void FindExistingThread_DebugOff_ShowsNothing()
        {
            string inbox = Path.Combine(_testDir, "Vault", "Inbox");
            Directory.CreateDirectory(inbox);
            File.WriteAllText(Path.Combine(inbox, "a.md"), "---\nthreadId: \"ID1\"\ndate: \"2026-09-23 18:18:52\"\n---\n\nBody");
            RecordingNotifier notifier = new RecordingNotifier();

            NewService(notifier).FindExistingThread("ID1", inbox);

            Assert.Empty(notifier.Shown);
        }
    }
}
