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

        private static string BuildNote(string threadId, string date, string body = "Body text")
        {
            return
                "---\n" +
                "title: \"Budget review - 2026-09-23\"\n" +
                "type: \"email\"\n" +
                "from: \"[[Sara Arnett]]\"\n" +
                "to:\n" +
                "  - \"[[Caleb Bennett]]\"\n" +
                "threadId: \"" + threadId + "\"\n" +
                "date: \"" + date + "\"\n" +
                "---\n\n" +
                body;
        }

        /// <summary>
        /// A conversation whose only note is still loose in the inbox has no thread folder yet, so
        /// no name is reported and the caller names the folder from the email itself (the old
        /// title-based reconstruction fused the date onto the subject).
        /// </summary>
        [Fact]
        public void FindExistingThread_NoteInInboxRoot_ReportsNoFolderName()
        {
            string inboxPath = Path.Combine(_testDir, "InboxRoot");
            Directory.CreateDirectory(inboxPath);
            File.WriteAllText(Path.Combine(inboxPath, "note.md"), BuildNote("THREAD00000000000001", "2026-09-23 18:18:52"));

            var result = _threadService.FindExistingThread("THREAD00000000000001", inboxPath);

            Assert.True(result.hasExistingThread);
            Assert.Null(result.earliestEmailThreadName);
        }

        /// <summary>
        /// The existing thread folder is reused even when an older email of the conversation is
        /// sitting in the inbox root.
        /// </summary>
        [Fact]
        public void FindExistingThread_ExistingFolderWinsOverOlderRootNote()
        {
            string inboxPath = Path.Combine(_testDir, "InboxFolderWins");
            string threadFolder = Path.Combine(inboxPath, "Budget review-SaraA-CalebB");
            Directory.CreateDirectory(threadFolder);
            File.WriteAllText(Path.Combine(threadFolder, "0-Budget review-SaraA-CalebB.md"), "# Budget review");
            File.WriteAllText(Path.Combine(threadFolder, "reply.md"), BuildNote("THREAD00000000000002", "2026-09-24 09:00:00"));
            File.WriteAllText(Path.Combine(inboxPath, "older.md"), BuildNote("THREAD00000000000002", "2026-09-23 08:00:00"));

            var result = _threadService.FindExistingThread("THREAD00000000000002", inboxPath);

            Assert.Equal("Budget review-SaraA-CalebB", result.earliestEmailThreadName);
            Assert.Equal(2, result.emailCount);
            Assert.Equal(new DateTime(2026, 9, 23, 8, 0, 0), result.earliestEmailDate.Value);
        }

        /// <summary>
        /// A subfolder without a 0- summary note (e.g. a batch-sling folder) is not a thread folder.
        /// </summary>
        [Fact]
        public void FindExistingThread_NoteInNonThreadSubfolder_IsIgnored()
        {
            string inboxPath = Path.Combine(_testDir, "InboxBatch");
            string batchFolder = Path.Combine(inboxPath, "Job Alerts");
            Directory.CreateDirectory(batchFolder);
            File.WriteAllText(Path.Combine(batchFolder, "alert.md"), BuildNote("THREAD00000000000003", "2026-09-23 08:00:00"));

            var result = _threadService.FindExistingThread("THREAD00000000000003", inboxPath);

            Assert.False(result.hasExistingThread);
            Assert.Equal(0, result.emailCount);
        }

        // -----------------------------------------------------------------------------------------
        // Moving the first email of a conversation into its thread folder
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void AddThreadNoteField_AppendsFieldToFrontMatter()
        {
            string updated = ThreadService.AddThreadNoteField(BuildNote("T", "2026-09-23 18:18:52"), "Budget review-SaraA-CalebB");

            Assert.Contains("date: \"2026-09-23 18:18:52\"\nthreadNote: \"[[0-Budget review-SaraA-CalebB]]\"\n---\n", updated);
            Assert.EndsWith("Body text", updated);
        }

        [Fact]
        public void AddThreadNoteField_ExistingField_LeavesContentUntouched()
        {
            string content = "---\nthreadNote: \"[[0-Other]]\"\n---\n\nBody";

            Assert.Equal(content, ThreadService.AddThreadNoteField(content, "Budget review"));
        }

        [Fact]
        public void MoveToThreadFolder_CarriesPerNoteAttachmentFolderAndAddsThreadNote()
        {
            string inboxPath = Path.Combine(_testDir, "InboxMove");
            string threadFolder = Path.Combine(inboxPath, "Budget review-SaraA-CalebB");
            Directory.CreateDirectory(inboxPath);
            string noteName = "Budget review-SaraA-2026-09-23-1818";
            string notePath = Path.Combine(inboxPath, noteName + ".md");
            File.WriteAllText(notePath, BuildNote("T", "2026-09-23 18:18:52", "![[" + noteName + "/logo.png]]"));
            Directory.CreateDirectory(Path.Combine(inboxPath, noteName));
            File.WriteAllText(Path.Combine(inboxPath, noteName, "logo.png"), "png");

            string movedPath = _threadService.MoveToThreadFolder(notePath, threadFolder, "Budget review-SaraA-CalebB");

            string movedName = Path.GetFileNameWithoutExtension(movedPath);
            Assert.Equal(threadFolder, Path.GetDirectoryName(movedPath));
            Assert.False(Directory.Exists(Path.Combine(inboxPath, noteName)));
            Assert.True(File.Exists(Path.Combine(threadFolder, movedName, "logo.png")));
            string content = File.ReadAllText(movedPath);
            Assert.Contains("![[" + movedName + "/logo.png]]", content);
            Assert.Contains("threadNote: \"[[0-Budget review-SaraA-CalebB]]\"", content);
        }

        [Fact]
        public void MoveToThreadFolder_CarriesAttachmentsStoredBesideTheNote()
        {
            string inboxPath = Path.Combine(_testDir, "InboxMoveLoose");
            string threadFolder = Path.Combine(inboxPath, "Thread");
            Directory.CreateDirectory(inboxPath);
            string notePath = Path.Combine(inboxPath, "note.md");
            File.WriteAllText(notePath, BuildNote("T", "2026-09-23 18:18:52", "![[chart.png]] and [[Sara Arnett]]"));
            File.WriteAllText(Path.Combine(inboxPath, "chart.png"), "png");
            File.WriteAllText(Path.Combine(inboxPath, "unrelated.png"), "png");

            _threadService.MoveToThreadFolder(notePath, threadFolder);

            Assert.True(File.Exists(Path.Combine(threadFolder, "chart.png")));
            Assert.False(File.Exists(Path.Combine(inboxPath, "chart.png")));
            Assert.True(File.Exists(Path.Combine(inboxPath, "unrelated.png")));
        }

        /// <summary>
        /// The first email of a thread (written standalone, date at the end) and later replies
        /// (written threaded, date at the front) end up with the same name shape, and a renamed
        /// note keeps its attachment folder.
        /// </summary>
        [Fact]
        public void ResuffixThreadNotes_NamesFirstAndLaterNotesConsistently()
        {
            string threadFolder = Path.Combine(_testDir, "Resuffix");
            Directory.CreateDirectory(threadFolder);
            string first = "Budget review-SaraA-2026-09-23-1818-eidABC123";
            File.WriteAllText(Path.Combine(threadFolder, first + ".md"),
                BuildNote("T", "2026-09-23 18:18:52", "![[" + first + "/logo.png]]"));
            Directory.CreateDirectory(Path.Combine(threadFolder, first));
            File.WriteAllText(Path.Combine(threadFolder, first, "logo.png"), "png");
            string replyPath = Path.Combine(threadFolder, "2026-09-24-0900-Budget review-CalebB-eidDEF456.md");
            File.WriteAllText(replyPath, BuildNote("T", "2026-09-24 09:00:05"));

            string newReplyPath = _threadService.ResuffixThreadNotes(threadFolder, null, replyPath);

            string firstFinal = "2026-09-23_181852_Budget review-SaraA";
            Assert.True(File.Exists(Path.Combine(threadFolder, firstFinal + ".md")));
            Assert.Equal(Path.Combine(threadFolder, "2026-09-24_090005_Budget review-CalebB.md"), newReplyPath);
            Assert.True(File.Exists(Path.Combine(threadFolder, firstFinal, "logo.png")));
            Assert.Contains("![[" + firstFinal + "/logo.png]]", File.ReadAllText(Path.Combine(threadFolder, firstFinal + ".md")));
        }

        // -----------------------------------------------------------------------------------------
        // Notes grouped with their attachments (GroupNoteWithAttachments)
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void FindExistingThread_NoteGroupedWithAttachments_CountsAsInboxNote()
        {
            string inboxPath = Path.Combine(_testDir, "InboxGrouped");
            string noteFolder = Path.Combine(inboxPath, "Budget review-SaraA-2026-09-23-1818");
            Directory.CreateDirectory(noteFolder);
            File.WriteAllText(Path.Combine(noteFolder, "Budget review-SaraA-2026-09-23-1818.md"), BuildNote("THREAD00000000000004", "2026-09-23 18:18:52"));

            var result = _threadService.FindExistingThread("THREAD00000000000004", inboxPath);

            Assert.True(result.hasExistingThread);
            Assert.Equal(1, result.emailCount);
            Assert.Null(result.earliestEmailThreadName);
        }

        [Fact]
        public void MoveToThreadFolder_GroupedNote_TurnsItsFolderIntoAttachmentFolder()
        {
            string inboxPath = Path.Combine(_testDir, "InboxMoveGrouped");
            string threadFolder = Path.Combine(inboxPath, "Budget review-SaraA-CalebB");
            string noteName = "Budget review-SaraA-2026-09-23-1818";
            string noteFolder = Path.Combine(inboxPath, noteName);
            Directory.CreateDirectory(noteFolder);
            string notePath = Path.Combine(noteFolder, noteName + ".md");
            File.WriteAllText(notePath, BuildNote("T", "2026-09-23 18:18:52", "![[logo.png]] and [report](Q3%20report.pdf)"));
            File.WriteAllText(Path.Combine(noteFolder, "logo.png"), "png");
            File.WriteAllText(Path.Combine(noteFolder, "Q3 report.pdf"), "pdf");

            string movedPath = _threadService.MoveToThreadFolder(notePath, threadFolder, "Budget review-SaraA-CalebB");

            Assert.Equal(Path.Combine(threadFolder, noteName + ".md"), movedPath);
            Assert.False(Directory.Exists(noteFolder));
            Assert.True(File.Exists(Path.Combine(threadFolder, noteName, "logo.png")));
            Assert.True(File.Exists(Path.Combine(threadFolder, noteName, "Q3 report.pdf")));
            string content = File.ReadAllText(movedPath);
            Assert.Contains("![[" + noteName + "/logo.png]]", content);
            Assert.Contains("](Budget%20review-SaraA-2026-09-23-1818/Q3%20report.pdf)", content);
        }

        // -----------------------------------------------------------------------------------------
        // Stable note names
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void GetThreadedNoteName_StandaloneAndThreadedBaseNames_GiveTheSameShape()
        {
            DateTime received = new DateTime(2026, 9, 23, 18, 18, 52);

            Assert.Equal("2026-09-23_181852_Budget review-SaraA",
                _threadService.GetThreadedNoteName("Budget review-SaraA-2026-09-23-1818", received));
            Assert.Equal("2026-09-23_181852_Budget review-SaraA",
                _threadService.GetThreadedNoteName("2026-09-23-1818-Budget review-SaraA", received));
        }

        /// <summary>
        /// A standalone note written under its threaded name keeps that exact name through the move
        /// into a thread folder and the rename pass, so links to it by name still resolve.
        /// </summary>
        [Fact]
        public void MoveThenResuffix_NoteWrittenWithThreadedName_KeepsItsName()
        {
            string inboxPath = Path.Combine(_testDir, "InboxStable");
            string threadFolder = Path.Combine(inboxPath, "Budget review-SaraA-CalebB");
            Directory.CreateDirectory(inboxPath);
            string noteName = _threadService.GetThreadedNoteName("Budget review-SaraA-2026-09-23-1818", new DateTime(2026, 9, 23, 18, 18, 52));
            string notePath = Path.Combine(inboxPath, noteName + ".md");
            string note = BuildNote("T", "2026-09-23 18:18:52", "- [ ] [[" + noteName + "]] follow up")
                .Replace("date:", "internetMessageId: \"<abc123@example.com>\"\ndate:");
            File.WriteAllText(notePath, note);

            _threadService.MoveToThreadFolder(notePath, threadFolder, "Budget review-SaraA-CalebB");
            _threadService.ResuffixThreadNotes(threadFolder, null);

            string finalPath = Path.Combine(threadFolder, noteName + ".md");
            Assert.True(File.Exists(finalPath));
            Assert.Contains("[[" + noteName + "]] follow up", File.ReadAllText(finalPath));
        }

        /// <summary>
        /// The summary note renders through Dataview, so an existing one is never rewritten — that used
        /// to wipe anything the user had added to it on every sling into the thread.
        /// </summary>
        [Fact]
        public void UpdateThreadNote_ExistingSummary_KeepsUserContent()
        {
            string threadFolder = Path.Combine(_testDir, "Vault", "Inbox", "Budget review-SaraA-CalebB");
            Directory.CreateDirectory(threadFolder);
            string summaryPath = Path.Combine(threadFolder, "0-Budget review-SaraA-CalebB.md");
            string userContent = "---\nthreadId: \"T\"\n---\n\nMy notes about this thread";
            File.WriteAllText(summaryPath, userContent);

            _threadService.UpdateThreadNote(threadFolder, summaryPath, "T", "Budget review-SaraA-CalebB", null).Wait();

            Assert.Equal(userContent, File.ReadAllText(summaryPath));
        }

        // -----------------------------------------------------------------------------------------
        // Path length budget
        // -----------------------------------------------------------------------------------------

        private const string LongBaseName =
            "Quarterly loan review documentation request for the downtown branch and the follow up items-SaraA";

        /// <summary>
        /// Returns a thread folder path of about <paramref name="length"/> characters under <paramref name="inboxPath"/>.
        /// </summary>
        private static string DeepThreadFolder(string inboxPath, int length)
        {
            return Path.Combine(inboxPath, "Thread" + new string('x', Math.Max(1, length - inboxPath.Length - 7)));
        }

        [Fact]
        public void GetTemporaryThreadNoteName_LongMessageId_UsesShortHashTheRenamePassStrips()
        {
            string id = "<BN6PR11MB1234567890ABCDEF1234567890ABCDEF@BN6PR11MB1234.namprd11.prod.outlook.com>";

            string name = _threadService.GetTemporaryThreadNoteName(Path.Combine(_testDir, "Inbox", "T"), "Budget review-SaraA", id);

            Assert.Equal(12, ThreadService.HashEmailId(id).Length);
            Assert.Equal("Budget review-SaraA-eid" + ThreadService.HashEmailId(id), name);
            Assert.Equal("Budget review-SaraA", new SlingMD.Outlook.Services.Formatting.LegacyFilenameStripper().Strip(name));
        }

        [Fact]
        public void ResuffixThreadNotes_LongNameInDeepFolder_KeepsPathWithinBudget()
        {
            string threadFolder = DeepThreadFolder(Path.Combine(_testDir, "InboxLong"), 130);
            Directory.CreateDirectory(threadFolder);
            string notePath = Path.Combine(threadFolder, LongBaseName + "-eidABC123.md");
            File.WriteAllText(notePath, BuildNote("T", "2026-09-23 18:18:52"));

            string newPath = _threadService.ResuffixThreadNotes(threadFolder, null, notePath);

            Assert.True(File.Exists(newPath));
            Assert.True(newPath.Length <= SlingMD.Outlook.Services.Formatting.PathLengthBudget.MaxFilePath);
            Assert.StartsWith("2026-09-23_181852_Quarterly loan review", Path.GetFileName(newPath));
        }

        /// <summary>
        /// A name shortened for a deep thread folder survives the move and the rename pass unchanged,
        /// so links to it by name keep resolving.
        /// </summary>
        [Fact]
        public void MoveThenResuffix_LongThreadedName_KeepsItsShortenedName()
        {
            string inboxPath = Path.Combine(_testDir, "InboxStableLong");
            string threadFolder = DeepThreadFolder(inboxPath, 130);
            Directory.CreateDirectory(inboxPath);
            string noteName = _threadService.GetThreadedNoteName(LongBaseName + "-2026-09-23-1818", new DateTime(2026, 9, 23, 18, 18, 52), threadFolder);
            string notePath = Path.Combine(inboxPath, noteName + ".md");
            string note = BuildNote("T", "2026-09-23 18:18:52")
                .Replace("date:", "internetMessageId: \"<BN6PR11MB1234567890ABCDEF@BN6PR11MB1234.namprd11.prod.outlook.com>\"\ndate:");
            File.WriteAllText(notePath, note);

            _threadService.MoveToThreadFolder(notePath, threadFolder, "Budget review-SaraA-CalebB");
            _threadService.ResuffixThreadNotes(threadFolder, null);

            string finalPath = Path.Combine(threadFolder, noteName + ".md");
            Assert.True(noteName.Length < 18 + LongBaseName.Length);
            Assert.True(File.Exists(finalPath));
            Assert.True(finalPath.Length <= SlingMD.Outlook.Services.Formatting.PathLengthBudget.MaxFilePath);
        }
    }
}
