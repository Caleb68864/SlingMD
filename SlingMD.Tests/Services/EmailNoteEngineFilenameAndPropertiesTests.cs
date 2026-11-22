using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// The settings brought in from Steve O'Sullivan's fork (steveosullivan/SlingMD 4b4026e): the
    /// {SenderFull} filename token, a date format for the filename's {Date}, and leaving the Outlook
    /// ids and threadId out of the note properties. Defaults leave existing notes unchanged.
    /// Linked into SlingMD.Core.Tests.
    /// </summary>
    public class EmailNoteEngineFilenameAndPropertiesTests : IDisposable
    {
        private static readonly DateTime Received = new DateTime(2026, 9, 23, 18, 18, 52);
        private readonly string _testDir;

        public EmailNoteEngineFilenameAndPropertiesTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "EmailFilenameProps", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(_testDir, "Vault"));
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
                IncludeDailyNoteLink = false,
                DefaultNoteTags = new List<string>()
            };
        }

        private static EmailNoteEngine NewEngine(ObsidianSettings settings)
        {
            return new EmailNoteEngine(settings, new TemplateService(new FileService(settings)), (name, email) => $"[[{name}]]");
        }

        private static string[] Keys(ObsidianSettings settings)
        {
            Dictionary<string, object> metadata = NewEngine(settings).BuildMetadata(
                "T", "Sara Adams", "sara@example.com", new List<string>(), new List<string>(), "thread-1",
                Received, "<msg-1@example.com>", "ENTRY1", new List<string>(), new List<string>(), false, null);
            return metadata.Keys.ToArray();
        }

        // --- Defaults ---

        [Fact]
        public void Defaults_KeepEveryPropertyAndTheIsoFilenameDate()
        {
            ObsidianSettings settings = new ObsidianSettings();

            Assert.True(settings.IncludeEmailIdProperties);
            Assert.True(settings.IncludeThreadIdProperty);
            Assert.Equal("yyyy-MM-dd", settings.EmailFilenameDateFormat);
        }

        [Fact]
        public void FromJson_OldSettingsFile_GetsTheDefaults()
        {
            ObsidianSettings settings = ObsidianSettings.FromJson("{ \"VaultName\": \"Notes\" }");

            Assert.True(settings.IncludeEmailIdProperties);
            Assert.True(settings.IncludeThreadIdProperty);
            Assert.Equal("yyyy-MM-dd", settings.EmailFilenameDateFormat);
        }

        [Fact]
        public void FromJson_BlankFilenameDateFormat_FallsBackToDefault()
        {
            ObsidianSettings settings = ObsidianSettings.FromJson("{ \"EmailFilenameDateFormat\": \"  \" }");

            Assert.Equal("yyyy-MM-dd", settings.EmailFilenameDateFormat);
        }

        [Fact]
        public void Validate_InvalidFilenameDateFormat_Throws()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameDateFormat = "yyyy-MM-dd\\";

            // Validate is private (Save calls it); reach it by reflection rather than writing to AppData.
            System.Reflection.MethodInfo validate = typeof(ObsidianSettings).GetMethod(
                "Validate", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            System.Reflection.TargetInvocationException ex = Assert.Throws<System.Reflection.TargetInvocationException>(
                () => validate.Invoke(settings, null));
            Assert.IsType<ArgumentException>(ex.InnerException);
            Assert.Contains("Email filename date format", ex.InnerException.Message);
        }

        [Fact]
        public void Settings_RoundTripThroughJson()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameDateFormat = "MM-dd-yyyy";
            settings.IncludeEmailIdProperties = false;
            settings.IncludeThreadIdProperty = false;

            ObsidianSettings copy = ObsidianSettings.FromJson(settings.ToJson());

            Assert.Equal("MM-dd-yyyy", copy.EmailFilenameDateFormat);
            Assert.False(copy.IncludeEmailIdProperties);
            Assert.False(copy.IncludeThreadIdProperty);
        }

        // --- Filename ---

        [Fact]
        public void BuildBaseFileName_SenderFullToken_UsesTheDisplayName()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameFormat = "{Subject} - {SenderFull} - {Date}";

            string name = NewEngine(settings).BuildBaseFileName(Received, "Budget review", "SaraA", "2026-09-23-1818", false, "Sara Adams");

            Assert.Equal("Budget review - Sara Adams - 2026-09-23", name);
        }

        [Fact]
        public void BuildBaseFileName_SteveOFormat_MonthDayYear()
        {
            // The fork's hard-coded name, now a format plus a date format.
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameFormat = "{Subject} - {SenderFull} - {Date}";
            settings.EmailFilenameDateFormat = "MM-dd-yyyy";

            string name = NewEngine(settings).BuildBaseFileName(Received, "Budget review", "SaraA", "2026-09-23-1818", false, "Sara Adams");

            Assert.Equal("Budget review - Sara Adams - 09-23-2026", name);
        }

        [Fact]
        public void BuildBaseFileName_SenderFullUnknown_FallsBackToShortName()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameFormat = "{Subject} - {SenderFull}";

            Assert.Equal("Budget review - SaraA", NewEngine(settings).BuildBaseFileName(Received, "Budget review", "SaraA", "t", false, null));
            Assert.Equal("Budget review - SaraA", NewEngine(settings).BuildBaseFileName(Received, "Budget review", "SaraA", "t", false, "  "));
        }

        [Fact]
        public void BuildBaseFileName_SenderFullWithInvalidCharacters_IsSanitised()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameFormat = "{Subject} - {SenderFull}";

            string name = NewEngine(settings).BuildBaseFileName(Received, "Budget", "SaraA", "t", false, "Adams, Sara: \"Finance\" <x>");

            Assert.Equal("Budget - Adams, Sara_ _Finance_ _x", name);
        }

        [Fact]
        public void BuildBaseFileName_FilenameDateWithSlashes_IsSanitised()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameFormat = "{Date} {Subject}";
            settings.EmailFilenameDateFormat = "MM/dd/yyyy";

            Assert.Equal("09_23_2026 Budget", NewEngine(settings).BuildBaseFileName(Received, "Budget", "SaraA", "t", false, null));
        }

        [Fact]
        public void BuildBaseFileName_FilenameDateFormatDoesNotChangeTheLegacyName()
        {
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameDateFormat = "MM-dd-yyyy";

            Assert.Equal("Budget review-SaraA-2026-09-23-1818",
                NewEngine(settings).BuildBaseFileName(Received, "Budget review", "SaraA", "2026-09-23-1818", false, "Sara Adams"));
        }

        // --- Properties ---

        [Fact]
        public void BuildMetadata_Defaults_SameKeysInTheSameOrder()
        {
            Assert.Equal(
                new[] { "title", "type", "from", "fromEmail", "to", "toEmail", "threadId", "date", "internetMessageId", "entryId" },
                Keys(NewSettings()));
        }

        [Fact]
        public void BuildMetadata_EmailIdsOff_DropsBothIds()
        {
            ObsidianSettings settings = NewSettings();
            settings.IncludeEmailIdProperties = false;

            Assert.Equal(new[] { "title", "type", "from", "fromEmail", "to", "toEmail", "threadId", "date" }, Keys(settings));
        }

        [Fact]
        public void BuildMetadata_ThreadIdOffWhileGrouping_KeepsThreadId()
        {
            ObsidianSettings settings = NewSettings();
            settings.GroupEmailThreads = true;
            settings.IncludeThreadIdProperty = false;

            Assert.Contains("threadId", Keys(settings));
        }

        [Fact]
        public void BuildMetadata_ThreadIdOffAndNotGrouping_DropsThreadId()
        {
            ObsidianSettings settings = NewSettings();
            settings.GroupEmailThreads = false;
            settings.IncludeThreadIdProperty = false;

            Assert.Equal(new[] { "title", "type", "from", "fromEmail", "to", "toEmail", "date", "internetMessageId", "entryId" }, Keys(settings));
        }

        [Fact]
        public void BuildMetadata_AllTrimmed_LaterKeysStayInOrder()
        {
            // The fork's trimmed front matter. Optional keys added afterwards must not move up into
            // the slots of the dropped ones.
            ObsidianSettings settings = NewSettings();
            settings.GroupEmailThreads = false;
            settings.IncludeThreadIdProperty = false;
            settings.IncludeEmailIdProperties = false;
            settings.IncludeDailyNoteLink = true;
            settings.DefaultNoteTags = new List<string> { "email" };

            Dictionary<string, object> metadata = NewEngine(settings).BuildMetadata(
                "T", "Sara Adams", "sara@example.com", new List<string>(), new List<string>(), "thread-1",
                Received, "<msg-1@example.com>", "ENTRY1", new List<string> { "[[Dana Lee]]" }, new List<string> { "dana@example.com" }, false, null);

            Assert.Equal(
                new[] { "title", "type", "from", "fromEmail", "to", "toEmail", "date", "dailyNoteLink", "tags", "cc", "ccEmail" },
                metadata.Keys.ToArray());
        }

        [Fact]
        public void ProcessedEmailCache_IdsOff_NoteGivesNothingToFindAfterARestart()
        {
            // Documents the cost the setting's help text names: the restart-proof duplicate check
            // reads internetMessageId/entryId from the notes.
            ObsidianSettings settings = NewSettings();
            settings.IncludeEmailIdProperties = false;
            Dictionary<string, object> metadata = NewEngine(settings).BuildMetadata(
                "T", "Sara Adams", "sara@example.com", new List<string>(), new List<string>(), "thread-1",
                Received, "<msg-1@example.com>", "ENTRY1", new List<string>(), new List<string>(), false, null);
            string inbox = Path.Combine(_testDir, "Vault", "Inbox");
            Directory.CreateDirectory(inbox);
            File.WriteAllText(Path.Combine(inbox, "T.md"), new TemplateService(new FileService(settings)).BuildFrontMatter(metadata));

            Assert.False(new ProcessedEmailCache().IsDuplicate(inbox, "<msg-1@example.com>", "ENTRY1", Received));
        }
    }
}
