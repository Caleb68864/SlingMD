using System;
using System.IO;
using System.Text;
using Xunit;
using Moq;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Models;
using System.Collections.Generic;

namespace SlingMD.Tests.Services
{
    public class FileServiceTests
    {
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly string _testDir;

        public FileServiceTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "FileService");
            
            // Clean up any previous test data
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);

            // Create settings
            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                InboxFolder = "Inbox",
                ContactsFolder = "Contacts",
                SubjectCleanupPatterns = new List<string> 
                { 
                    @"^(?:(?:Re|Fwd|FW|RE|FWD)[:\s_-])*",
                    @"\[EXTERNAL\]\s*"
                }
            };

            _fileService = new FileService(_settings);
        }

        [Fact]
        public void GetSettings_ReturnsCorrectSettings()
        {
            // Act
            var settings = _fileService.GetSettings();

            // Assert
            Assert.Same(_settings, settings);
        }

        [Fact]
        public void CleanFileName_InvalidCleanupPattern_DoesNotThrow_AndStillCleans()
        {
            // An invalid regex could reach SubjectCleanupPatterns via hand-edited JSON (Load does not
            // validate). CleanFileName must skip the bad pattern rather than abort the export.
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                SubjectCleanupPatterns = new List<string> { "[unclosed", @"\[EXTERNAL\]\s*" }
            };
            FileService fileService = new FileService(settings);

            string result = fileService.CleanFileName("[EXTERNAL] Hello");

            // The valid pattern still applied; no exception from the invalid one.
            Assert.Equal("Hello", result);
        }

        [Fact]
        public void CleanFileName_NullCleanupPatterns_DoesNotThrow()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                SubjectCleanupPatterns = null
            };
            FileService fileService = new FileService(settings);

            string result = fileService.CleanFileName("Plain subject");

            Assert.Equal("Plain subject", result);
        }

        [Fact]
        public void WriteUtf8File_CreatesDirectoryAndFile()
        {
            // Arrange
            string testFilePath = Path.Combine(_testDir, "TestFolder", "test.txt");
            string testContent = "Hello, world!";

            // Act
            _fileService.WriteUtf8File(testFilePath, testContent);

            // Assert
            Assert.True(Directory.Exists(Path.Combine(_testDir, "TestFolder")));
            Assert.True(File.Exists(testFilePath));
            
            string fileContent = File.ReadAllText(testFilePath);
            Assert.Equal(testContent, fileContent);
        }

        [Fact]
        public void CleanFileName_RemovesInvalidCharacters()
        {
            // Arrange
            string dirtyName = "Test: File* Name? with <> invalid | chars";

            // Act
            string cleanName = _fileService.CleanFileName(dirtyName);

            // Assert
            Assert.DoesNotContain(":", cleanName);
            Assert.DoesNotContain("*", cleanName);
            Assert.DoesNotContain("?", cleanName);
            Assert.DoesNotContain("<", cleanName);
            Assert.DoesNotContain(">", cleanName);
            Assert.DoesNotContain("|", cleanName);
        }

        [Fact]
        public void CleanFileName_AppliesCleanupPatterns()
        {
            // Arrange
            string dirtyName = "Re: [EXTERNAL] This is a test email";

            // Act
            string cleanName = _fileService.CleanFileName(dirtyName);

            // Assert
            Assert.Equal("This is a test email", cleanName);
        }

        [Fact]
        public void EnsureDirectoryExists_CreatesDirectoryIfNotExists()
        {
            // Arrange
            string testPath = Path.Combine(_testDir, "NewTestDir");
            if (Directory.Exists(testPath))
            {
                Directory.Delete(testPath);
            }

            // Act
            bool result = _fileService.EnsureDirectoryExists(testPath);

            // Assert
            Assert.True(result);
            Assert.True(Directory.Exists(testPath));
        }

        [Fact]
        public void EnsureDirectoryExists_ReturnsTrueIfDirectoryExists()
        {
            // Arrange
            string testPath = Path.Combine(_testDir, "ExistingDir");
            Directory.CreateDirectory(testPath);

            // Act
            bool result = _fileService.EnsureDirectoryExists(testPath);

            // Assert
            Assert.True(result);
            Assert.True(Directory.Exists(testPath));
        }

        [Fact]
        public void AppendToFile_AppendsToExistingNoteWithoutBomOrTempFile()
        {
            string notePath = Path.Combine(_testDir, "Append", "Note.md");
            _fileService.WriteUtf8File(notePath, "---\ntitle: \"Café\"\n---\nBody");

            _fileService.AppendToFile(notePath, "\n\n## Attachments\n");

            byte[] bytes = File.ReadAllBytes(notePath);
            Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF);
            Assert.Equal("---\ntitle: \"Café\"\n---\nBody\n\n## Attachments\n", File.ReadAllText(notePath, Encoding.UTF8));
            Assert.False(File.Exists(notePath + ".slingtmp"));
        }

        [Fact]
        public void AppendToFile_CreatesNoteWhenMissing()
        {
            string notePath = Path.Combine(_testDir, "AppendNew", "Note.md");

            _fileService.AppendToFile(notePath, "content");

            Assert.Equal("content", File.ReadAllText(notePath));
        }

        [Fact]
        public void EnsureDirectoryExists_ReturnsFalseForInvalidPath()
        {
            // A path with an illegal character can't be created; the failure is logged, not thrown.
            // (Concatenated, because Path.Combine itself rejects the character on .NET Framework.)
            string testPath = _testDir + Path.DirectorySeparatorChar + "bad\0name";

            bool result = _fileService.EnsureDirectoryExists(testPath);

            Assert.False(result);
        }

        [Fact]
        public void BuildObsidianOpenUri_AbsolutePathInsideVault_IsMadeVaultRelative()
        {
            string vaultRoot = Path.Combine(Path.GetTempPath(), "Notes", "My Vault");
            string uri = FileService.BuildObsidianOpenUri("My Vault", Path.Combine(vaultRoot, "Contacts"), vaultRoot);

            Assert.Equal("obsidian://open?vault=My%20Vault&file=Contacts", uri);
        }

        [Fact]
        public void BuildObsidianOpenUri_RelativeNotePath_DropsMdExtensionAndUsesForwardSlashes()
        {
            string uri = FileService.BuildObsidianOpenUri("Vault", @"Inbox\Thread\Note.MD", Path.Combine(Path.GetTempPath(), "Notes", "Vault"));

            Assert.Equal("obsidian://open?vault=Vault&file=Inbox%2FThread%2FNote", uri);
        }

        [Fact]
        public void BuildObsidianOpenUri_AbsolutePathOutsideVault_IsPassedThrough()
        {
            // "...\Notes\Vault2" shares a prefix with the vault root but is a different folder.
            string notesRoot = Path.Combine(Path.GetTempPath(), "Notes");
            string uri = FileService.BuildObsidianOpenUri("Vault", Path.Combine(notesRoot, "Vault2", "Note.md"), Path.Combine(notesRoot, "Vault"));

            string expectedFile = Path.Combine(notesRoot, "Vault2", "Note").Replace('\\', '/');
            Assert.Equal("obsidian://open?vault=Vault&file=" + Uri.EscapeDataString(expectedFile), uri);
        }

        [Fact]
        public void GetInboxPath_ReturnsSettingsInboxPath()
        {
            // Act
            string inboxPath = _fileService.GetInboxPath();

            // Assert
            Assert.Equal(Path.Combine(_testDir, "TestVault", "Inbox"), inboxPath);
        }
    }
}