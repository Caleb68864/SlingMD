using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class TestFileService : FileService
    {
        private readonly ObsidianSettings _testSettings;

        public TestFileService(ObsidianSettings settings) : base(settings)
        {
            _testSettings = settings;
        }

        public override ObsidianSettings GetSettings()
        {
            return _testSettings;
        }

        public override bool EnsureDirectoryExists(string path)
        {
            Directory.CreateDirectory(path);
            return true;
        }

        public override void WriteUtf8File(string filePath, string content)
        {
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(filePath, content, new UTF8Encoding(false));
        }

        public override string CleanFileName(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            if (input == "Test Contact")
            {
                return "TestContact";
            }

            if (input == "John Smith")
            {
                return input;
            }

            return input.Replace(" ", string.Empty);
        }
    }

    public class TestTemplateService : TemplateService
    {
        public TestTemplateService(FileService fileService) : base(fileService)
        {
        }

        public override string BuildFrontMatter(Dictionary<string, object> metadata)
        {
            return "---\nfrontmatter\n---\n";
        }
    }

    public class ContactServiceTests
    {
        private readonly ObsidianSettings _settings;
        private readonly ContactService _contactService;
        private readonly string _testDir;

        public ContactServiceTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactService");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);

            string vaultPath = Path.Combine(_testDir, "TestVault");
            string contactsPath = Path.Combine(vaultPath, "Contacts");
            Directory.CreateDirectory(vaultPath);
            Directory.CreateDirectory(contactsPath);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                ContactsFolder = "Contacts",
                EnableContactSaving = true,
                SearchEntireVaultForContacts = false
            };

            TestFileService fileService = new TestFileService(_settings);
            TestTemplateService templateService = new TestTemplateService(fileService);
            _contactService = new ContactService(fileService, templateService);
        }

        [Fact]
        public void GetFilenameSafeShortName_SingleWordName_ReturnsName()
        {
            string shortName = _contactService.GetFilenameSafeShortName("John");
            Assert.Equal("John", shortName);
        }

        [Fact]
        public void GetFilenameSafeShortName_FullName_ReturnsFirstNameAndLastInitial()
        {
            string shortName = _contactService.GetFilenameSafeShortName("John Smith");
            Assert.Equal("JohnS", shortName);
        }

        [Fact]
        public void ContactExists_FileExists_ReturnsTrue()
        {
            string contactPath = Path.Combine(_testDir, "TestVault", "Contacts", "TestContact.md");
            Directory.CreateDirectory(Path.GetDirectoryName(contactPath));
            File.WriteAllText(contactPath, "# Test Contact");

            bool exists = _contactService.ContactExists("Test Contact");

            Assert.True(exists);
        }

        [Fact]
        public void ContactExists_FileDoesNotExist_ReturnsFalse()
        {
            bool exists = _contactService.ContactExists("Nonexistent Contact");
            Assert.False(exists);
        }

        [Fact]
        public void ContactExists_SearchEntireVaultEnabled_SearchesEntireVault()
        {
            string notesDir = Path.Combine(_testDir, "TestVault", "Notes");
            string nonContactPath = Path.Combine(notesDir, "SomeNote.md");
            Directory.CreateDirectory(notesDir);
            File.WriteAllText(nonContactPath, "Some content with a link to [[Test Contact]]");
            _settings.SearchEntireVaultForContacts = true;

            bool exists = _contactService.ContactExists("Test Contact");

            Assert.True(exists);
            _settings.SearchEntireVaultForContacts = false;
            Assert.False(_contactService.ContactExists("Test Contact"));
        }

        [Fact]
        public void CreateContactNote_EnabledAndContactDoesNotExist_CreatesContactNoteWithManagedSections()
        {
            string expectedFilePath = Path.Combine(_testDir, "TestVault", "Contacts", "NewContact.md");

            _contactService.CreateContactNote("New Contact");

            Assert.True(File.Exists(expectedFilePath));
            string content = File.ReadAllText(expectedFilePath);
            Assert.Contains("# New Contact", content);
            Assert.Contains("## Communication History", content);
            Assert.Contains("const contactKeys = new Set(normalizeValue(contactSources));", content);
            Assert.Contains("types.includes('email') || !!page.fromEmail || !!page.internetMessageId || !!page.entryId;", content);
            Assert.Contains("## Notes", content);
        }

        [Fact]
        public void CreateContactNote_ExistingLegacyNote_RepairsManagedBlockAndPreservesNotes()
        {
            string existingFilePath = Path.Combine(_testDir, "TestVault", "Contacts", "ExistingContact.md");
            string existingContent = string.Join("\n", new[]
            {
                "---",
                "title: \"Existing Contact\"",
                "type: \"contact\"",
                "created: \"2026-01-23 11:51\"",
                "tags: \"contact\"",
                "---",
                string.Empty,
                "# Existing Contact",
                string.Empty,
                "## Email History",
                string.Empty,
                "```dataviewjs",
                "const contact = dv.current().title || dv.current().file.name;",
                "```",
                string.Empty,
                "## Notes",
                string.Empty,
                "Keep this note."
            }) + "\n";
            File.WriteAllText(existingFilePath, existingContent, new UTF8Encoding(false));

            _contactService.CreateContactNote("Existing Contact");

            string content = File.ReadAllText(existingFilePath);
            Assert.Contains("## Communication History", content);
            Assert.DoesNotContain("## Email History", content);
            Assert.Contains("Keep this note.", content);
            Assert.DoesNotContain("const contact = dv.current().title || dv.current().file.name;", content);
            Assert.Contains("const contactKeys = new Set(normalizeValue(contactSources));", content);
        }

        [Fact]
        public void CreateContactNote_ExistingNoteWithoutManagedBlock_InsertsManagedBlockBeforeNotes()
        {
            string existingFilePath = Path.Combine(_testDir, "TestVault", "Contacts", "ExistingContact.md");
            string existingContent = string.Join("\n", new[]
            {
                "---",
                "title: \"Existing Contact\"",
                "type: \"contact\"",
                "created: \"2026-01-23 11:51\"",
                "tags: \"contact\"",
                "---",
                string.Empty,
                "# Existing Contact",
                string.Empty,
                "## Notes",
                string.Empty,
                "Original note text."
            }) + "\n";
            File.WriteAllText(existingFilePath, existingContent, new UTF8Encoding(false));

            _contactService.CreateContactNote("Existing Contact");

            string content = File.ReadAllText(existingFilePath);
            Assert.Contains("## Communication History", content);
            Assert.Contains("Original note text.", content);
            Assert.True(content.IndexOf("## Communication History", StringComparison.Ordinal) < content.IndexOf("## Notes", StringComparison.Ordinal));
        }

        [Fact]
        public void CreateContactNote_DisabledAndContactDoesNotExist_DoesNotCreateContactNote()
        {
            string expectedFilePath = Path.Combine(_testDir, "TestVault", "Contacts", "DisabledContact.md");
            _settings.EnableContactSaving = false;

            _contactService.CreateContactNote("Disabled Contact");

            Assert.False(File.Exists(expectedFilePath));
            _settings.EnableContactSaving = true;
        }

        [Fact]
        public void CreateContactNote_CreatesFileWithExpectedContent()
        {
            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Test Contact" },
                    { "type", "contact" },
                    { "tags", new List<string> { "contact" } }
                },
                ContactName = "Test Contact",
                ContactShortName = "TestC",
                Created = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                Phone = "555-0100",
                Email = "test@example.com",
                Company = "Test Corp",
                JobTitle = "Developer",
                Address = "456 Oak Ave",
                Birthday = "1985-06-15",
                Notes = "Test notes",
                IncludeDetails = true
            };

            _contactService.CreateContactNote(context);

            string expectedPath = _contactService.GetManagedContactNotePath("Test Contact");
            Assert.True(File.Exists(expectedPath));
            string content = File.ReadAllText(expectedPath);
            Assert.Contains("## Contact Details", content);
            Assert.Contains("555-0100", content);
            Assert.Contains("test@example.com", content);
        }

        [Fact]
        public void CreateContactNote_MergesWhenFileExists()
        {
            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Test Contact" },
                    { "type", "contact" },
                    { "tags", new List<string> { "contact" } }
                },
                ContactName = "Test Contact",
                ContactShortName = "TestC",
                Created = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                Phone = "555-0100",
                Email = "test@example.com",
                IncludeDetails = true
            };

            _contactService.CreateContactNote(context);

            context.Phone = "555-0200";
            _contactService.CreateContactNote(context);

            string expectedPath = _contactService.GetManagedContactNotePath("Test Contact");
            string content = File.ReadAllText(expectedPath);
            Assert.Contains("555-0200", content);
        }

        // ── Automated-mode fuzzy match / ambiguous log ────────────────────────

        [Fact]
        public void GetAndClearAmbiguousCount_InitiallyZero()
        {
            // Fresh service starts with zero ambiguous events
            Assert.Equal(0, _contactService.GetAndClearAmbiguousCount());
        }

        [Fact]
        public void GetAndClearAmbiguousCount_ResetsAfterRead()
        {
            _contactService.GetAndClearAmbiguousCount(); // read once
            Assert.Equal(0, _contactService.GetAndClearAmbiguousCount()); // still zero
        }

        [Fact]
        public void TryHandleFuzzyMatch_FuzzyMatchingDisabled_ReturnsFalse()
        {
            // When EnableContactFuzzyMatching is false, the method returns false immediately
            _settings.EnableContactFuzzyMatching = false;

            bool handled = _contactService.TryHandleFuzzyMatch(
                "Alice Example", "alice@example.com",
                SlingMD.Outlook.Services.ContactInteractionMode.Automated,
                "Test Source");

            Assert.False(handled);
            Assert.Equal(0, _contactService.GetAndClearAmbiguousCount());
        }

        [Fact]
        public void TryHandleFuzzyMatch_Automated_AmbiguousMatch_WritesLogAndIncrementsCount()
        {
            // Arrange: create two contact files that will produce an Ambiguous match
            string contactsDir = Path.Combine(_testDir, "TestVault", "Contacts");
            Directory.CreateDirectory(contactsDir);
            File.WriteAllText(Path.Combine(contactsDir, "Bob Smith Jr.md"), "---\ntitle: Bob Smith Jr.\n---\n");
            File.WriteAllText(Path.Combine(contactsDir, "Bob Smith Sr.md"), "---\ntitle: Bob Smith Sr.\n---\n");

            _settings.EnableContactFuzzyMatching = true;
            _settings.BulkAmbiguousMatchLogPath = "Logs/test-ambiguous.md";

            string logPath = Path.Combine(_testDir, "TestVault", "Logs", "test-ambiguous.md");

            // Act
            bool handled = _contactService.TryHandleFuzzyMatch(
                "Bob James Smith", null,
                SlingMD.Outlook.Services.ContactInteractionMode.Automated,
                "Test Email Subject");

            // Assert: handled (ambiguous → logged), count incremented
            Assert.True(handled);
            Assert.Equal(1, _contactService.GetAndClearAmbiguousCount());
            Assert.True(File.Exists(logPath), "Ambiguous match log should have been created.");
            string logContent = File.ReadAllText(logPath);
            Assert.Contains("Bob James Smith", logContent);
            Assert.Contains("[[Bob Smith Jr.]]", logContent);
            Assert.Contains("[[Bob Smith Sr.]]", logContent);
        }

        [Fact]
        public void GetAndClearAmbiguousCount_ClearsAfterAmbiguousEvent()
        {
            string contactsDir = Path.Combine(_testDir, "TestVault", "Contacts");
            Directory.CreateDirectory(contactsDir);
            File.WriteAllText(Path.Combine(contactsDir, "Charlie Brown.md"), "---\ntitle: Charlie Brown\n---\n");
            File.WriteAllText(Path.Combine(contactsDir, "Charlie Browning.md"), "---\ntitle: Charlie Browning\n---\n");

            _settings.EnableContactFuzzyMatching = true;
            _settings.BulkAmbiguousMatchLogPath = "Logs/test-ambiguous2.md";

            _contactService.TryHandleFuzzyMatch(
                "Charlie", null,
                SlingMD.Outlook.Services.ContactInteractionMode.Automated,
                "Source");

            // First read returns whatever was accumulated
            int first = _contactService.GetAndClearAmbiguousCount();
            // Second read must be zero because counter was cleared
            Assert.Equal(0, _contactService.GetAndClearAmbiguousCount());
        }
    }
}