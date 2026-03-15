using System;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class ContactProcessorTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;

        public ContactProcessorTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactProcessor");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                ContactsFolder = "Contacts",
                EnableContactSaving = true,
                ContactNoteIncludeDetails = true
            };
        }

        [Fact]
        public void Constructor_WithValidSettings_CreatesInstance()
        {
            ContactProcessor processor = new ContactProcessor(_settings);
            Assert.NotNull(processor);
        }

        [Fact]
        public void Constructor_WithNullSettings_DoesNotThrow()
        {
            System.Exception caughtException = null;
            ContactProcessor processor = null;
            try
            {
                processor = new ContactProcessor(null);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
            Assert.NotNull(processor);
        }

        [Fact]
        public void GetBulkErrors_InitiallyEmpty()
        {
            ContactProcessor processor = new ContactProcessor(_settings);
            System.Collections.Generic.List<string> errors = processor.GetBulkErrors();
            Assert.Empty(errors);
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
            catch { }
        }
    }
}
