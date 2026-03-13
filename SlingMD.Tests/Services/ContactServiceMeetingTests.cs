using System;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class ContactServiceMeetingTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly ContactService _contactService;

        public ContactServiceMeetingTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactMeeting");
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
                EnableContactSaving = true
            };
            TestFileService fileService = new TestFileService(_settings);
            TestTemplateService templateService = new TestTemplateService(fileService);
            _contactService = new ContactService(fileService, templateService);
        }

        [Fact]
        public void GetShortName_WithFullName_ReturnsFirstNameAndLastInitial()
        {
            // Arrange & Act (existing method, verify still works)
            string result = _contactService.GetShortName("John Smith");

            // Assert - GetShortName returns first name + last initial when two parts
            Assert.Equal("JohnS", result);
        }

        // Note: GetSMTPEmailAddress, BuildLinkedNames (meeting overload),
        // BuildEmailList (meeting overload), and GetMeetingResourceData
        // all take COM objects (Recipient, Recipients) that require
        // running Outlook. These methods should be verified via:
        // 1. Build compilation (structural verification)
        // 2. Manual testing with Outlook running
        // 3. Integration tests if COM interop mocking is feasible

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                try { Directory.Delete(_testDir, true); }
                catch (System.Exception) { }
            }
        }
    }
}
