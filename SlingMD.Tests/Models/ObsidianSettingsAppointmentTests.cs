using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using Xunit;

namespace SlingMD.Tests.Models
{
    public class ObsidianSettingsAppointmentTests : IDisposable
    {
        private readonly string _testDir;

        public ObsidianSettingsAppointmentTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "SettingsAppointment");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);
        }

        [Fact]
        public void DefaultValues_AllAppointmentProperties_HaveCorrectDefaults()
        {
            // Arrange & Act
            ObsidianSettings settings = new ObsidianSettings();

            // Assert
            Assert.Equal("Appointments", settings.AppointmentsFolder);
            Assert.Equal("{Date} - {Subject}", settings.AppointmentNoteTitleFormat);
            Assert.Equal(50, settings.AppointmentNoteTitleMaxLength);
            Assert.Single(settings.AppointmentDefaultNoteTags);
            Assert.Contains("Appointment", settings.AppointmentDefaultNoteTags);
            Assert.True(settings.AppointmentSaveAttachments);
            Assert.True(settings.CreateMeetingNotes);
            Assert.Equal(string.Empty, settings.MeetingNoteTemplate);
            Assert.True(settings.GroupRecurringMeetings);
            Assert.False(settings.SaveCancelledAppointments);
            Assert.Equal("None", settings.AppointmentTaskCreation);
        }

        [Fact]
        public void GetAppointmentsPath_ReturnsCorrectCombinedPath()
        {
            // Arrange
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Notes",
                VaultName = "MyVault",
                AppointmentsFolder = "Appointments"
            };

            // Act
            string result = settings.GetAppointmentsPath();

            // Assert
            Assert.Equal(Path.Combine(@"C:\Notes", "MyVault", "Appointments"), result);
        }

        [Fact]
        public void RoundTrip_AppointmentSettings_PreservedThroughSaveLoad()
        {
            // Arrange - use testable subclass
            string settingsPath = Path.Combine(_testDir, "settings.json");
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = settingsPath,
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentsFolder = "CustomAppointments",
                AppointmentNoteTitleFormat = "{Subject} on {Date}",
                AppointmentNoteTitleMaxLength = 75,
                AppointmentDefaultNoteTags = new List<string> { "Meeting", "Work" },
                AppointmentSaveAttachments = false,
                CreateMeetingNotes = false,
                MeetingNoteTemplate = "custom-template.md",
                GroupRecurringMeetings = false,
                SaveCancelledAppointments = true,
                AppointmentTaskCreation = "Both"
            };

            // Act
            settings.Save();
            ObsidianSettingsTestable loaded = new ObsidianSettingsTestable
            {
                TestSettingsPath = settingsPath
            };
            loaded.Load();

            // Assert
            Assert.Equal("CustomAppointments", loaded.AppointmentsFolder);
            Assert.Equal("{Subject} on {Date}", loaded.AppointmentNoteTitleFormat);
            Assert.Equal(75, loaded.AppointmentNoteTitleMaxLength);
            Assert.Equal(2, loaded.AppointmentDefaultNoteTags.Count);
            Assert.Contains("Meeting", loaded.AppointmentDefaultNoteTags);
            Assert.False(loaded.AppointmentSaveAttachments);
            Assert.False(loaded.CreateMeetingNotes);
            Assert.Equal("custom-template.md", loaded.MeetingNoteTemplate);
            Assert.False(loaded.GroupRecurringMeetings);
            Assert.True(loaded.SaveCancelledAppointments);
            Assert.Equal("Both", loaded.AppointmentTaskCreation);
        }

        [Fact]
        public void Save_AppointmentNoteTitleMaxLength_OutOfRange_Throws()
        {
            // Arrange - use testable subclass so Save() writes to temp path
            string settingsPath = Path.Combine(_testDir, "invalid-settings.json");
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = settingsPath,
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentNoteTitleMaxLength = 5  // Below minimum of 10
            };

            // Act & Assert - Validate is called by Save
            Assert.Throws<System.ArgumentException>(() => settings.Save());
        }

        [Fact]
        public void Save_AppointmentTaskCreation_InvalidValue_Throws()
        {
            // Arrange - use testable subclass so Save() writes to temp path
            string settingsPath = Path.Combine(_testDir, "invalid-settings2.json");
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = settingsPath,
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentTaskCreation = "Invalid"
            };

            // Act & Assert - Validate is called by Save
            Assert.Throws<System.ArgumentException>(() => settings.Save());
        }

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
