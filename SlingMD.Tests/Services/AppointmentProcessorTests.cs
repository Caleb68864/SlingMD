using System;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class AppointmentProcessorTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;

        public AppointmentProcessorTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "AppointmentProcessor");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault",
                AppointmentsFolder = "Appointments",
                CreateMeetingNotes = true,
                GroupRecurringMeetings = true,
                SaveCancelledAppointments = false,
                AppointmentTaskCreation = "None"
            };
        }

        [Fact]
        public void Constructor_WithValidSettings_CreatesInstance()
        {
            // Arrange & Act
            System.Exception caughtException = null;
            AppointmentProcessor processor = null;
            try
            {
                processor = new AppointmentProcessor(_settings);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            // Assert
            Assert.Null(caughtException);
            Assert.NotNull(processor);
        }

        [Fact]
        public void Constructor_WithNullSettings_DoesNotThrow()
        {
            // Arrange & Act
            // All internal service constructors accept null settings (store-only pattern).
            // Passing null is structural -- exceptions only surface when methods are called.
            System.Exception caughtException = null;
            AppointmentProcessor processor = null;
            try
            {
                processor = new AppointmentProcessor(null);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            // Assert
            Assert.Null(caughtException);
            Assert.NotNull(processor);
        }

        // Note: ProcessAppointment tests require AppointmentItem COM objects
        // which need Outlook running. The following are structural tests
        // that verify the processor was wired correctly.

        [Fact]
        public void Constructor_CreatesWithDifferentSettings()
        {
            // Arrange
            ObsidianSettings settings1 = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault1"
            };
            ObsidianSettings settings2 = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault2"
            };

            // Act
            AppointmentProcessor first = new AppointmentProcessor(settings1);
            AppointmentProcessor second = new AppointmentProcessor(settings2);

            // Assert
            Assert.NotNull(first);
            Assert.NotNull(second);
            Assert.NotSame(first, second);
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
