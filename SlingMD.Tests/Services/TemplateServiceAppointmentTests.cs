using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class TemplateServiceAppointmentTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;

        public TemplateServiceAppointmentTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "TemplateAppointment");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
            Directory.CreateDirectory(_testDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "TestVault"
            };
            _fileService = new FileService(_settings);
            _templateService = new TemplateService(_fileService);
        }

        [Fact]
        public void AppointmentTemplateContext_AllProperties_Settable()
        {
            // Arrange & Act
            AppointmentTemplateContext context = new AppointmentTemplateContext
            {
                Metadata = new Dictionary<string, object> { { "title", "Test" } },
                NoteTitle = "Test Meeting",
                Subject = "Weekly Standup",
                Organizer = "[[John Smith]]",
                OrganizerEmail = "john@example.com",
                Attendees = "[[Jane]], [[Bob]]",
                OptionalAttendees = "[[Alice]]",
                Resources = "Conference Room A",
                Location = "Room 101",
                StartDateTime = "2026-03-13 09:00",
                EndDateTime = "2026-03-13 10:00",
                Recurrence = "Weekly",
                Date = "2026-03-13",
                Body = "Meeting body content",
                TaskBlock = "",
                FileName = "test.md",
                FileNameWithoutExtension = "test"
            };

            // Assert
            Assert.Equal("Weekly Standup", context.Subject);
            Assert.Equal("[[John Smith]]", context.Organizer);
            Assert.Equal("Room 101", context.Location);
        }

        [Fact]
        public void MeetingNoteTemplateContext_AllProperties_Settable()
        {
            // Arrange & Act
            MeetingNoteTemplateContext context = new MeetingNoteTemplateContext
            {
                Metadata = new Dictionary<string, object> { { "title", "Notes" } },
                AppointmentTitle = "Weekly Standup",
                AppointmentLink = "[[2026-03-13 - Weekly Standup]]",
                Organizer = "[[John Smith]]",
                Attendees = "[[Jane]], [[Bob]]",
                Date = "2026-03-13",
                Location = "Room 101"
            };

            // Assert
            Assert.Equal("[[2026-03-13 - Weekly Standup]]", context.AppointmentLink);
            Assert.Equal("Room 101", context.Location);
        }

        [Fact]
        public void RenderAppointmentContent_WithValidContext_ProducesExpectedMarkdown()
        {
            // Arrange
            AppointmentTemplateContext context = new AppointmentTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Weekly Standup" },
                    { "type", "Appointment" }
                },
                NoteTitle = "2026-03-13 - Weekly Standup",
                Subject = "Weekly Standup",
                Organizer = "[[John Smith]]",
                Attendees = "[[Jane Doe]]",
                OptionalAttendees = "",
                Resources = "",
                Location = "Room 101",
                StartDateTime = "2026-03-13 09:00",
                EndDateTime = "2026-03-13 10:00",
                Recurrence = "Weekly",
                Body = "Standup meeting content",
                TaskBlock = "",
                Date = "2026-03-13"
            };

            // Act
            string result = _templateService.RenderAppointmentContent(context);

            // Assert
            Assert.Contains("Weekly Standup", result);
            Assert.Contains("[[John Smith]]", result);
            Assert.Contains("[[Jane Doe]]", result);
            Assert.Contains("Room 101", result);
            Assert.Contains("Standup meeting content", result);
        }

        [Fact]
        public void RenderMeetingNoteContent_WithValidContext_ProducesStubWithBacklink()
        {
            // Arrange
            MeetingNoteTemplateContext context = new MeetingNoteTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Weekly Standup - Meeting Notes" },
                    { "type", "Meeting Notes" }
                },
                AppointmentTitle = "Weekly Standup",
                AppointmentLink = "[[2026-03-13 - Weekly Standup]]",
                Organizer = "[[John Smith]]",
                Attendees = "[[Jane Doe]]",
                Date = "2026-03-13",
                Location = "Room 101"
            };

            // Act
            string result = _templateService.RenderMeetingNoteContent(context);

            // Assert
            Assert.Contains("[[2026-03-13 - Weekly Standup]]", result);
            Assert.Contains("[[John Smith]]", result);
            Assert.Contains("Action Items", result);
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
