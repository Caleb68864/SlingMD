using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Moq;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Drives <see cref="AppointmentProcessor.ProcessAppointment"/> end to end with a Moq
    /// <see cref="AppointmentItem"/> (bulk mode, so no dialogs) and pins the exact note files it
    /// writes. The double has no organizer and no recipients: the processor releases those with
    /// <c>Marshal.ReleaseComObject</c>, which throws on a non-COM object. Organizer and attendee
    /// output is pinned by <c>AppointmentNoteEngineTests</c> instead.
    /// <para>
    /// Every member the processor reads must be set up (or otherwise referenced) here. SlingMD.Tests
    /// embeds the Outlook interop types, and the embedded <c>_AppointmentItem</c> only carries the
    /// members this assembly uses, so on .NET Framework a call to any other member of the Moq proxy
    /// throws <see cref="MissingMethodException"/> (Mono does not). That is how
    /// <c>GetRecurrencePattern()</c> failed the Windows CI run before the processor stopped calling it.
    /// </para>
    /// </summary>
    public class AppointmentProcessorExportTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;

        public AppointmentProcessorExportTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "AppointmentExport", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(_testDir, "Vault"));
            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                AppointmentsFolder = "Appointments",
                TemplatesFolder = "Templates",
                CreateMeetingNotes = true,
                GroupRecurringMeetings = true,
                SaveCancelledAppointments = false,
                AppointmentTaskCreation = "None",
                EnableContactSaving = false,
                IncludeDailyNoteLink = true,
                AppointmentSaveAttachments = false,
                AppointmentDefaultNoteTags = new List<string> { "Meeting" },
                LaunchObsidian = false,
                ShowNoteInExplorer = false
            };
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

        private static Mock<AppointmentItem> Appointment(string subject, OlRecurrenceState recurrence, string globalId)
        {
            Mock<AppointmentItem> appointment = new Mock<AppointmentItem>();
            appointment.Setup(a => a.MeetingStatus).Returns(OlMeetingStatus.olMeeting);
            appointment.Setup(a => a.RecurrenceState).Returns(recurrence);
            appointment.Setup(a => a.Subject).Returns(subject);
            appointment.Setup(a => a.Body).Returns("Agenda below.\nJoin: https://teams.microsoft.com/l/meetup-join/abc123\n");
            appointment.Setup(a => a.Location).Returns("Room 4");
            appointment.Setup(a => a.Start).Returns(new DateTime(2026, 10, 5, 9, 30, 0));
            appointment.Setup(a => a.End).Returns(new DateTime(2026, 10, 5, 10, 15, 0));
            appointment.Setup(a => a.GlobalAppointmentID).Returns(globalId);
            appointment.Setup(a => a.GetOrganizer()).Returns((AddressEntry)null);
            appointment.Setup(a => a.Recipients).Returns((Recipients)null);
            // No parent folder: the calendar name is empty, without relying on GetCalendarName's
            // catch of the MissingMethodException described above.
            appointment.Setup(a => a.Parent).Returns((object)null);
            return appointment;
        }

        private string AppointmentsPath(params string[] parts)
        {
            List<string> all = new List<string> { _testDir, "Vault", "Appointments" };
            all.AddRange(parts);
            return Path.Combine(all.ToArray());
        }

        // The built-in layout ends lines with Environment.NewLine; compare with \n only.
        private static string ReadNote(string path)
        {
            return File.ReadAllText(path).Replace("\r\n", "\n");
        }

        [Fact]
        public async Task ProcessAppointment_SingleMeeting_WritesNoteAndMeetingNote()
        {
            AppointmentProcessor processor = new AppointmentProcessor(_settings);
            Mock<AppointmentItem> appointment = Appointment("Budget sync", OlRecurrenceState.olApptNotRecurring, "GID-export-single");

            AppointmentProcessingResult result = await processor.ProcessAppointment(appointment.Object, bulkMode: true);

            Assert.Equal(AppointmentProcessingResult.Success, result);
            string note = ReadNote(AppointmentsPath("2026-10-05 - Budget sync.md"));
            string meetingNote = ReadNote(AppointmentsPath("2026-10-05 - Budget sync - Meeting Notes.md"));
            Assert.Equal(SingleNote, note);
            Assert.Equal(SingleMeetingNote, meetingNote);
        }

        [Fact]
        public async Task ProcessAppointment_RecurringOccurrence_WritesIntoSeriesFolder()
        {
            AppointmentProcessor processor = new AppointmentProcessor(_settings);
            Mock<AppointmentItem> appointment = Appointment("Weekly standup", OlRecurrenceState.olApptOccurrence, "GID-export-recurring");

            AppointmentProcessingResult result = await processor.ProcessAppointment(appointment.Object, bulkMode: true);

            Assert.Equal(AppointmentProcessingResult.Success, result);
            string note = ReadNote(AppointmentsPath("Weekly standup", "2026-10-05 - Weekly standup.md"));
            string summary = ReadNote(AppointmentsPath("Weekly standup", "0-Weekly standup.md"));
            Assert.Contains("threadId: \"Weekly standup\"", summary);
            Assert.True(File.Exists(AppointmentsPath("Weekly standup", "2026-10-05 - Weekly standup - Meeting Notes.md")));
            Assert.Equal(RecurringNote, note);
        }

        [Fact]
        public async Task ProcessAppointment_RecurringOccurrence_DoesNotReadTheRecurrencePattern()
        {
            // The series folder is named from the subject alone; the pattern used to be fetched and
            // thrown away (one COM round trip per instance).
            AppointmentProcessor processor = new AppointmentProcessor(_settings);
            Mock<AppointmentItem> appointment = Appointment("Weekly standup", OlRecurrenceState.olApptOccurrence, "GID-export-pattern");

            await processor.ProcessAppointment(appointment.Object, bulkMode: true);

            appointment.Verify(a => a.GetRecurrencePattern(), Times.Never());
        }

        [Fact]
        public async Task ProcessAppointment_RecurringSubjectEndingInDot_FolderHasNoTrailingDot()
        {
            AppointmentProcessor processor = new AppointmentProcessor(_settings);
            Mock<AppointmentItem> appointment = Appointment("Sync w/ Smith Jr.", OlRecurrenceState.olApptException, "GID-export-dot");

            AppointmentProcessingResult result = await processor.ProcessAppointment(appointment.Object, bulkMode: true);

            Assert.Equal(AppointmentProcessingResult.Success, result);
            string[] folders = Directory.GetDirectories(AppointmentsPath());
            Assert.Single(folders);
            Assert.False(Path.GetFileName(folders[0]).EndsWith("."), folders[0]);
            Assert.True(File.Exists(Path.Combine(folders[0], "0-" + Path.GetFileName(folders[0]) + ".md")));
        }

        [Fact]
        public async Task ProcessAppointment_ThaiCulture_GregorianFileNames()
        {
            System.Globalization.CultureInfo original = System.Globalization.CultureInfo.CurrentCulture;
            try
            {
                System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("th-TH");
                Assert.Equal("2569", new DateTime(2026, 10, 5).ToString("yyyy"));
                AppointmentProcessor processor = new AppointmentProcessor(_settings);

                await processor.ProcessAppointment(Appointment("Retro", OlRecurrenceState.olApptOccurrence, "GID-export-thai").Object, bulkMode: true);

                Assert.True(File.Exists(AppointmentsPath("Retro", "2026-10-05 - Retro.md")),
                    string.Join(", ", Directory.GetFiles(AppointmentsPath(), "*", SearchOption.AllDirectories)));
            }
            finally
            {
                System.Globalization.CultureInfo.CurrentCulture = original;
            }
        }

        [Fact]
        public async Task ProcessAppointment_SameGlobalIdTwice_SkipsTheSecond()
        {
            AppointmentProcessor processor = new AppointmentProcessor(_settings);
            string globalId = "GID-" + Guid.NewGuid().ToString("N");

            AppointmentProcessingResult first = await processor.ProcessAppointment(Appointment("Budget sync", OlRecurrenceState.olApptNotRecurring, globalId).Object, bulkMode: true);
            AppointmentProcessingResult second = await processor.ProcessAppointment(Appointment("Budget sync", OlRecurrenceState.olApptNotRecurring, globalId).Object, bulkMode: true);

            Assert.Equal(AppointmentProcessingResult.Success, first);
            Assert.Equal(AppointmentProcessingResult.Skipped, second);
        }

        private const string SingleNote = ""
            + "---\n"
            + "title: \"2026-10-05 - Budget sync\"\n"
            + "type: \"Appointment\"\n"
            + "calendar: \"\"\n"
            + "organizer: \"\"\n"
            + "organizerEmail: \"\"\n"
            + "location: \"Room 4\"\n"
            + "startDateTime: \"2026-10-05T09:30\"\n"
            + "endDateTime: \"2026-10-05T10:15\"\n"
            + "recurrence: \"None\"\n"
            + "globalAppointmentId: \"GID-export-single\"\n"
            + "meetingNotes: \"[[2026-10-05 - Budget sync - Meeting Notes]]\"\n"
            + "dailyNoteLink: \"[[2026-10-05]]\"\n"
            + "tags: \n"
            + "  - \"Meeting\"\n"
            + "meetingUrl: \"https://teams.microsoft.com/l/meetup-join/abc123\"\n"
            + "---\n"
            + "\n"
            + "# 2026-10-05 - Budget sync\n"
            + "\n"
            + "**Location:** Room 4\n"
            + "\n"
            + "**Join:** [Join the meeting](https://teams.microsoft.com/l/meetup-join/abc123)\n"
            + "\n"
            + "**Start:** 2026-10-05T09:30\n"
            + "**End:** 2026-10-05T10:15\n"
            + "\n"
            + "## Description\n"
            + "\n"
            + "Agenda below.\n"
            + "Join: https://teams.microsoft.com/l/meetup-join/abc123\n"
            + "\n"
            + "## Meeting Notes\n"
            + "\n";

        private const string SingleMeetingNote = ""
            + "---\n"
            + "title: \"2026-10-05 - Budget sync - Meeting Notes\"\n"
            + "type: \"Meeting Notes\"\n"
            + "appointment: \"[[2026-10-05 - Budget sync]]\"\n"
            + "date: \"2026-10-05\"\n"
            + "tags: \n"
            + "  - \"Meeting\"\n"
            + "---\n"
            + "\n"
            + "\n"
            + "## Meeting Notes\n"
            + "\n"
            + "**Appointment:** [[2026-10-05 - Budget sync]]\n"
            + "**Organizer:** \n"
            + "**Attendees:** \n"
            + "**Date:** 2026-10-05\n"
            + "**Location:** Room 4\n"
            + "\n"
            + "## Agenda\n"
            + "\n"
            + "-\n"
            + "\n"
            + "## Notes\n"
            + "\n"
            + "-\n"
            + "\n"
            + "## Action Items\n"
            + "\n"
            + "- [ ]";

        private const string RecurringNote = ""
            + "---\n"
            + "title: \"2026-10-05 - Weekly standup\"\n"
            + "type: \"Appointment\"\n"
            + "calendar: \"\"\n"
            + "organizer: \"\"\n"
            + "organizerEmail: \"\"\n"
            + "location: \"Room 4\"\n"
            + "startDateTime: \"2026-10-05T09:30\"\n"
            + "endDateTime: \"2026-10-05T10:15\"\n"
            + "recurrence: \"Occurrence\"\n"
            + "globalAppointmentId: \"GID-export-recurring\"\n"
            + "meetingNotes: \"[[2026-10-05 - Weekly standup - Meeting Notes]]\"\n"
            + "dailyNoteLink: \"[[2026-10-05]]\"\n"
            + "tags: \n"
            + "  - \"Meeting\"\n"
            + "meetingUrl: \"https://teams.microsoft.com/l/meetup-join/abc123\"\n"
            + "---\n"
            + "\n"
            + "# 2026-10-05 - Weekly standup\n"
            + "\n"
            + "**Location:** Room 4\n"
            + "\n"
            + "**Join:** [Join the meeting](https://teams.microsoft.com/l/meetup-join/abc123)\n"
            + "\n"
            + "**Start:** 2026-10-05T09:30\n"
            + "**End:** 2026-10-05T10:15\n"
            + "\n"
            + "## Description\n"
            + "\n"
            + "Agenda below.\n"
            + "Join: https://teams.microsoft.com/l/meetup-join/abc123\n"
            + "\n"
            + "## Meeting Notes\n"
            + "\n";
    }
}
