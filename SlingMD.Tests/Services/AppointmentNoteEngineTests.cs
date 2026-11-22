using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Characterization tests for <see cref="AppointmentNoteEngine"/>, the COM-free half of
    /// AppointmentProcessor: given an <see cref="AppointmentData"/>, the exact title, front matter
    /// and note text. The expected values are what AppointmentProcessor produced before the split
    /// (same code, moved; <c>AppointmentProcessorExportTests</c> pins the processor's output on both
    /// sides of the move). Linked into SlingMD.Core.Tests, so they run on Linux.
    /// </summary>
    public class AppointmentNoteEngineTests : IDisposable
    {
        private readonly string _testDir;
        private readonly string _templatesDir;

        public AppointmentNoteEngineTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "AppointmentNoteEngine", Guid.NewGuid().ToString("N"));
            _templatesDir = Path.Combine(_testDir, "Vault", "Templates");
            Directory.CreateDirectory(_templatesDir);
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
                TemplatesFolder = "Templates",
                IncludeDailyNoteLink = false,
                CreateMeetingNotes = true,
                GroupRecurringMeetings = true,
                AppointmentDefaultNoteTags = new List<string> { "Meeting" }
            };
        }

        private static AppointmentNoteEngine NewEngine(ObsidianSettings settings, Func<string, bool> contactNoteExists = null)
        {
            TemplateService templates = new TemplateService(new FileService(settings));
            return new AppointmentNoteEngine(settings, templates, contactNoteExists);
        }

        private static AppointmentData SampleAppointment()
        {
            return new AppointmentData
            {
                Subject = "Budget sync",
                Body = "Agenda below.\nJoin: https://teams.microsoft.com/l/meetup-join/abc123\n",
                Location = "Room 4",
                Start = new DateTime(2026, 10, 5, 9, 30, 0),
                End = new DateTime(2026, 10, 5, 10, 15, 0),
                OrganizerName = "Sara Adams",
                OrganizerEmail = "sara@example.com",
                GlobalAppointmentId = "GID-1",
                Recurrence = AppointmentRecurrence.NotRecurring,
                CalendarName = "Calendar (caleb@example.com)",
                RequiredAttendees = new List<string> { "Caleb Bennett", "Dana Cole" },
                OptionalAttendees = new List<string> { "Eli Ford" },
                ResourceAttendees = new List<string> { "Room 4" }
            };
        }

        private static string Lf(string text)
        {
            return text.Replace("\r\n", "\n");
        }

        // --- Title ---

        [Fact]
        public void BuildNoteTitle_DefaultFormat_DateThenSubject()
        {
            AppointmentNoteEngine engine = NewEngine(NewSettings());

            Assert.Equal("2026-10-05 - Budget sync", engine.BuildNoteTitle("Budget sync", "SaraA", new DateTime(2026, 10, 5, 9, 30, 0)));
        }

        [Fact]
        public void BuildNoteTitle_SenderTokenAndMaxLength()
        {
            ObsidianSettings settings = NewSettings();
            settings.AppointmentNoteTitleFormat = "{Sender} - {Subject} - {Date}";
            settings.AppointmentNoteTitleMaxLength = 20;
            AppointmentNoteEngine engine = NewEngine(settings);

            string title = engine.BuildNoteTitle("Quarterly planning review", "SaraA", new DateTime(2026, 10, 5));

            Assert.StartsWith("SaraA - ", title);
            Assert.True(title.Length <= 20, title);
        }

        // --- Recurrence ---

        [Theory]
        [InlineData(AppointmentRecurrence.NotRecurring, "None")]
        [InlineData(AppointmentRecurrence.Master, "Series")]
        [InlineData(AppointmentRecurrence.Occurrence, "Occurrence")]
        [InlineData(AppointmentRecurrence.Exception, "Exception")]
        public void DescribeRecurrence_ReturnsReadableValue(AppointmentRecurrence recurrence, string expected)
        {
            Assert.Equal(expected, AppointmentNoteEngine.DescribeRecurrence(recurrence));
        }

        [Theory]
        [InlineData(AppointmentRecurrence.NotRecurring, true, false)]
        [InlineData(AppointmentRecurrence.Master, true, false)]
        [InlineData(AppointmentRecurrence.Occurrence, true, true)]
        [InlineData(AppointmentRecurrence.Exception, true, true)]
        [InlineData(AppointmentRecurrence.Occurrence, false, false)]
        public void ShouldGroupAsRecurring_OnlyOccurrencesAndExceptionsWhenGroupingIsOn(AppointmentRecurrence recurrence, bool group, bool expected)
        {
            ObsidianSettings settings = NewSettings();
            settings.GroupRecurringMeetings = group;

            Assert.Equal(expected, NewEngine(settings).ShouldGroupAsRecurring(recurrence));
        }

        // --- Front matter ---

        [Fact]
        public void BuildMetadata_KeysInOrderWithLinkedPeople()
        {
            ObsidianSettings settings = NewSettings();
            settings.IncludeDailyNoteLink = true;
            AppointmentNoteEngine engine = NewEngine(settings);

            Dictionary<string, object> metadata = engine.BuildMetadata("2026-10-05 - Budget sync", SampleAppointment(), "2026-10-05 - Budget sync");

            Assert.Equal(
                new[]
                {
                    "title", "type", "calendar", "organizer", "organizerEmail", "location", "startDateTime",
                    "endDateTime", "recurrence", "globalAppointmentId", "meetingNotes", "attendees",
                    "optionalAttendees", "resources", "dailyNoteLink", "tags", "meetingUrl"
                },
                metadata.Keys.ToArray());
            Assert.Equal("Appointment", metadata["type"]);
            Assert.Equal("Calendar (caleb@example.com)", metadata["calendar"]);
            Assert.Equal("[[Sara Adams]]", metadata["organizer"]);
            Assert.Equal("sara@example.com", metadata["organizerEmail"]);
            Assert.Equal("2026-10-05T09:30", metadata["startDateTime"]);
            Assert.Equal("2026-10-05T10:15", metadata["endDateTime"]);
            Assert.Equal("None", metadata["recurrence"]);
            Assert.Equal("GID-1", metadata["globalAppointmentId"]);
            Assert.Equal("[[2026-10-05 - Budget sync - Meeting Notes]]", metadata["meetingNotes"]);
            Assert.Equal(new List<string> { "[[Caleb Bennett]]", "[[Dana Cole]]" }, metadata["attendees"]);
            Assert.Equal(new List<string> { "[[Eli Ford]]" }, metadata["optionalAttendees"]);
            Assert.Equal(new List<string> { "[[Room 4]]" }, metadata["resources"]);
            Assert.Equal("[[2026-10-05]]", metadata["dailyNoteLink"]);
            Assert.Equal(new List<string> { "Meeting" }, metadata["tags"]);
            Assert.Equal("https://teams.microsoft.com/l/meetup-join/abc123", metadata["meetingUrl"]);
        }

        [Fact]
        public void BuildMetadata_NoPeopleNoMeetingNotesNoLink_LeavesThoseKeysOut()
        {
            ObsidianSettings settings = NewSettings();
            settings.CreateMeetingNotes = false;
            settings.AppointmentDefaultNoteTags = new List<string>();
            AppointmentData appointment = new AppointmentData
            {
                Subject = "Focus time",
                Start = new DateTime(2026, 10, 5, 13, 0, 0),
                End = new DateTime(2026, 10, 5, 14, 0, 0),
                Recurrence = AppointmentRecurrence.Exception
            };

            Dictionary<string, object> metadata = NewEngine(settings).BuildMetadata("2026-10-05 - Focus time", appointment, "2026-10-05 - Focus time");

            Assert.Equal(
                new[]
                {
                    "title", "type", "calendar", "organizer", "organizerEmail", "location", "startDateTime",
                    "endDateTime", "recurrence", "globalAppointmentId"
                },
                metadata.Keys.ToArray());
            Assert.Equal(string.Empty, metadata["organizer"]);
            Assert.Equal("Exception", metadata["recurrence"]);
        }

        [Fact]
        public void BuildMetadata_ContactLinkFormats_UseTheFirstFormatWhoseNoteExists()
        {
            ObsidianSettings settings = NewSettings();
            settings.ContactLinkFormats = new List<string> { "[[{FullName}]]", "[[{LastName}, {FirstName}]]" };
            List<string> asked = new List<string>();
            AppointmentNoteEngine engine = NewEngine(settings, stem =>
            {
                asked.Add(stem);
                return stem == "Adams, Sara";
            });

            Dictionary<string, object> metadata = engine.BuildMetadata("t", SampleAppointment(), "t");

            Assert.Equal("[[Adams, Sara]]", metadata["organizer"]);
            Assert.Contains("Sara Adams", asked);
            // Nobody else has a note: they fall back to the first format.
            Assert.Equal(new List<string> { "[[Caleb Bennett]]", "[[Dana Cole]]" }, metadata["attendees"]);
        }

        [Fact]
        public void BuildMetadata_MalformedDailyNoteFormat_FallsBackToIsoDate()
        {
            ObsidianSettings settings = NewSettings();
            settings.IncludeDailyNoteLink = true;
            settings.DailyNoteLinkFormat = "[[%]]";

            Dictionary<string, object> metadata = NewEngine(settings).BuildMetadata("t", SampleAppointment(), "t");

            Assert.Equal("[[2026-10-05]]", metadata["dailyNoteLink"]);
        }

        // --- Note text ---

        [Fact]
        public void BuildNoteContent_BuiltInLayout_ExactText()
        {
            ObsidianSettings settings = NewSettings();
            AppointmentNoteEngine engine = NewEngine(settings);
            AppointmentData appointment = SampleAppointment();
            Dictionary<string, object> metadata = engine.BuildMetadata("2026-10-05 - Budget sync", appointment, "2026-10-05 - Budget sync");

            string content = engine.BuildNoteContent(metadata, "2026-10-05 - Budget sync", "Budget sync", appointment, "2026-10-05", "2026-10-05 - Budget sync", "- [ ] Follow up");

            string expected = ""
                + "---\n"
                + "title: \"2026-10-05 - Budget sync\"\n"
                + "type: \"Appointment\"\n"
                + "calendar: \"Calendar (caleb@example.com)\"\n"
                + "organizer: \"[[Sara Adams]]\"\n"
                + "organizerEmail: \"sara@example.com\"\n"
                + "location: \"Room 4\"\n"
                + "startDateTime: \"2026-10-05T09:30\"\n"
                + "endDateTime: \"2026-10-05T10:15\"\n"
                + "recurrence: \"None\"\n"
                + "globalAppointmentId: \"GID-1\"\n"
                + "meetingNotes: \"[[2026-10-05 - Budget sync - Meeting Notes]]\"\n"
                + "attendees: \n"
                + "  - \"[[Caleb Bennett]]\"\n"
                + "  - \"[[Dana Cole]]\"\n"
                + "optionalAttendees: \n"
                + "  - \"[[Eli Ford]]\"\n"
                + "resources: \n"
                + "  - \"[[Room 4]]\"\n"
                + "tags: \n"
                + "  - \"Meeting\"\n"
                + "meetingUrl: \"https://teams.microsoft.com/l/meetup-join/abc123\"\n"
                + "---\n"
                + "\n"
                + "# 2026-10-05 - Budget sync\n"
                + "\n"
                + "- [ ] Follow up\n"
                + "\n"
                + "**Location:** Room 4\n"
                + "\n"
                + "**Join:** [Join the meeting](https://teams.microsoft.com/l/meetup-join/abc123)\n"
                + "\n"
                + "**Start:** 2026-10-05T09:30\n"
                + "**End:** 2026-10-05T10:15\n"
                + "\n"
                + "**Organizer:** [[Sara Adams]]\n"
                + "\n"
                + "**Calendar:** Calendar (caleb@example.com)\n"
                + "\n"
                + "## Description\n"
                + "\n"
                + "Agenda below.\n"
                + "Join: https://teams.microsoft.com/l/meetup-join/abc123\n"
                + "\n"
                + "## Meeting Notes\n"
                + "\n";
            Assert.Equal(expected, Lf(content));
        }

        [Fact]
        public void BuildNoteContent_BuiltInLayout_EmptyPartsAndNoMeetingNotes_AreLeftOut()
        {
            ObsidianSettings settings = NewSettings();
            settings.CreateMeetingNotes = false;
            settings.AppointmentDefaultNoteTags = new List<string>();
            AppointmentNoteEngine engine = NewEngine(settings);
            AppointmentData appointment = new AppointmentData
            {
                Subject = "Focus time",
                Start = new DateTime(2026, 10, 5, 13, 0, 0),
                End = new DateTime(2026, 10, 5, 14, 0, 0)
            };
            Dictionary<string, object> metadata = engine.BuildMetadata("Focus", appointment, "Focus");

            string content = Lf(engine.BuildNoteContent(metadata, "Focus", "Focus time", appointment, "2026-10-05", "Focus"));

            string afterFrontMatter = content.Substring(content.IndexOf("---\n\n", StringComparison.Ordinal) + 5);
            Assert.Equal("# Focus\n\n**Start:** 2026-10-05T13:00\n**End:** 2026-10-05T14:00\n\n", afterFrontMatter);
        }

        [Fact]
        public void BuildNoteContent_NoTemplateFile_UsesBuiltInLayout()
        {
            // The default AppointmentTemplateFile ("AppointmentTemplate.md") is not in the templates
            // folder, so the built-in layout is used.
            AppointmentNoteEngine engine = NewEngine(NewSettings());

            Assert.False(engine.HasCustomTemplate());
        }

        [Fact]
        public void BuildNoteContent_CustomTemplate_FillsEveryToken()
        {
            File.WriteAllText(Path.Combine(_templatesDir, "AppointmentTemplate.md"),
                "{{noteTitle}}|{{subject}}|{{organizer}}|{{organizerEmail}}|{{attendees}}|{{optionalAttendees}}|{{resources}}|{{location}}|{{startDateTime}}|{{endDateTime}}|{{recurrence}}|{{date}}|{{body}}|{{taskBlock}}|{{fileName}}|{{fileNameNoExt}}");
            ObsidianSettings settings = NewSettings();
            AppointmentNoteEngine engine = NewEngine(settings);
            AppointmentData appointment = SampleAppointment();
            appointment.Recurrence = AppointmentRecurrence.Occurrence;
            Dictionary<string, object> metadata = engine.BuildMetadata("T", appointment, "F");

            Assert.True(engine.HasCustomTemplate());
            string content = engine.BuildNoteContent(metadata, "T", "Budget sync", appointment, "2026-10-05", "F", "TASK");

            Assert.Equal(
                "T|Budget sync|[[Sara Adams]]|sara@example.com|[[Caleb Bennett]], [[Dana Cole]]|[[Eli Ford]]|[[Room 4]]|Room 4|2026-10-05T09:30|2026-10-05T10:15|Occurrence|2026-10-05|Agenda below.\nJoin: https://teams.microsoft.com/l/meetup-join/abc123|TASK|F.md|F",
                Lf(content));
        }
        // --- Companion meeting note and recurring summary ---

        [Fact]
        public void MeetingNoteFileName_AppendsMeetingNotes()
        {
            Assert.Equal("2026-10-05 - Budget sync - Meeting Notes.md", AppointmentNoteEngine.MeetingNoteFileName("2026-10-05 - Budget sync"));
        }

        [Fact]
        public void BuildMeetingNoteContent_DefaultTemplate_ExactText()
        {
            AppointmentNoteEngine engine = NewEngine(NewSettings());

            string content = engine.BuildMeetingNoteContent(
                "2026-10-05 - Budget sync", "2026-10-05 - Budget sync", "Sara Adams", "Caleb Bennett, Dana Cole", "2026-10-05", "Room 4");

            Assert.Equal(""
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
                + "**Organizer:** Sara Adams\n"
                + "**Attendees:** Caleb Bennett, Dana Cole\n"
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
                + "- [ ]",
                Lf(content));
        }

        [Fact]
        public void BuildMeetingNoteContent_NoTags_LeavesTagsOut()
        {
            ObsidianSettings settings = NewSettings();
            settings.AppointmentDefaultNoteTags = new List<string>();

            string content = NewEngine(settings).BuildMeetingNoteContent("A", "A", null, null, "2026-10-05", null);

            Assert.DoesNotContain("tags:", content);
            Assert.Contains("**Date:** 2026-10-05", content);
        }

        [Fact]
        public void BuildMetadata_RecurringThreadId_IsTheLastKey()
        {
            // The series summary's Dataview query matches instances on threadId.
            Dictionary<string, object> metadata = NewEngine(NewSettings()).BuildMetadata("t", SampleAppointment(), "t", "Weekly standup");

            Assert.Equal("threadId", metadata.Keys.Last());
            Assert.Equal("Weekly standup", metadata["threadId"]);
        }

        [Fact]
        public void BuildMetadata_NotRecurring_HasNoThreadId()
        {
            Assert.False(NewEngine(NewSettings()).BuildMetadata("t", SampleAppointment(), "t").ContainsKey("threadId"));
        }

        [Fact]
        public void BuildRecurringSummaryContent_AbsoluteFolder_QueriesTheVaultRelativePath()
        {
            ObsidianSettings settings = NewSettings();
            string folder = Path.Combine(settings.GetFullVaultPath(), "Appointments", "Weekly standup");

            string content = NewEngine(settings).BuildRecurringSummaryContent(folder, "Weekly standup", "Weekly standup");

            Assert.Contains("dv.pages('\"Appointments/Weekly standup\"')", content);
        }

        [Fact]
        public void BuildRecurringSummaryContent_UsesTheThreadTemplate()
        {
            AppointmentNoteEngine engine = NewEngine(NewSettings());

            string content = Lf(engine.BuildRecurringSummaryContent("Appointments/Weekly standup", "Weekly standup", "Weekly standup"));

            Assert.StartsWith("---\ntitle: \"Weekly standup\"\ntype: email-thread\nthreadId: \"Weekly standup\"\n", content);
            Assert.Contains("const threadId = \"Weekly standup\";", content);
        }
    }
}
