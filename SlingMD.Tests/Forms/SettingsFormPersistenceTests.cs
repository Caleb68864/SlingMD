using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;
using SlingMD.Tests.Models;
using Xunit;

namespace SlingMD.Tests.Forms
{
    /// <summary>
    /// End-to-end reproduction for issue #13: settings on the Appointment, Contacts, Tasks,
    /// Threading, and Attachments tabs revert to defaults on restart. Drives the real
    /// <see cref="SettingsForm"/> through the same LoadSettings -> controls -> Save handler ->
    /// disk round-trip the user exercises, then reloads from a fresh instance ("restart").
    /// </summary>
    public class SettingsFormPersistenceTests
    {
        /// <summary>
        /// Runs the given action on a dedicated STA thread (WinForms controls require STA).
        /// Re-throws any exception raised on that thread so the test fails normally.
        /// </summary>
        private static void RunSta(Action action)
        {
            System.Exception thrown = null;
            Thread staThread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (System.Exception ex)
                {
                    thrown = ex;
                }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            if (thrown != null)
            {
                throw new System.Exception("STA thread failed: " + thrown, thrown);
            }
        }

        private static void InvokeSaveHandler(SettingsForm form)
        {
            MethodInfo handler = typeof(SettingsForm).GetMethod(
                "btnSave_Click",
                BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(handler);
            handler.Invoke(form, new object[] { form, EventArgs.Empty });
        }

        private static void SetControlText(SettingsForm form, string fieldName, string text)
        {
            FieldInfo field = typeof(SettingsForm).GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(field);
            Control control = (Control)field.GetValue(form);
            control.Text = text;
        }

        [Fact]
        public void FiveTabs_NonDefaultValues_SurviveSaveAndReload()
        {
            RunSta(() =>
            {
                string dir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "FormPersist_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir);
                string settingsPath = Path.Combine(dir, "ObsidianSettings.json");

                // Non-default values across exactly the five tabs the reporter named.
                ObsidianSettingsTestable seed = new ObsidianSettingsTestable
                {
                    TestSettingsPath = settingsPath,

                    // Appointments tab
                    AppointmentsFolder = "MyAppointments",
                    AppointmentNoteTitleMaxLength = 123,
                    AppointmentSaveAttachments = false,      // default true -> flip
                    CreateMeetingNotes = false,              // default true -> flip
                    GroupRecurringMeetings = false,          // default true -> flip
                    SaveCancelledAppointments = true,        // default false -> flip
                    AppointmentTaskCreation = "Both",
                    MeetingNoteTemplateFile = "MyMeetingTemplate.md",

                    // Contacts tab
                    ContactsFolder = "MyContacts",
                    EnableContactSaving = false,             // default true -> flip
                    SearchEntireVaultForContacts = true,
                    EnableContactFuzzyMatching = true,
                    AutoSaveAliasOnMatchConfirmed = false,   // default true -> flip
                    ContactNoteIncludeDetails = false,       // default true -> flip

                    // Tasks tab
                    CreateObsidianTask = false,              // default true -> flip
                    CreateOutlookTask = true,                // default false -> flip
                    AskForDates = true,
                    UseRelativeReminder = true,             // default false -> flip
                    DefaultDueDays = 7,
                    DefaultReminderDays = 3,
                    DefaultReminderHour = 14,

                    // Threading tab
                    GroupEmailThreads = false,               // default true -> flip
                    MoveDateToFrontInThread = false,         // default true -> flip

                    // Attachments tab
                    AttachmentsFolder = "MyAttachments",
                    AttachmentStorageMode = AttachmentStorageMode.Centralized,
                    SaveRealAttachments = false,             // default true -> flip
                    SaveInlineImages = true,                 // default false -> flip
                    SaveAllAttachments = true,               // default false -> flip
                    UseObsidianWikilinks = false             // default true -> flip
                };

                // Constructing the form runs LoadSettings(), populating controls from `seed`.
                using (SettingsForm form = new SettingsForm(seed))
                {
                    // The user clicks Save without touching anything: controls -> settings -> disk.
                    InvokeSaveHandler(form);
                }

                // Simulate an Outlook restart: brand-new settings object loads from disk.
                ObsidianSettingsTestable reloaded = new ObsidianSettingsTestable
                {
                    TestSettingsPath = settingsPath
                };
                reloaded.Load();

                // Appointments
                Assert.Equal("MyAppointments", reloaded.AppointmentsFolder);
                Assert.Equal(123, reloaded.AppointmentNoteTitleMaxLength);
                Assert.False(reloaded.AppointmentSaveAttachments);
                Assert.False(reloaded.CreateMeetingNotes);
                Assert.False(reloaded.GroupRecurringMeetings);
                Assert.True(reloaded.SaveCancelledAppointments);
                Assert.Equal("Both", reloaded.AppointmentTaskCreation);
                Assert.Equal("MyMeetingTemplate.md", reloaded.MeetingNoteTemplateFile);

                // Contacts
                Assert.Equal("MyContacts", reloaded.ContactsFolder);
                Assert.False(reloaded.EnableContactSaving);
                Assert.True(reloaded.SearchEntireVaultForContacts);
                Assert.True(reloaded.EnableContactFuzzyMatching);
                Assert.False(reloaded.AutoSaveAliasOnMatchConfirmed);
                Assert.False(reloaded.ContactNoteIncludeDetails);

                // Tasks
                Assert.False(reloaded.CreateObsidianTask);
                Assert.True(reloaded.CreateOutlookTask);
                Assert.True(reloaded.AskForDates);
                Assert.True(reloaded.UseRelativeReminder);
                Assert.Equal(7, reloaded.DefaultDueDays);
                Assert.Equal(3, reloaded.DefaultReminderDays);
                Assert.Equal(14, reloaded.DefaultReminderHour);

                // Threading
                Assert.False(reloaded.GroupEmailThreads);
                Assert.False(reloaded.MoveDateToFrontInThread);

                // Attachments
                Assert.Equal("MyAttachments", reloaded.AttachmentsFolder);
                Assert.Equal(AttachmentStorageMode.Centralized, reloaded.AttachmentStorageMode);
                Assert.False(reloaded.SaveRealAttachments);
                Assert.True(reloaded.SaveInlineImages);
                Assert.True(reloaded.SaveAllAttachments);
                Assert.False(reloaded.UseObsidianWikilinks);
            });
        }

        /// <summary>
        /// Regression guard for issue #13's real failure mode: a single invalid field (here an
        /// Appointments folder containing a path separator, which fails ObsidianSettings.Validate)
        /// must NOT silently discard the save. The dialog must stay open (DialogResult != OK) and
        /// nothing must be written to disk, so the user is told rather than losing every changed tab.
        /// </summary>
        [Fact]
        public void InvalidField_DoesNotCloseDialogAndDoesNotPersist()
        {
            RunSta(() =>
            {
                string dir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "FormInvalid_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir);
                string settingsPath = Path.Combine(dir, "ObsidianSettings.json");

                ObsidianSettingsTestable seed = new ObsidianSettingsTestable
                {
                    TestSettingsPath = settingsPath
                };

                using (TestableSettingsForm form = new TestableSettingsForm(seed))
                {
                    // A backslash is an invalid file-name character -> ObsidianSettings.Validate throws.
                    SetControlText(form, "txtAppointmentsFolder", "bad\\folder");

                    // btnSave_Click swallows the validation error, notifies the user, and returns early.
                    InvokeSaveHandler(form);

                    // The user must be told (not left thinking it saved)...
                    Assert.NotNull(form.LastSaveError);
                    // ...and the dialog must remain open (not accepted) so the change is not silently lost.
                    Assert.NotEqual(DialogResult.OK, form.DialogResult);
                }

                // Nothing should have been written to disk.
                Assert.False(File.Exists(settingsPath), "Invalid settings must not be persisted to disk.");
            });
        }

        /// <summary>
        /// Test double that records the save-failure notification instead of popping a modal
        /// MessageBox (which would block the STA test thread).
        /// </summary>
        private class TestableSettingsForm : SettingsForm
        {
            public string LastSaveError { get; private set; }

            public TestableSettingsForm(ObsidianSettings settings) : base(settings)
            {
            }

            protected override void ShowSaveFailed(string message, System.Exception ex)
            {
                LastSaveError = message;
            }
        }
    }
}
