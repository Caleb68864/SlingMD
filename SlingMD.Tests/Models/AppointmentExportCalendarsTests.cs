using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using Xunit;

namespace SlingMD.Tests.Models
{
    /// <summary>
    /// Covers <see cref="ObsidianSettings.AppointmentExportCalendars"/>, the calendar choice made in
    /// the Save Date Range dialog and reused by Save Today.
    /// </summary>
    public class AppointmentExportCalendarsTests : IDisposable
    {
        private readonly string _testSettingsDir;
        private readonly string _testSettingsPath;

        public AppointmentExportCalendarsTests()
        {
            _testSettingsDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ExportCalendars_" + Guid.NewGuid().ToString("N"));
            _testSettingsPath = Path.Combine(_testSettingsDir, "ObsidianSettings.json");
            Directory.CreateDirectory(_testSettingsDir);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testSettingsDir))
            {
                Directory.Delete(_testSettingsDir, true);
            }
        }

        [Fact]
        public void Default_IsEmpty_MeaningAccountDefaultCalendars()
        {
            ObsidianSettings settings = new ObsidianSettings();

            Assert.NotNull(settings.AppointmentExportCalendars);
            Assert.Empty(settings.AppointmentExportCalendars);
        }

        [Fact]
        public void SaveAndLoad_PersistsCalendarIdentityButNotDefaultFlag()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                AppointmentExportCalendars = new List<CalendarFolderInfo>
                {
                    new CalendarFolderInfo { EntryId = "00AA", StoreId = "S1", DisplayName = "Calendar (jane@firm.com)", IsAccountDefault = true },
                    new CalendarFolderInfo { EntryId = "00BB", StoreId = "S2", DisplayName = "Docket (Shared Calendars)" }
                }
            };
            settings.Save();

            ObsidianSettingsTestable reloaded = new ObsidianSettingsTestable { TestSettingsPath = _testSettingsPath };
            reloaded.Load();

            Assert.Equal(2, reloaded.AppointmentExportCalendars.Count);
            Assert.Equal("00AA", reloaded.AppointmentExportCalendars[0].EntryId);
            Assert.Equal("S1", reloaded.AppointmentExportCalendars[0].StoreId);
            Assert.Equal("Calendar (jane@firm.com)", reloaded.AppointmentExportCalendars[0].DisplayName);
            Assert.False(reloaded.AppointmentExportCalendars[0].IsAccountDefault);
            Assert.Equal("Docket (Shared Calendars)", reloaded.AppointmentExportCalendars[1].DisplayName);
        }

        [Fact]
        public void Load_DropsBlankAndDuplicateEntries()
        {
            // A hand-edited file must never make one calendar export twice, and an entry without an
            // EntryID cannot be reopened at all. EntryIDs are compared without regard to case.
            File.WriteAllText(_testSettingsPath, @"{
  ""AppointmentExportCalendars"": [
    { ""EntryId"": ""00aa"", ""StoreId"": ""S1"", ""DisplayName"": ""Calendar"" },
    { ""EntryId"": """", ""StoreId"": ""S1"", ""DisplayName"": ""Blank"" },
    null,
    { ""EntryId"": ""00AA"", ""StoreId"": ""S1"", ""DisplayName"": ""Calendar again"" },
    { ""EntryId"": ""00BB"" }
  ]
}");

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable { TestSettingsPath = _testSettingsPath };
            settings.Load();

            Assert.Equal(2, settings.AppointmentExportCalendars.Count);
            Assert.Equal("Calendar", settings.AppointmentExportCalendars[0].DisplayName);
            Assert.Equal("00BB", settings.AppointmentExportCalendars[1].EntryId);
            Assert.Equal(string.Empty, settings.AppointmentExportCalendars[1].StoreId);
            Assert.Equal(string.Empty, settings.AppointmentExportCalendars[1].DisplayName);
        }

        [Fact]
        public void Load_NullList_BecomesEmpty()
        {
            File.WriteAllText(_testSettingsPath, @"{ ""AppointmentExportCalendars"": null }");

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable { TestSettingsPath = _testSettingsPath };
            settings.Load();

            Assert.NotNull(settings.AppointmentExportCalendars);
            Assert.Empty(settings.AppointmentExportCalendars);
        }
    }
}
