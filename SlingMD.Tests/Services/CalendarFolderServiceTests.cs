using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Covers the Outlook-independent parts of <see cref="CalendarFolderService"/>: list labels and
    /// which calendars the Save Date Range dialog pre-selects.
    /// </summary>
    public class CalendarFolderServiceTests
    {
        private static CalendarFolderInfo Calendar(string entryId, string name, bool isDefault = false)
        {
            return new CalendarFolderInfo { EntryId = entryId, StoreId = "S", DisplayName = name, IsAccountDefault = isDefault };
        }

        [Theory]
        [InlineData("Calendar", "jane@firm.com", "Calendar (jane@firm.com)")]
        [InlineData("Court Dates", "jane@firm.com", "Court Dates (jane@firm.com)")]
        [InlineData("Calendar - Jane Partner", "Jane Partner", "Calendar - Jane Partner")]
        [InlineData("Docket", "", "Docket")]
        [InlineData("Docket", null, "Docket")]
        [InlineData("", "jane@firm.com", "Calendar (jane@firm.com)")]
        [InlineData("  Docket  ", "  Shared Calendars ", "Docket (Shared Calendars)")]
        public void FormatCalendarLabel_QualifiesNameWithSource(string folderName, string qualifier, string expected)
        {
            Assert.Equal(expected, CalendarFolderService.FormatCalendarLabel(folderName, qualifier));
        }

        [Fact]
        public void GetInitialSelection_NothingSaved_SelectsAccountDefaults()
        {
            List<CalendarFolderInfo> available = new List<CalendarFolderInfo>
            {
                Calendar("A", "Calendar (jane@firm.com)", isDefault: true),
                Calendar("B", "Court Dates (jane@firm.com)"),
                Calendar("C", "Calendar (jane@gmail.com)", isDefault: true)
            };

            List<CalendarFolderInfo> selected = CalendarFolderService.GetInitialSelection(available, new List<CalendarFolderInfo>());

            Assert.Equal(new[] { "A", "C" }, selected.ConvertAll(c => c.EntryId));
        }

        [Fact]
        public void GetInitialSelection_SavedChoice_WinsOverDefaults()
        {
            List<CalendarFolderInfo> available = new List<CalendarFolderInfo>
            {
                Calendar("A", "Calendar (jane@firm.com)", isDefault: true),
                Calendar("B", "Court Dates (jane@firm.com)"),
                Calendar("C", "Docket (Shared Calendars)")
            };
            List<CalendarFolderInfo> saved = new List<CalendarFolderInfo> { Calendar("b", "old label"), Calendar("C", "Docket") };

            List<CalendarFolderInfo> selected = CalendarFolderService.GetInitialSelection(available, saved);

            // Returns the live entries (current labels), matched by EntryID regardless of case.
            Assert.Equal(new[] { "B", "C" }, selected.ConvertAll(c => c.EntryId));
            Assert.Equal("Court Dates (jane@firm.com)", selected[0].DisplayName);
        }

        [Fact]
        public void GetInitialSelection_SavedCalendarsAllGone_FallsBackToDefaults()
        {
            List<CalendarFolderInfo> available = new List<CalendarFolderInfo>
            {
                Calendar("A", "Calendar (jane@firm.com)", isDefault: true),
                Calendar("B", "Court Dates (jane@firm.com)")
            };
            List<CalendarFolderInfo> saved = new List<CalendarFolderInfo> { Calendar("Z", "Removed shared calendar") };

            List<CalendarFolderInfo> selected = CalendarFolderService.GetInitialSelection(available, saved);

            Assert.Single(selected);
            Assert.Equal("A", selected[0].EntryId);
        }

        [Fact]
        public void GetInitialSelection_NullInputs_ReturnEmpty()
        {
            Assert.Empty(CalendarFolderService.GetInitialSelection(null, null));
        }

        [Fact]
        public void IsSameFolder_ComparesEntryIdIgnoringCase()
        {
            Assert.True(Calendar("00ab", "x").IsSameFolder(Calendar("00AB", "y")));
            Assert.False(Calendar("00AB", "x").IsSameFolder(Calendar("00AC", "x")));
            Assert.False(Calendar("", "x").IsSameFolder(Calendar("", "x")));
            Assert.False(Calendar("00AB", "x").IsSameFolder(null));
        }
    }
}
