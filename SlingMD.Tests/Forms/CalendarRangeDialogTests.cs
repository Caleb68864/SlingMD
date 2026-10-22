using System;
using System.Collections.Generic;
using System.Threading;
using Xunit;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;

namespace SlingMD.Tests.Forms
{
    public class CalendarRangeDialogTests
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

        private static List<CalendarFolderInfo> ThreeCalendars()
        {
            return new List<CalendarFolderInfo>
            {
                new CalendarFolderInfo { EntryId = "A", StoreId = "S1", DisplayName = "Calendar (jane@firm.com)", IsAccountDefault = true },
                new CalendarFolderInfo { EntryId = "B", StoreId = "S1", DisplayName = "Court Dates (jane@firm.com)" },
                new CalendarFolderInfo { EntryId = "C", StoreId = "S2", DisplayName = "Docket (Shared Calendars)" }
            };
        }

        [Fact]
        public void CalendarRangeDialog_DefaultDates_AreToday()
        {
            using (CalendarRangeDialog dialog = new CalendarRangeDialog())
            {
                Assert.Equal(DateTime.Today, dialog.StartDate);
                Assert.Equal(DateTime.Today, dialog.EndDate);
            }
        }

        [Fact]
        public void CalendarRangeDialog_ValidRange_EnablesOkButton()
        {
            using (CalendarRangeDialog dialog = new CalendarRangeDialog())
            {
                dialog.dtpStart.Value = DateTime.Today;
                dialog.dtpEnd.Value = DateTime.Today.AddDays(3);

                Assert.True(dialog.IsOkEnabled);
            }
        }

        [Fact]
        public void CalendarRangeDialog_StartDateAfterEndDate_DisablesOkButton()
        {
            using (CalendarRangeDialog dialog = new CalendarRangeDialog())
            {
                dialog.dtpEnd.Value = DateTime.Today;
                dialog.dtpStart.Value = DateTime.Today.AddDays(2);

                Assert.False(dialog.IsOkEnabled);
            }
        }

        [Fact]
        public void CalendarRangeDialog_ListsCalendarsAndChecksPreselected()
        {
            RunSta(() =>
            {
                List<CalendarFolderInfo> calendars = ThreeCalendars();
                List<CalendarFolderInfo> preselected = new List<CalendarFolderInfo>
                {
                    new CalendarFolderInfo { EntryId = "a" },
                    new CalendarFolderInfo { EntryId = "C" }
                };

                using (CalendarRangeDialog dialog = new CalendarRangeDialog(calendars, preselected))
                {
                    Assert.Equal(3, dialog.clbCalendars.Items.Count);
                    Assert.Equal("Court Dates (jane@firm.com)", dialog.clbCalendars.Items[1].ToString());
                    Assert.Equal(new[] { "A", "C" }, dialog.SelectedCalendars.ConvertAll(c => c.EntryId));
                    Assert.True(dialog.IsOkEnabled);
                }
            });
        }

        [Fact]
        public void CalendarRangeDialog_NoCalendarChecked_DisablesOkUntilOneIsChecked()
        {
            RunSta(() =>
            {
                List<CalendarFolderInfo> calendars = ThreeCalendars();
                using (CalendarRangeDialog dialog = new CalendarRangeDialog(calendars, new[] { calendars[0] }))
                {
                    dialog.clbCalendars.SetItemChecked(0, false);
                    Assert.False(dialog.IsOkEnabled);
                    Assert.Empty(dialog.SelectedCalendars);

                    dialog.clbCalendars.SetItemChecked(2, true);
                    Assert.True(dialog.IsOkEnabled);
                    Assert.Equal("C", Assert.Single(dialog.SelectedCalendars).EntryId);
                }
            });
        }

        [Fact]
        public void CalendarRangeDialog_InvalidDates_StayDisabledEvenWithCalendarChecked()
        {
            RunSta(() =>
            {
                List<CalendarFolderInfo> calendars = ThreeCalendars();
                using (CalendarRangeDialog dialog = new CalendarRangeDialog(calendars, calendars))
                {
                    dialog.dtpEnd.Value = DateTime.Today;
                    dialog.dtpStart.Value = DateTime.Today.AddDays(1);
                    Assert.False(dialog.IsOkEnabled);

                    dialog.clbCalendars.SetItemChecked(1, false);
                    Assert.False(dialog.IsOkEnabled);
                }
            });
        }

        [Fact]
        public void CalendarRangeDialog_NoCalendarsOffered_KeepsOkEnabledAndDisablesList()
        {
            // When Outlook yields no calendars the export falls back to the saved or default
            // calendars, so the dialog must not block on an empty, uncheckable list.
            RunSta(() =>
            {
                using (CalendarRangeDialog dialog = new CalendarRangeDialog(new List<CalendarFolderInfo>(), null))
                {
                    Assert.True(dialog.IsOkEnabled);
                    Assert.False(dialog.clbCalendars.Enabled);
                    Assert.Empty(dialog.SelectedCalendars);
                }
            });
        }
    }
}
