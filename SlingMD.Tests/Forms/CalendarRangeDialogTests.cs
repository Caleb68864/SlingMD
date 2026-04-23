using System;
using Xunit;
using SlingMD.Outlook.Forms;

namespace SlingMD.Tests.Forms
{
    public class CalendarRangeDialogTests
    {
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
    }
}
