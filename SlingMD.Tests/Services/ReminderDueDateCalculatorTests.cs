using System;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class ReminderDueDateCalculatorTests
    {
        private readonly ReminderDueDateCalculator _calc = new ReminderDueDateCalculator();

        [Fact]
        public void Calculate_DueDate_IsNowPlusDefaultDueDays()
        {
            DateTime now = new DateTime(2026, 4, 21, 9, 0, 0);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings { DefaultDueDays = 3, DefaultReminderDays = 1, DefaultReminderHour = 9 });
            Assert.Equal(new DateTime(2026, 4, 24), result.DueDate);
        }

        [Fact]
        public void Calculate_RelativeReminder_IsBeforeDueDate()
        {
            DateTime now = new DateTime(2026, 4, 21, 9, 0, 0);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings
            {
                DefaultDueDays = 7,
                UseRelativeReminder = true,
                DefaultReminderDays = 2,
                DefaultReminderHour = 9
            });
            Assert.Equal(new DateTime(2026, 4, 26), result.ReminderDate);
        }

        [Fact]
        public void Calculate_AbsoluteReminder_IsFromNow()
        {
            DateTime now = new DateTime(2026, 4, 21, 9, 0, 0);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings
            {
                DefaultDueDays = 7,
                UseRelativeReminder = false,
                DefaultReminderDays = 2,
                DefaultReminderHour = 9
            });
            Assert.Equal(new DateTime(2026, 4, 23), result.ReminderDate);
        }

        [Fact]
        public void Calculate_FutureReminder_NoPastWarning()
        {
            DateTime now = new DateTime(2026, 4, 21, 9, 0, 0);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings
            {
                DefaultDueDays = 5,
                DefaultReminderDays = 2,
                DefaultReminderHour = 14
            });
            Assert.False(result.PastReminderWarning);
            Assert.True(result.ReminderDateTime > now);
        }

        [Fact]
        public void Calculate_ReminderToday_HourPassed_BumpsToNowPlusOneHour()
        {
            // now = 14:00, reminder lands today (DefaultReminderDays=0) at hour=10 → past
            DateTime now = new DateTime(2026, 4, 21, 14, 0, 0);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings
            {
                DefaultDueDays = 1,
                DefaultReminderDays = 0,
                DefaultReminderHour = 10
            });
            Assert.True(result.PastReminderWarning);
            Assert.Equal(now.AddHours(1), result.ReminderDateTime);
        }

        [Fact]
        public void Calculate_ReminderInPastDate_BumpsToTomorrowAtConfiguredHour()
        {
            // now = 14:00 today, reminder = -2 days from now at hour=10 → in the past, bump to tomorrow @10
            DateTime now = new DateTime(2026, 4, 21, 14, 0, 0);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings
            {
                DefaultDueDays = 1,
                UseRelativeReminder = false,
                DefaultReminderDays = -2,
                DefaultReminderHour = 10
            });
            Assert.True(result.PastReminderWarning);
            Assert.Equal(new DateTime(2026, 4, 22, 10, 0, 0), result.ReminderDateTime);
        }

        [Fact]
        public void Calculate_NullSettings_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => _calc.Calculate(DateTime.Now, null));
        }

        [Fact]
        public void Calculate_DueDateUsesDateOnly_StripsTimeFromNow()
        {
            DateTime now = new DateTime(2026, 4, 21, 23, 59, 59);
            TaskDueDates result = _calc.Calculate(now, new TaskDueSettings { DefaultDueDays = 1 });
            Assert.Equal(new DateTime(2026, 4, 22), result.DueDate);
            Assert.Equal(TimeSpan.Zero, result.DueDate.TimeOfDay);
        }
    }
}
