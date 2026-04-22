using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Tests.Infrastructure;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Tests TaskService date math against an injected fake clock — the wire-up that
    /// makes "given today is Friday, due is X" testable without DateTime.Now.
    /// </summary>
    public class TaskServiceClockTests
    {
        [Fact]
        public void GenerateObsidianTask_DueDate_IsFakeNowPlusDueDays()
        {
            ObsidianSettings settings = new ObsidianSettings();
            FakeClock clock = new FakeClock(new System.DateTime(2026, 4, 21, 9, 0, 0));
            TaskService svc = new TaskService(settings, null, clock);
            svc.InitializeTaskSettings(dueDays: 3, reminderDays: 1, reminderHour: 9, useRelativeReminder: false);

            string line = svc.GenerateObsidianTask("Note", new List<string> { "FollowUp" });

            // Plugin date strings are always yyyy-MM-dd regardless of EmailDateFormat.
            Assert.Contains("📅 2026-04-24", line);
        }

        [Fact]
        public void GenerateObsidianTask_RelativeReminder_BeforeDueDate()
        {
            ObsidianSettings settings = new ObsidianSettings();
            FakeClock clock = new FakeClock(new System.DateTime(2026, 4, 21, 9, 0, 0));
            TaskService svc = new TaskService(settings, null, clock);
            svc.InitializeTaskSettings(dueDays: 7, reminderDays: 2, reminderHour: 9, useRelativeReminder: true);

            string line = svc.GenerateObsidianTask("Note", new List<string> { "FollowUp" });

            // Due 7 days from "now"; relative reminder is 2 days before that.
            Assert.Contains("📅 2026-04-28", line);
            Assert.Contains("🛫 2026-04-26", line);
        }

        [Fact]
        public void GenerateObsidianTask_AbsoluteReminder_FromNow()
        {
            ObsidianSettings settings = new ObsidianSettings();
            FakeClock clock = new FakeClock(new System.DateTime(2026, 4, 21, 9, 0, 0));
            TaskService svc = new TaskService(settings, null, clock);
            svc.InitializeTaskSettings(dueDays: 7, reminderDays: 2, reminderHour: 9, useRelativeReminder: false);

            string line = svc.GenerateObsidianTask("Note", new List<string> { "FollowUp" });

            // Absolute reminder is 2 days from "now", independent of due date.
            Assert.Contains("🛫 2026-04-23", line);
        }

        [Fact]
        public void GenerateObsidianTask_CreatedDate_IsFakeNow()
        {
            ObsidianSettings settings = new ObsidianSettings();
            FakeClock clock = new FakeClock(new System.DateTime(2026, 4, 21, 14, 30, 0));
            TaskService svc = new TaskService(settings, null, clock);
            svc.InitializeTaskSettings();

            string line = svc.GenerateObsidianTask("Note", new List<string> { "FollowUp" });

            Assert.Contains("➕ 2026-04-21", line);
        }

        [Fact]
        public void GenerateObsidianTask_DefaultClock_IsSystemClock()
        {
            // Sanity check: TaskService still works without an injected clock.
            ObsidianSettings settings = new ObsidianSettings();
            TaskService svc = new TaskService(settings);
            svc.InitializeTaskSettings(1, 1, 9, false);

            string line = svc.GenerateObsidianTask("Note", new List<string> { "FollowUp" });

            Assert.Contains("📅 ", line);
            Assert.Contains("🛫 ", line);
        }
    }
}
